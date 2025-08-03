VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrLibDef 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrLibDef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.PictureBox pbcHighlight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2190
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   6750
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8880
      Top             =   6990
   End
   Begin VB.PictureBox pbcConflict 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8400
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmcImport 
      Caption         =   "&Import"
      Height          =   345
      Left            =   6060
      TabIndex        =   59
      Top             =   6915
      Width           =   1335
   End
   Begin VB.PictureBox imcTrash 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1260
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   58
      Top             =   6660
      Width           =   540
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConflicts 
      Height          =   1035
      Left            =   15
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1140
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
      ScrollBars      =   0
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   15
      Width           =   45
   End
   Begin VB.CommandButton cmcConflict 
      Caption         =   "Con&flict Check"
      Enabled         =   0   'False
      Height          =   345
      Left            =   9540
      TabIndex        =   52
      Top             =   7110
      Visible         =   0   'False
      Width           =   1335
   End
   Begin V10EngineeringDev.CSI_Calendar cccLib 
      Height          =   240
      Left            =   4455
      TabIndex        =   8
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   423
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowDropDownOnFocus=   -1  'True
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
   Begin V10EngineeringDev.CSI_HourPicker hpcEvent 
      Height          =   225
      Left            =   630
      TabIndex        =   38
      Top             =   1785
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   397
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDayPartButtons=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_HourOnColor =   4638790
      CSI_HourOffColor=   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
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
   Begin V10EngineeringDev.CSI_DayPicker dpcEvent 
      Height          =   210
      Left            =   1830
      TabIndex        =   39
      Top             =   1890
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   370
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_DayOnColor  =   4638790
      CSI_DayOffColor =   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
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
   Begin V10EngineeringDev.CSI_TimeLength ltcEvent 
      Height          =   195
      Left            =   975
      TabIndex        =   37
      Top             =   4170
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "00:00:00.0"
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
   Begin V10EngineeringDev.CSI_DayPicker dpcLib 
      Height          =   210
      Left            =   8925
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   370
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_DayOnColor  =   4638790
      CSI_DayOffColor =   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
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
   Begin V10EngineeringDev.CSI_HourPicker hpcLib 
      Height          =   225
      Left            =   7470
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   397
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDayPartButtons=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_HourOnColor =   4638790
      CSI_HourOffColor=   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
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
   Begin VB.ListBox lbcCTE_1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":030A
      Left            =   10080
      List            =   "EngrLibDef.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4095
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmcDefine 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "[New]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5730
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmcNone 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "[None]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5730
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   1170
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
      TabIndex        =   36
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
      Left            =   7095
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcFNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":030E
      Left            =   10140
      List            =   "EngrLibDef.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMTE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0312
      Left            =   8505
      List            =   "EngrLibDef.frx":0314
      Sorted          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcANE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0316
      Left            =   3135
      List            =   "EngrLibDef.frx":0318
      Sorted          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcDefine 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   5670
      ScaleHeight     =   165
      ScaleWidth      =   1035
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ListBox lbcSCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":031A
      Left            =   4695
      List            =   "EngrLibDef.frx":031C
      Sorted          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5205
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcNNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":031E
      Left            =   9030
      List            =   "EngrLibDef.frx":0320
      Sorted          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCTE_2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0322
      Left            =   8715
      List            =   "EngrLibDef.frx":0324
      Sorted          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcASE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0326
      Left            =   2850
      List            =   "EngrLibDef.frx":0328
      Sorted          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcRNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":032A
      Left            =   1275
      List            =   "EngrLibDef.frx":032C
      Sorted          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcETE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":032E
      Left            =   405
      List            =   "EngrLibDef.frx":0330
      Sorted          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTTE_E 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0332
      Left            =   6645
      List            =   "EngrLibDef.frx":0334
      Sorted          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5115
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTTE_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0336
      Left            =   6840
      List            =   "EngrLibDef.frx":0338
      Sorted          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcBuses 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":033A
      Left            =   4650
      List            =   "EngrLibDef.frx":033C
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCCE_A 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":033E
      Left            =   4455
      List            =   "EngrLibDef.frx":0340
      Sorted          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3105
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCCE_B 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0342
      Left            =   2850
      List            =   "EngrLibDef.frx":0344
      Sorted          =   -1  'True
      TabIndex        =   44
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
   Begin VB.ListBox lbcBDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0346
      Left            =   8025
      List            =   "EngrLibDef.frx":0348
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4080
      TabIndex        =   13
      Top             =   1755
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcDropdown 
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
      Left            =   5025
      Picture         =   "EngrLibDef.frx":034A
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcBGE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0444
      Left            =   6990
      List            =   "EngrLibDef.frx":0446
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcDSE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":0448
      Left            =   3645
      List            =   "EngrLibDef.frx":044A
      Sorted          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcDNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrLibDef.frx":044C
      Left            =   2595
      List            =   "EngrLibDef.frx":044E
      Sorted          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   1410
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
      Picture         =   "EngrLibDef.frx":0450
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   3795
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcReplace 
      Caption         =   "&Replace"
      Height          =   345
      Left            =   4425
      TabIndex        =   51
      Top             =   6915
      Width           =   1335
   End
   Begin VB.PictureBox pbcETab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   47
      Top             =   6675
      Width           =   60
   End
   Begin VB.PictureBox pbcESTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   22
      Top             =   1575
      Width           =   60
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4455
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcLib 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7200
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   19
      Top             =   1365
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   360
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
      TabIndex        =   1
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
      Picture         =   "EngrLibDef.frx":054A
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   6870
      TabIndex        =   50
      Top             =   6495
      Width           =   1335
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5235
      TabIndex        =   49
      Top             =   6495
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   345
      Left            =   3585
      TabIndex        =   48
      Top             =   6495
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLibEvents 
      Height          =   4740
      Left            =   165
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1545
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   8361
      _Version        =   393216
      Rows            =   4
      Cols            =   43
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
      _Band(0).Cols   =   43
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLib 
      Height          =   795
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   285
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1402
      _Version        =   393216
      Rows            =   3
      Cols            =   13
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   0
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
   Begin VB.CommandButton cmcSearch 
      Caption         =   "Bus Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10230
      TabIndex        =   54
      Top             =   1200
      Width           =   1125
   End
   Begin VB.TextBox edcSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8535
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1695
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10125
      Top             =   6660
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
      TabIndex        =   55
      Top             =   6300
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.Label lacScreen 
      Caption         =   "Library Definition"
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11640
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   450
      Picture         =   "EngrLibDef.frx":0854
      Top             =   6630
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   10920
      Picture         =   "EngrLibDef.frx":111E
      Top             =   6570
      Width           =   480
   End
End
Attribute VB_Name = "EngrLibDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrLibDef - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private hmSEE As Integer
Private hmSOE As Integer
Private hmCME As Integer
Private hmCTE As Integer
Private hmDHE As Integer
Private hmDEE As Integer

Private imFieldChgd As Integer
Private bmIntegralSet As Boolean
Private smState As String
Private smConflict As String    'Audio; Bus; Both or No
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
Private smReplaceValues() As String
Private smGridValues() As String
Private imStartDateChangable As Integer 'Indicates if the start date can be changed
Private imOverlapCase As Integer    '1=Replace; 2=Terminate; 3=Change Start Date; 4=Split
Private lmCurrentDHE As Long        'DHE code that needs to be altered because of date overlap
Private smOverlapMsg As String
Private imGapExist As Integer    'Gap exist between this library and another library (True or False)
Private lmGapDHE As Long
Private lmGapEndDate As Long
Private lmCharacterWidth As Long
Private imMaxColChars As Integer
Private imIgnoreBDEChg As Integer
Private imLimboAllowed As Integer   'Limbo only allowed for New or currently saved as Limbo
Private imMaxCols As Integer
Private imLibCallType As Integer
Private lmLibCallCode As Long
Private bmInBranch As Boolean
Private bmInSave As Boolean

Private smESCValue As String    'Value used if ESC pressed

Private bmPrinting As Boolean

Private tmDHE As DHE
Private smCurrDEEStamp
Private smDHEComment As String
Private smDHEBusGroups As String
Private smDHEBuses As String
Private tmCurrDEE() As DEE
Private tmCTE As CTE

Private lmDeleteCodes() As Long

Private tmSchdChgInfo As SCHDCHGINFO

Private smCurrBSEStamp As String
Private tmCurrBSE() As BSE
Private smBusGroups() As String
Private smBuses() As String
Private smCurrDBEStamp As String
Private tmCurrDBE() As DBE
Private smCurrEBEStamp As String
Private tmCurrEBE() As EBE
Private smT1Comment() As String
Private tmCurr1CTE_Name() As DEECTE
Private smT2Comment() As String
Private tmCurr2CTE_Name() As DEECTE
Private smEBuses() As String

Private smCurrASEStamp As String
Private tmCurrASE() As ASE

Private smCurrLibDBEStamp As String
Private tmCurrLibDBE() As DBE
Private smCurrLibDEEStamp As String
Private tmCurrLibDEE() As DEE
Private smCurrLibEBEStamp As String
Private tmCurrLibEBE() As EBE

Private tmDBE As DBE
Private tmEBE As EBE

Private fmUsedWidth As Single
Private fmUnusedWidth As Single
Private imUnusedCount As Integer

Private smGridRow(0 To 37) As String

Private tmConflictList() As CONFLICTLIST
Private lmConflictRow As Long

Private tmConflictTest() As CONFLICTTEST

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer
Private lmEEnableRow As Long         'Current or last row focus was on
Private lmEEnableCol As Long         'Current or last column focus was on
Private imDefaultProgIndex As Integer

Private lmHighlightRow As Integer

'Const CATEGORYINDEX = 0
'Const NAMEINDEX = 1
'Const DESCRIPTIONINDEX = 2
'Const AUTOCODEINDEX = 3
'Const STATEINDEX = 4
'Const CODEINDEX = 5
'Const USEDFLAGINDEX = 6
'Const USEDINDEX = 7
'Const MANDATORYINDEX = 8
Const NAMEINDEX = 0
Const SUBLIBNAMEINDEX = 1
Const DESCRIPTIONINDEX = 2
Const STARTDATEINDEX = 3
Const ENDDATEINDEX = 4
Const DAYSINDEX = 5
'Const STARTTIMEINDEX = 6
'Const LENGTHINDEX = 7
Const HOURSINDEX = 6    '8
Const BUSGROUPSINDEX = 7    '9
Const BUSESINDEX = 8    '10
Const CONFLICTINDEX = 9
Const STATEINDEX = 10    '11
Const CODEINDEX = 11    '12
Const USEDFLAGINDEX = 12    '13


Const HIGHLIGHTINDEX = 0
Const EVENTTYPEINDEX = 1
Const BUSNAMEINDEX = 2
Const BUSCTRLINDEX = 3
Const TIMEINDEX = 4
Const AIRHOURSINDEX = 5
Const STARTTYPEINDEX = 6
Const FIXEDINDEX = 7
Const ENDTYPEINDEX = 8
Const DURATIONINDEX = 9
Const AIRDAYSINDEX = 10
Const MATERIALINDEX = 11
Const AUDIONAMEINDEX = 12
Const AUDIOITEMIDINDEX = 13
Const AUDIOISCIINDEX = 14
Const AUDIOCTRLINDEX = 15
Const BACKUPNAMEINDEX = 16  '17
Const BACKUPCTRLINDEX = 17  '18
Const PROTNAMEINDEX = 18    '14
Const PROTITEMIDINDEX = 19  '15
Const PROTISCIINDEX = 20  '15
Const PROTCTRLINDEX = 21    '16
Const RELAY1INDEX = 22
Const RELAY2INDEX = 23
Const FOLLOWINDEX = 24
Const SILENCETIMEINDEX = 25
Const SILENCE1INDEX = 26
Const SILENCE2INDEX = 27
Const SILENCE3INDEX = 28
Const SILENCE4INDEX = 29
Const NETCUE1INDEX = 30
Const NETCUE2INDEX = 31
Const TITLE1INDEX = 32
Const TITLE2INDEX = 33
Const ABCFORMATINDEX = 34
Const ABCPGMCODEINDEX = 35
Const ABCXDSMODEINDEX = 36
Const ABCRECORDITEMINDEX = 37
Const PCODEINDEX = 38
Const SORTTIMEINDEX = 39
Const ERRORFLAGINDEX = 40
Const CHGSTATUSINDEX = 41
Const EVTCONFLICTINDEX = 42

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


Private Sub cccLib_Change()
    Dim ilDay As Integer
    Dim slStr As String
    Dim slCurrDays As String
    Dim slDays As String
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    
    If StrComp(Trim$(grdLib.text), Trim$(cccLib.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        Select Case grdLib.Col
            Case STARTDATEINDEX
                If (Trim$(cccLib.text) <> "") And (grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX) <> "") Then
                    If gIsDate(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                        If gIsDate(cccLib.text) Then
                            If gDateValue(cccLib.text) + 6 > gDateValue(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                                slStr = Trim$(grdLib.TextMatrix(grdLib.Row, DAYSINDEX))
                                slCurrDays = gCreateDayStr(slStr)
                                slDays = String(7, "N")
                                ilStartDay = Weekday(cccLib.text, vbMonday)
                                ilEndDay = Weekday(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX), vbMonday)
                                If ilStartDay <= ilEndDay Then
                                    For ilDay = ilStartDay To ilEndDay Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                Else
                                    For ilDay = ilStartDay To 7 Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                    For ilDay = 1 To ilEndDay Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                End If
                                grdLib.TextMatrix(grdLib.Row, DAYSINDEX) = gDayMap(slDays)
                            End If
                        End If
                    End If
                End If
            Case ENDDATEINDEX
                If (grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "") And (Trim$(cccLib.text) <> "") Then
                    If gIsDate(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) Then
                        If gIsDate(cccLib.text) Then
                            If gDateValue(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) + 6 > gDateValue(cccLib.text) Then
                                slStr = Trim$(grdLib.TextMatrix(grdLib.Row, DAYSINDEX))
                                slCurrDays = gCreateDayStr(slStr)
                                slDays = String(7, "N")
                                ilStartDay = Weekday(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX), vbMonday)
                                ilEndDay = Weekday(cccLib.text, vbMonday)
                                If ilStartDay <= ilEndDay Then
                                    For ilDay = ilStartDay To ilEndDay Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                Else
                                    For ilDay = ilStartDay To 7 Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                    For ilDay = 1 To ilEndDay Step 1
                                        If Mid$(slCurrDays, ilDay, 1) = "Y" Then
                                            Mid$(slDays, ilDay, 1) = "Y"
                                        End If
                                    Next ilDay
                                End If
                                grdLib.TextMatrix(grdLib.Row, DAYSINDEX) = gDayMap(slDays)
                            End If
                        End If
                    End If
                End If
            'Case STARTTIMEINDEX
            '    grdLib.TextMatrix(grdLib.Row, HOURSINDEX) = ""
            'Case LENGTHINDEX
            '    grdLib.TextMatrix(grdLib.Row, HOURSINDEX) = ""
        End Select
        grdLib.text = cccLib.text
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub cmcCancel_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcConflict_Click()
    Dim ilLibRet As Integer
    Dim ilEvtRet As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilColumn As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    gSetMousePointer grdLib, grdLibEvents, vbHourglass
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    grdLibEvents.Redraw = False
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            grdLibEvents.Row = llRow
            For ilColumn = EVENTTYPEINDEX To imMaxCols Step 1
                grdLibEvents.Col = ilColumn
                If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                    grdLibEvents.CellForeColor = vbBlue
                Else
                    grdLibEvents.CellForeColor = vbBlack
                End If
            Next ilColumn
        End If
    Next llRow
    gConflictPop
    mInitConflictTest
    ilLibRet = mCheckLibConflicts()
    ilEvtRet = mCheckEventConflicts()
    lmConflictRow = -1
    grdLibEvents.Redraw = True
    gSetMousePointer grdLib, grdLibEvents, vbDefault
End Sub

Private Sub cmcConflict_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcDefine_Click()
    Dim ilRet As Integer
    ilRet = mBranch()
    cmcDefine.SetFocus
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdLib.Col
        Case NAMEINDEX
            lbcDNE.Visible = Not lbcDNE.Visible
        Case SUBLIBNAMEINDEX
            lbcDSE.Visible = Not lbcDSE.Visible
        Case BUSGROUPSINDEX
            lbcBGE.Visible = Not lbcBGE.Visible
        Case BUSESINDEX
            lbcBDE.Visible = Not lbcBDE.Visible
    End Select
End Sub

Private Sub cmcEDropDown_Click()
    Select Case grdLibEvents.Col
        Case BUSCTRLINDEX
            lbcCCE_B.Visible = Not lbcCCE_B.Visible
        Case EVENTTYPEINDEX
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

Private Sub cmcImport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilAudio As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim llOffsetEventTime As Long
    Dim llEventStartTime As Long
    Dim slHours As String
    Dim ilHour As Integer
    Dim ilOffset As Integer
    Dim ilBus1 As Integer
    Dim ilBus2 As Integer
    Dim ilBusCount As Integer
    Dim slBuses As String
    Dim slAudio As String
    Dim ilRowChgd As Integer
    Dim ilExtract As Integer
    Dim ilMatch As Integer
    Dim ilBusMatch As Integer
    Dim ilIndex As Integer
    Dim ilNext As Integer
    Dim ilDay As Integer
    Dim slDays As String
    
    If bmInSave Then
        Exit Sub
    End If
    ilRet = mMinHeaderFieldsDefined()
    If Not ilRet Then
        Exit Sub
    End If
    sgExtractType = "L"
    sgExtractName = Trim$(grdLib.TextMatrix(grdLib.FixedRows, NAMEINDEX)) & "/" & Trim$(grdLib.TextMatrix(grdLib.FixedRows, SUBLIBNAMEINDEX))
    ReDim sgExtractBusNames(0 To 0) As String
    slStr = grdLib.TextMatrix(grdLib.FixedRows, BUSESINDEX)
    gParseCDFields slStr, False, smBuses()
    For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
        slStr = Trim$(smBuses(ilLoop))
        If slStr <> "" Then
            sgExtractBusNames(UBound(sgExtractBusNames)) = Trim$(slStr)
            ReDim Preserve sgExtractBusNames(0 To UBound(sgExtractBusNames) + 1) As String
        End If
    Next ilLoop
    slStr = grdLib.TextMatrix(grdLib.FixedRows, HOURSINDEX)
    sgExtractHours = gCreateHourStr(slStr)
    ReDim sgExtractAudios(0 To 0) As String
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
            ilFound = False
            For ilAudio = 0 To UBound(sgExtractAudios) - 1 Step 1
                If StrComp(sgExtractAudios(ilAudio), slStr, vbTextCompare) = 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilAudio
            If Not ilFound Then
                sgExtractAudios(UBound(sgExtractAudios)) = slStr
                ReDim Preserve sgExtractAudios(0 To UBound(sgExtractAudios) + 1) As String
            End If
        End If
    Next llRow
    EngrExtract.Show vbModal
    If igReturnCallStatus = CALLDONE Then
        If sgCurrANEStamp = "" Then
            mPopANE
            smCurrASEStamp = ""
            mPopASE
        End If
        mPopBDE
        mPopNNE
        mPopRNE
        mPopCCE_Audio
        mPopCCE_Bus
        mPopCTE
        mPopFNE
        mPopMTE
        mPopSCE
        mPopTTE_EndType
        mPopTTE_StartType
        'Remove and merge
        For llRow = grdLibEvents.Rows - 1 To grdLibEvents.FixedRows Step -1
            If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                llOffsetEventTime = gStrLengthInTenthToLong(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX))
                slHours = gCreateHourStr(slStr)
                slStr = grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX)
                slDays = gCreateDayStr(slStr)
                slBuses = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
                slAudio = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
                gParseCDFields slBuses, False, smBuses()
                ilRowChgd = False
                For ilDay = 1 To 7 Step 1
                    If (Mid$(slDays, ilDay, 1) = "Y") And (Mid$(sgExtractDays, ilDay, 1) = "Y") Then
                        For ilHour = 1 To 24 Step 1
                            If (Mid$(slHours, ilHour, 1) = "Y") And (Mid$(sgExtractHours, ilHour, 1) = "Y") Then
                                llEventStartTime = 36000 * (ilHour - 1) + llOffsetEventTime
                                If (llEventStartTime >= lgExtractStartTime) And (llEventStartTime < lgExtractEndTime) Then
                                    For ilOffset = LBound(lgExtractOffsetStart) To UBound(lgExtractOffsetStart) Step 1
                                        If (llOffsetEventTime >= lgExtractOffsetStart(ilOffset)) And (llOffsetEventTime <= lgExtractOffsetEnd(ilOffset)) Then
                                            For ilBus1 = LBound(smBuses) To UBound(smBuses) Step 1
                                                For ilBus2 = 0 To UBound(sgExtractBusNames) - 1 Step 1
                                                    If StrComp(Trim$(smBuses(ilBus1)), Trim$(sgExtractBusNames(ilBus2)), vbTextCompare) = 0 Then
                                                        If UBound(sgExtractAudios) > LBound(sgExtractAudios) Then
                                                            For ilAudio = 0 To UBound(sgExtractAudios) - 1 Step 1
                                                                If StrComp(slAudio, Trim$(sgExtractAudios(ilAudio)), vbTextCompare) = 0 Then
                                                                    smBuses(ilBus1) = ""
                                                                    Mid$(slHours, ilHour, 1) = "N"
                                                                    Mid$(slDays, ilDay, 1) = "N"
                                                                    ilRowChgd = True
                                                                    Exit For
                                                                End If
                                                            Next ilAudio
                                                        Else
                                                            smBuses(ilBus1) = ""
                                                            Mid$(slHours, ilHour, 1) = "N"
                                                            Mid$(slDays, ilDay, 1) = "N"
                                                            ilRowChgd = True
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilBus2
                                            Next ilBus1
                                            Exit For
                                        End If
                                    Next ilOffset
                                End If
                            End If
                        Next ilHour
                    End If
                Next ilDay
                If ilRowChgd Then
                    imFieldChgd = True
                    ilBusCount = 0
                    slBuses = ""
                    For ilBus1 = LBound(smBuses) To UBound(smBuses) Step 1
                        If Trim(smBuses(ilBus1)) <> "" Then
                            ilBusCount = ilBusCount + 1
                            If slBuses = "" Then
                                slBuses = smBuses(ilBus1)
                            Else
                                slBuses = slBuses & ", " & smBuses(ilBus1)
                            End If
                        End If
                    Next ilBus1
                    If (slHours = "NNNNNNNNNNNNNNNNNNNNNNNN") And (slDays = "NNNNNNN") And (ilBusCount = 0) And (grdLib.TextMatrix(grdLib.FixedRows, USEDFLAGINDEX) <> "Y") Then
                        If (Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) <> 0) Then
                            lmDeleteCodes(UBound(lmDeleteCodes)) = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
                            ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
                        End If
                        grdLibEvents.RemoveItem llRow
                    Else
                        If (slHours = "NNNNNNNNNNNNNNNNNNNNNNNN") And (slDays = "NNNNNNN") And (ilBusCount = 0) Then
                            grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = ""
                            grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = ""
                            grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = ""
                        Else
                            If ilBusCount > 0 Then
                                grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = slBuses
                            End If
                            If (slHours <> "NNNNNNNNNNNNNNNNNNNNNNNN") Then
                                grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = gHourMap(slHours)
                            End If
                            If (slDays <> "NNNNNNN") Then
                                grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = gDayMap(slDays)
                            End If
                        End If
                        grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                    End If
                End If
            End If
        Next llRow
        For ilExtract = 0 To UBound(tgExtract) - 1 Step 1
            If tgExtract(ilExtract).lLinkBus <> 0 Then
                ilMatch = False
                For llRow = grdLibEvents.Rows - 1 To grdLibEvents.FixedRows Step -1
                    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                        slStr = grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX)
                        slHours = gCreateHourStr(slStr)
                        If mCompareExtract(llRow, tgExtract(ilExtract)) And (tgExtract(ilExtract).sHours = slHours) Then
                            ilMatch = True
                            'For ilHour = 1 To 24 Step 1
                            '    slHours = tgExtract(ilExtract).sHours
                            '    If Mid$(slHours, ilHour, 1) = "Y" Then
                            '        slStr = grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX)
                            '        slHours = gCreateHourStr(slStr)
                            '        Mid$(slHours, ilHour, 1) = "Y"
                            '        grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = gHourMap(slHours)
                            '    End If
                            'Next ilHour
                            'Create bus string
                            slBuses = ""
                            For ilBus1 = LBound(smBuses) To UBound(smBuses) Step 1
                                If Trim(smBuses(ilBus1)) <> "" Then
                                    If slBuses = "" Then
                                        slBuses = smBuses(ilBus1)
                                    Else
                                        slBuses = slBuses & ", " & smBuses(ilBus1)
                                    End If
                                End If
                            Next ilBus1
                            For ilBus2 = 0 To UBound(sgExtractBusNames) - 1 Step 1
                                ilBusMatch = False
                                For ilBus1 = LBound(smBuses) To UBound(smBuses) Step 1
                                    If Trim(smBuses(ilBus1)) <> "" Then
                                        If StrComp(Trim$(smBuses(ilBus1)), Trim$(sgExtractBusNames(ilBus2)), vbTextCompare) = 0 Then
                                            ilBusMatch = True
                                            Exit For
                                        End If
                                    End If
                                Next ilBus1
                                If Not ilBusMatch Then
                                    If slBuses = "" Then
                                        slBuses = sgExtractBusNames(ilBus2)
                                    Else
                                        slBuses = slBuses & ", " & sgExtractBusNames(ilBus2)
                                    End If
                                End If
                            Next ilBus2
                            grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = slBuses
                            grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                            imFieldChgd = True
                            Exit For
                        End If
                    End If
                Next llRow
                If Not ilMatch Then
                    imFieldChgd = True
                    'Add Row
                    llRow = grdLibEvents.Rows - 1
                    Do
                        
                        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                            Exit Do
                        End If
                        llRow = llRow - 1
                    Loop While llRow > grdLibEvents.FixedRows
                    If llRow = grdLibEvents.FixedRows Then
                        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                            llRow = llRow + 1
                            If llRow >= grdLibEvents.Rows Then
                                grdLibEvents.AddItem ""
                            End If
                        End If
                    Else
                        llRow = llRow + 1
                        If llRow >= grdLibEvents.Rows Then
                            grdLibEvents.AddItem ""
                        End If
                    End If
                    'Move Extract into Grid
                    If tgExtract(ilExtract).sEventType = "P" Then
                        grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "Program"
                    Else
                        grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "Avail"
                    End If
                    slBuses = ""
                    slStr = String(7, "N")
                    grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = gDayMap(slStr)
                    ilIndex = ilExtract
                    Do
                        If slBuses = "" Then
                            slBuses = Trim$(tgExtract(ilIndex).sBus)
                        Else
                            slBuses = slBuses & ", " & Trim$(tgExtract(ilIndex).sBus)
                        End If
                        For ilDay = 1 To 7 Step 1
                            slStr = tgExtract(ilIndex).sDays
                            If Mid(slStr, ilDay, 1) = "Y" Then
                                slStr = grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX)
                                If slStr <> "" Then
                                    slDays = gCreateDayStr(slStr)
                                Else
                                    slDays = String(7, "N")
                                End If
                                Mid(slDays, ilDay, 1) = "Y"
                                grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = gDayMap(slDays)
                            End If
                        Next ilDay
                        ilNext = tgExtract(ilIndex).lLinkBus
                        tgExtract(ilIndex).lLinkBus = 0
                        ilIndex = ilNext
                    Loop While ilIndex <> -1
                    grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = slBuses
                    grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX) = tgExtract(ilExtract).sBusCtrl
                    grdLibEvents.TextMatrix(llRow, TIMEINDEX) = tgExtract(ilExtract).sOffset
                    grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = tgExtract(ilExtract).sStartType
                    grdLibEvents.TextMatrix(llRow, FIXEDINDEX) = tgExtract(ilExtract).sFixedTime
                    grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = tgExtract(ilExtract).sEndType
                    grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = tgExtract(ilExtract).sDuration
                    grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = gHourMap(tgExtract(ilExtract).sHours)
                    grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = tgExtract(ilExtract).sMaterialType
                    grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = tgExtract(ilExtract).sAudioName
                    grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = tgExtract(ilExtract).sAudioID
                    grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX) = tgExtract(ilExtract).sAudioISCI
                    grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = tgExtract(ilExtract).sAudioCtrl
                    grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = tgExtract(ilExtract).sBackupName
                    grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = tgExtract(ilExtract).sBackupCtrl
                    grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = tgExtract(ilExtract).sProtName
                    grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX) = tgExtract(ilExtract).sProtItemID
                    grdLibEvents.TextMatrix(llRow, PROTISCIINDEX) = tgExtract(ilExtract).sProtISCI
                    grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = tgExtract(ilExtract).sProtCtrl
                    grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = tgExtract(ilExtract).sRelay1
                    grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = tgExtract(ilExtract).sRelay2
                    grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = tgExtract(ilExtract).sFollow
                    grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = tgExtract(ilExtract).sSilenceTime
                    grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = tgExtract(ilExtract).sSilence1
                    grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = tgExtract(ilExtract).sSilence2
                    grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = tgExtract(ilExtract).sSilence3
                    grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = tgExtract(ilExtract).sSilence4
                    grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = tgExtract(ilExtract).sNetcue1
                    grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = tgExtract(ilExtract).sNetcue2
                    grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = tgExtract(ilExtract).sTitle1
                    grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = tgExtract(ilExtract).sTitle2
                    grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = ""
                    grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = ""
                    grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = ""
                    grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = ""
                    grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
                    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                    'grdLibEvents.AddItem ""
                    
                End If
            End If
        Next ilExtract
    End If
    mSetCommands
End Sub

Private Sub cmcNone_Click()
    Dim llRg As Long
    Dim llRet As Long
    Dim ilValue As Integer
    
    ilValue = False
    If lbcBGE.ListCount > 0 Then         'at least 1 entries exists in check box
        llRg = CLng(lbcBGE.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcBGE.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    If grdLib.TextMatrix(grdLib.Row, BUSGROUPSINDEX) <> "" Then
        grdLib.CellForeColor = vbBlack
        grdLib.TextMatrix(grdLib.Row, BUSGROUPSINDEX) = ""
        imFieldChgd = True
        mSetCommands
    End If
End Sub

Private Sub cmcReplace_Click()
    Dim ilCol As Integer
    Dim ilFilter As Integer
    Dim ilIndex As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    gSetMousePointer grdLib, grdLibEvents, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdLib, grdLibEvents, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Library Definition"
        Exit Sub
    End If
    ReDim tgLibReplaceValues(0 To 0) As LIBREPLACEVALUES
    mMoveDEECtrlsToRec
    mCreateUsedArrays
    mInitReplaceInfo
    igAnsReplace = 0
    igReplaceCallInfo = 1
    sgReplaceDefaultHours = grdLib.TextMatrix(grdLib.FixedRows, HOURSINDEX)
    gSetMousePointer grdLib, grdLibEvents, vbDefault
    EngrReplaceLib.Show vbModal
    If igAnsReplace = CALLDONE Then 'Apply
        gSetMousePointer grdLib, grdLibEvents, vbHourglass
        grdLibEvents.Redraw = False
        mReplaceValues
        grdLibEvents.Redraw = True
    End If
    gSetMousePointer grdLib, grdLibEvents, vbDefault
End Sub

Private Sub cmcReplace_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    If bmInSave Then
        Exit Sub
    End If
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_SearchByType(1, grdLibEvents, BUSNAMEINDEX, slStr)
    If llRow >= 0 Then
        mEEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub


Private Function mNameOk() As Integer
    Dim llRow As Long
    Dim slStrName As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMatch As Integer
    Dim slStrSubname As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slMsg As String
    Dim ilDHE As Integer
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim llDheCode As Long
    Dim llNowDate As Long
    
    imOverlapCase = 0
    lmCurrentDHE = 0
    imGapExist = False
    If (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Dormant") Or (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo") Then
        If lgLibCallCode > 0 Then
            smOverlapMsg = "Ok to Change Library"
        Else
            smOverlapMsg = "Ok to Add Library"
        End If
        mNameOk = True
        Exit Function
    End If
    llRow = grdLib.FixedRows
    If Trim$(grdLib.TextMatrix(llRow, CODEINDEX)) = "" Then
        llDheCode = 0
    Else
        llDheCode = Val(grdLib.TextMatrix(llRow, CODEINDEX))
    End If
    
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", sgCurrLibDHEStamp, "EngrLibDef-mNameOk", tgCurrLibDHE())
    slStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
    If Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX)) <> "" Then
        slEndDate = Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX))
    Else
        slEndDate = "12/31/2069"
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    slStrName = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
    slStrSubname = Trim$(grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX))
    smOverlapMsg = ""
    smNowDate = Format(gNow(), sgShowDateForm)
    llNowDate = gDateValue(smNowDate)
    For ilDHE = 0 To UBound(tgCurrLibDHE) - 1 Step 1
        If (tgCurrLibDHE(ilDHE).sState <> "L") And (tgCurrLibDHE(ilDHE).sState <> "D") Then
            For ilDNE = 0 To UBound(tgCurrLibDNE) - 1 Step 1
                If tgCurrLibDHE(ilDHE).lDneCode = tgCurrLibDNE(ilDNE).lCode Then
                    If StrComp(Trim$(tgCurrLibDNE(ilDNE).sName), slStrName, vbTextCompare) = 0 Then
                        For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                            If tgCurrLibDHE(ilDHE).lDseCode = tgCurrDSE(ilDSE).lCode Then
                                If StrComp(Trim$(tgCurrDSE(ilDSE).sName), slStrSubname, vbTextCompare) = 0 Then
                                    'Check dates
                                    slStartDate = tgCurrLibDHE(ilDHE).sStartDate
                                    slEndDate = tgCurrLibDHE(ilDHE).sEndDate
                                    'If (tmDHE.lCode <> tgCurrLibDHE(ilDHE).lCode) Or ((llDheCode = tgCurrLibDHE(ilDHE).lCode) And ((gDateValue(slStartDate) <> llStartDate) Or (gDateValue(slEndDate) <> llEndDate))) Then
                                    If (llDheCode <> tgCurrLibDHE(ilDHE).lCode) Then
                                        If (gDateValue(slEndDate) >= llStartDate) And (gDateValue(slStartDate) <= llEndDate) Then
                                            lmCurrentDHE = tgCurrLibDHE(ilDHE).lCode
                                            'Overlap
                                            If smOverlapMsg = "" Then
                                                'Dtermine the case: 1: Total replace; 2: Terminate Old; 3: Change Start date of Old; 4: Split old into two parts
                                                If (llStartDate <= gDateValue(slStartDate)) And (llEndDate >= gDateValue(slEndDate)) Then
                                                    smOverlapMsg = "Replace Current with this version"
                                                    imOverlapCase = 1
                                                ElseIf (llStartDate > gDateValue(slStartDate)) And (llEndDate >= gDateValue(slEndDate)) Then
                                                    smOverlapMsg = "Terminating Current as of " & Format$(llStartDate - 1, "ddddd")
                                                    imOverlapCase = 2
                                                ElseIf (llStartDate <= gDateValue(slStartDate)) And (llEndDate < gDateValue(slEndDate)) Then
                                                    smOverlapMsg = "Save disallowed as this Library is overlapping previously entered Library, modify previous Library first"
                                                    imOverlapCase = 3
                                                ElseIf (llStartDate > gDateValue(slStartDate)) And (llEndDate < gDateValue(slEndDate)) Then
                                                    'smOverlapMsg = "Splitting Current into two pieces. One ending on " & Format$(llStartDate - 1, "ddddd") & " and the other start on " & Format$(llEndDate + 1, "ddddd")
                                                    smOverlapMsg = "Save disallowed as this Library is splitting a previously entered Library, modify previous Library first"
                                                    imOverlapCase = 4
                                                End If
                                            Else
                                                MsgBox " Too many overlapping dates to process, reduce date range", vbOKOnly + vbExclamation, "Overlapping Dates"
                                                mNameOk = False
                                                Exit Function
                                            End If
                                        Else
                                            If gDateValue(slEndDate) >= llNowDate Then
                                                If (gDateValue(slEndDate) < llStartDate) Then
                                                    If imGapExist Then
                                                        If lmGapEndDate < gDateValue(slEndDate) Then
                                                            lmGapDHE = tgCurrLibDHE(ilDHE).lCode
                                                            lmGapEndDate = gDateValue(slEndDate)
                                                        End If
                                                    Else
                                                        imGapExist = True
                                                        lmGapDHE = tgCurrLibDHE(ilDHE).lCode
                                                        lmGapEndDate = gDateValue(slEndDate)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next ilDSE
                        'Exit For
                    End If
                End If
            Next ilDNE
        End If
    Next ilDHE
            
    If imOverlapCase <> 0 Then
        imGapExist = False
    End If
    '5/31/11: Disallow expanding gap
    'If imGapExist Then
    '    If lmGapEndDate = llStartDate - 1 Then
    '        imGapExist = False
    '    End If
    'End If
    imGapExist = False
    If smOverlapMsg = "" Then
        If lgLibCallCode > 0 Then
            smOverlapMsg = "Ok to Change Library"
        Else
            smOverlapMsg = "Ok to Add Library"
        End If
    Else
        '5/31/11: Disallow changes schedule area
        If imOverlapCase = 3 Then
            MsgBox smOverlapMsg, vbOKOnly + vbExclamation, "Overlapping Dates"
            mNameOk = False
            Exit Function
        End If
        If imOverlapCase = 4 Then
            MsgBox smOverlapMsg, vbOKOnly + vbExclamation, "Overlapping Dates"
            mNameOk = False
            Exit Function
        End If
    End If
    mNameOk = True
    
End Function

Private Sub mSortCol(ilCol As Integer)
    Dim llEndRow As Long
    Dim llRow As Long
    Dim slStr As String
    
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    'If (ilCol = TIMEINDEX) Or (ilCol = DURATIONINDEX) Or (ilCol = SILENCETIMEINDEX) Then
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            If slStr <> "" Then
                If (ilCol = TIMEINDEX) Then
                    slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
                    slStr = Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    'grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
                ElseIf (ilCol = DURATIONINDEX) Then
                    slStr = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
                    slStr = Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    'grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
                ElseIf (ilCol = SILENCETIMEINDEX) Then
                    slStr = grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX)
                    slStr = Trim$(Str$(gLengthToLong(slStr)))    'gStrLengthInTenthToLong(slStr)
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    'grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
                Else
                    slStr = grdLibEvents.TextMatrix(llRow, ilCol)
                End If
                grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr & grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX)
            End If
        Next llRow
        If imLastColSorted = ilCol Then
            gGrid_SortByCol grdLibEvents, EVENTTYPEINDEX, SORTTIMEINDEX, SORTTIMEINDEX, imLastSort
        Else
            gGrid_SortByCol grdLibEvents, EVENTTYPEINDEX, SORTTIMEINDEX, imLastColSorted, imLastSort
        End If
        imLastColSorted = ilCol
    'Else
    '    gGrid_SortByCol grdLibEvents, EVENTTYPEINDEX, ilCol, imLastColSorted, imLastSort
    'End If
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slStr As String
    
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            ilRet = mMinHeaderFieldsDefined()
            If (ilRet) Then
                For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                    If slStr <> "" Then
                        If (Trim$(grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX)) <> "Dormant") And (Trim$(grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX)) <> "Limbo") Then
                            cmcConflict.Enabled = True
                        Else
                            cmcConflict.Enabled = False
                        End If
                        cmcSave.Enabled = True
                        Exit Sub
                    End If
                Next llRow
                cmcConflict.Enabled = False
                cmcSave.Enabled = False
            Else
                cmcConflict.Enabled = False
                cmcSave.Enabled = False
            End If
        Else
            cmcConflict.Enabled = False
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
        ilRet = mMinHeaderFieldsDefined()
        If (ilRet) And (Trim$(grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX)) <> "Dormant") And (Trim$(grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX)) <> "Limbo") Then
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                If slStr <> "" Then
                    cmcConflict.Enabled = True
                    Exit Sub
                End If
            Next llRow
            cmcConflict.Enabled = False
        Else
            cmcConflict.Enabled = False
        End If
    End If
End Sub

Private Sub mEnableBox()
    Dim slStr As String
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilFieldChgd As Integer
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilStartHour As Integer
    Dim ilEndHour As Integer
    Dim ilHour As Integer
    Dim slHour As String
    Dim slLatestAirDate As String
    
    If igLibCallType = 3 Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (igLibCallType = 4) And (grdLib.CellBackColor = LIGHTYELLOW) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igJobStatus(LIBRARYJOB) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    
    If (grdLib.Row >= grdLib.FixedRows) And (grdLib.Row < grdLib.Rows) And (grdLib.Col >= 0) And (grdLib.Col < grdLib.Cols - 1) Then
        lmEnableRow = grdLib.Row
        lmEnableCol = grdLib.Col
        Select Case grdLib.Col
            Case NAMEINDEX
                edcDropdown.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - cmcDropDown.Width - 30, grdLib.RowHeight(grdLib.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcDNE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcDNE, CLng(grdLibEvents.Height / 2)
'                If lbcDNE.Top + lbcDNE.Height > cmcCancel.Top Then
'                    lbcDNE.Top = edcDropdown.Top - lbcDNE.Height
'                End If
                slStr = grdLib.text
                'ilIndex = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcDNE, slStr)
                If ilIndex >= 0 Then
                    lbcDNE.ListIndex = ilIndex
                    edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcDNE.ListCount <= 0 Then
                        lbcDNE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        lbcDNE.ListIndex = 0
                        edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                    End If
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcDNE.Visible = True
                edcDropdown.SetFocus
            Case SUBLIBNAMEINDEX
                edcDropdown.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - cmcDropDown.Width - 30, grdLib.RowHeight(grdLib.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcDSE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcDSE, CLng(grdLibEvents.Height / 2)
'                If lbcDSE.Top + lbcDSE.Height > cmcCancel.Top Then
'                    lbcDSE.Top = edcDropdown.Top - lbcDSE.Height
'                End If
                slStr = grdLib.text
                'ilIndex = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcDSE, slStr)
                If ilIndex >= 0 Then
                    lbcDSE.ListIndex = ilIndex
                    edcDropdown.text = lbcDSE.List(lbcDSE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcDSE.ListCount <= 0 Then
                        lbcDSE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        lbcDSE.ListIndex = 0
                        edcDropdown.text = lbcDSE.List(lbcDSE.ListIndex)
                    End If
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcDSE.Visible = True
                edcDropdown.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                edcLib.MaxLength = Len(tmCTE.sComment)
                edcLib.text = grdLib.text
                edcLib.Visible = True
                edcLib.SetFocus
            Case STARTDATEINDEX  'Date
'                edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
'                edcLib.MaxLength = 10
'                edcLib.Text = grdLib.Text
'                edcLib.Visible = True
'                edcLib.SetFocus
                cccLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                'cccLib.MaxLength = 10
                cccLib.CSI_AllowTFN = False
                cccLib.CSI_AllowBlankDate = False
                cccLib.CSI_DefaultDateType = csiNextMonday
                If grdLib.text = "" Then
                    slLatestAirDate = gGetLatestSchdDate(True)
                    If gDateValue(slLatestAirDate) < gDateValue(smNowDate) Then
                        slLatestAirDate = smNowDate
                    End If
                    cccLib.text = DateAdd("d", 1, slLatestAirDate)
                Else
                    cccLib.text = grdLib.text
                End If
                cccLib.Visible = True
                cccLib.SetFocus
            Case ENDDATEINDEX  'Date
'Can't use the date control until it allows for a blank date (tfn)
'                edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
'                edcLib.MaxLength = 10
'                edcLib.text = grdLib.text
'                edcLib.Visible = True
'                edcLib.SetFocus
                cccLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                'cccLib.MaxLength = 10
                cccLib.CSI_AllowTFN = True
                cccLib.CSI_AllowBlankDate = True
                cccLib.CSI_DefaultDateType = csiNoDate
                cccLib.text = grdLib.text
                cccLib.Visible = True
                cccLib.SetFocus
           Case DAYSINDEX  'Date
'                edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
'                edcLib.MaxLength = 0
'                If Trim$(grdLib.Text) = "" Then
'                    If grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX) <> "" Then
'                        If gIsDate(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
'                            If grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "" Then
'                                If gIsDate(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) Then
'                                    If gDateValue(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) + 6 > gDateValue(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
'                                        slDay = String(7, "N")
'                                        ilStartDay = Weekday(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX), vbMonday)
'                                        ilEndDay = Weekday(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX), vbMonday)
'                                        If ilStartDay <= ilEndDay Then
'                                            For ilDay = ilStartDay To ilEndDay Step 1
'                                                Mid$(slDay, ilDay, 1) = "Y"
'                                            Next ilDay
'                                        Else
'                                            For ilDay = ilStartDay To 7 Step 1
'                                                Mid$(slDay, ilDay, 1) = "Y"
'                                            Next ilDay
'                                            For ilDay = 1 To ilEndDay Step 1
'                                                Mid$(slDay, ilDay, 1) = "Y"
'                                            Next ilDay
'                                        End If
'                                        grdLib.Text = gDayMap(slDay)
'                                    Else
'                                        grdLib.Text = "Mo-Su"
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Else
'                        If grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "" Then
'                            grdLib.Text = "M-Su"
'                        End If
'                    End If
'                End If
'                edcLib.Text = grdLib.Text
'                edcLib.Visible = True
'                edcLib.SetFocus
                dpcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                dpcLib.MaxLength = 0
                If Trim$(grdLib.text) = "" Then
                    If grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX) <> "" Then
                        If gIsDate(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                            If grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "" Then
                                If gIsDate(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) Then
                                    If gDateValue(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) + 6 > gDateValue(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                                        slDay = String(7, "N")
                                        ilStartDay = Weekday(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX), vbMonday)
                                        ilEndDay = Weekday(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX), vbMonday)
                                        If ilStartDay <= ilEndDay Then
                                            For ilDay = ilStartDay To ilEndDay Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                        Else
                                            For ilDay = ilStartDay To 7 Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                            For ilDay = 1 To ilEndDay Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                        End If
                                        grdLib.text = gDayMap(slDay)
                                    Else
                                        grdLib.text = "Mo-Su"
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "" Then
                            grdLib.text = "M-Su"
                        End If
                    End If
                End If
                dpcLib.text = grdLib.text
                dpcLib.Visible = True
                dpcLib.SetFocus
            'Case STARTTIMEINDEX  'Date
            '    edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
            '    edcLib.MaxLength = 10
            '    edcLib.Text = grdLib.Text
            '    edcLib.Visible = True
            '    edcLib.SetFocus
            'Case LENGTHINDEX  'Date
            '    edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
            '    edcLib.MaxLength = 10
            '    edcLib.Text = grdLib.Text
            '    edcLib.Visible = True
            '    edcLib.SetFocus
            Case HOURSINDEX  'Date
'                edcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
'                edcLib.MaxLength = 0
'                If Trim$(grdLib.Text) = "" Then
'                    'If grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX) <> "" Then
'                    '    If gIsTime(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX)) Then
'                    '        If grdLib.TextMatrix(grdLib.Row, LENGTHINDEX) <> "" Then
'                    '            If gIsLength(grdLib.TextMatrix(grdLib.Row, LENGTHINDEX)) Then
'                    '                ilStartHour = gTimeToLong(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX), False) \ 3600
'                    '                ilEndHour = (gTimeToLong(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX), False) + gLengthToLong(grdLib.TextMatrix(grdLib.Row, LENGTHINDEX)) - 1) \ 3600
'                    '                If ilEndHour > 23 Then
'                    '                    ilEndHour = 23
'                    '                End If
'                                    ilStartHour = 0
'                                    ilEndHour = 23
'                                    slHour = String(24, "N")
'                                    For ilHour = ilStartHour + 1 To ilEndHour + 1 Step 1
'                                        Mid$(slHour, ilHour, 1) = "Y"
'                                    Next ilHour
'                                    grdLib.Text = gHourMap(slHour)
'                    '            End If
'                    '        End If
'                    '    End If
'                    'End If
'                End If
'                edcLib.Text = grdLib.Text
'                edcLib.Visible = True
'                edcLib.SetFocus
                hpcLib.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                'hpcLib.MaxLength = 0
                If Trim$(grdLib.text) = "" Then
                    'If grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX) <> "" Then
                    '    If gIsTime(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX)) Then
                    '        If grdLib.TextMatrix(grdLib.Row, LENGTHINDEX) <> "" Then
                    '            If gIsLength(grdLib.TextMatrix(grdLib.Row, LENGTHINDEX)) Then
                    '                ilStartHour = gTimeToLong(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX), False) \ 3600
                    '                ilEndHour = (gTimeToLong(grdLib.TextMatrix(grdLib.Row, STARTTIMEINDEX), False) + gLengthToLong(grdLib.TextMatrix(grdLib.Row, LENGTHINDEX)) - 1) \ 3600
                    '                If ilEndHour > 23 Then
                    '                    ilEndHour = 23
                    '                End If
                                    ilStartHour = 0
                                    ilEndHour = 23
                                    slHour = String(24, "N")
                                    For ilHour = ilStartHour + 1 To ilEndHour + 1 Step 1
                                        Mid$(slHour, ilHour, 1) = "Y"
                                    Next ilHour
                                    grdLib.text = gHourMap(slHour)
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If
                hpcLib.text = grdLib.text
                hpcLib.Visible = True
                hpcLib.SetFocus
            Case BUSGROUPSINDEX
                'edcDropdown.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - cmcDropdown.Width - 30, grdLib.RowHeight(grdLib.Row) - 15
                'cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
                pbcDefine.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                cmcDefine.Move pbcDefine.Left, pbcDefine.Top + pbcDefine.Height, pbcDefine.Width, pbcDefine.Height
                cmcNone.Move pbcDefine.Left, cmcDefine.Top + cmcDefine.Height, pbcDefine.Width, pbcDefine.Height
                lbcBGE.Move pbcDefine.Left, cmcNone.Top + cmcNone.Height, pbcDefine.Width
                gSetListBoxHeight lbcBGE, CLng(grdLibEvents.Height / 2)
                If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
                    cmcDefine.Caption = "[New]"
                Else
                    cmcDefine.Caption = "[View]"
                End If
'                If lbcBGE.Top + lbcBGE.Height > cmcCancel.Top Then
'                    lbcBGE.Top = edcDropdown.Top - lbcBGE.Height
'                End If
                slStr = grdLib.text
                gParseCDFields slStr, False, smBusGroups()
                ilFieldChgd = imFieldChgd
                lbcBGE.ListIndex = -1
                For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                    lbcBGE.Selected(ilLoop) = False
                Next ilLoop
                For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
                    slStr = Trim$(smBusGroups(ilLoop))
                    If slStr <> "" Then
                        'llRow = SendMessageByString(lbcBGE.hwnd, LB_FINDSTRING, -1, slStr)
                        llRow = gListBoxFind(lbcBGE, slStr)
                        If llRow >= 0 Then
                            lbcBGE.Selected(llRow) = True
                        End If
                    End If
                Next ilLoop
                imFieldChgd = ilFieldChgd
                mSetCommands
                'edcDropdown.Visible = True
                'cmcDropdown.Visible = True
                cmcDefine.Visible = True
                cmcNone.Visible = True
                pbcDefine.Visible = True
                lbcBGE.Visible = True
                'edcDropdown.SetFocus
                lbcBGE.SetFocus
            Case BUSESINDEX
                pbcDefine.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                cmcDefine.Move pbcDefine.Left, pbcDefine.Top + pbcDefine.Height, pbcDefine.Width, pbcDefine.Height
                lbcBDE.Move pbcDefine.Left, cmcDefine.Top + cmcDefine.Height, pbcDefine.Width
                gSetListBoxHeight lbcBDE, CLng(grdLibEvents.Height / 2)
                If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
                    cmcDefine.Caption = "[New]"
                Else
                    cmcDefine.Caption = "[View]"
                End If
'                If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
'                    lbcBDE.Top = edcDropdown.Top - lbcBDE.Height
'                End If
                slStr = grdLib.text
                gParseCDFields slStr, False, smBuses()
                ilFieldChgd = imFieldChgd
                lbcBDE.ListIndex = -1
                For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
                    lbcBDE.Selected(ilLoop) = False
                Next ilLoop
                For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
                    slStr = Trim$(smBuses(ilLoop))
                    If slStr <> "" Then
                        'llRow = SendMessageByString(lbcBGE.hwnd, LB_FINDSTRING, -1, slStr)
                        llRow = gListBoxFind(lbcBDE, slStr)
                        If llRow >= 0 Then
                            lbcBDE.Selected(llRow) = True
                        End If
                    End If
                Next ilLoop
                imFieldChgd = ilFieldChgd
                mSetCommands
                'edcDropdown.Visible = True
                'cmcDropDown.Visible = True
                'lbcBDE.Visible = True
                'edcDropdown.SetFocus
                cmcDefine.Visible = True
                pbcDefine.Visible = True
                lbcBDE.Visible = True
                lbcBDE.SetFocus
            Case CONFLICTINDEX
                pbcConflict.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                smConflict = grdLib.text
                If (Trim$(smConflict) = "") Or (smConflict = "Missing") Then
                    smConflict = "No"
                End If
                pbcConflict.Visible = True
                pbcConflict.SetFocus
            Case STATEINDEX
                pbcState.Move grdLib.Left + grdLib.ColPos(grdLib.Col) + 30, grdLib.Top + grdLib.RowPos(grdLib.Row) + 15, grdLib.ColWidth(grdLib.Col) - 30, grdLib.RowHeight(grdLib.Row) - 15
                smState = grdLib.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdLib.text
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
    
    If igLibCallType = 3 Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If igLibCallType = 4 Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igJobStatus(LIBRARYJOB) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If Not mMinHeaderFieldsDefined() Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (grdLibEvents.Row >= grdLibEvents.FixedRows) And (grdLibEvents.Row < grdLibEvents.Rows) And (grdLibEvents.Col >= 0) And (grdLibEvents.Col < grdLibEvents.Cols - 1) Then
        lmEEnableRow = grdLibEvents.Row
        mPaintRowColor grdLibEvents.Row
        ilCol = grdLibEvents.Col
        If grdLibEvents.Col >= TITLE1INDEX Then
            grdLibEvents.LeftCol = grdLibEvents.LeftCol + 6
            'This do event is required so that the column is moved now
            DoEvents
        End If
        If grdLibEvents.Col <= STARTTYPEINDEX Then
            grdLibEvents.LeftCol = HIGHLIGHTINDEX
            'This do event is required so that the column is moved now
            DoEvents
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
            If grdLibEvents.ColIsVisible(ilCol) And (grdLibEvents.ColWidth(ilCol) > 0) Then
                llColPos = llColPos + grdLibEvents.ColWidth(ilCol)
            End If
        Next ilCol
        Select Case grdLibEvents.Col
            Case HIGHLIGHTINDEX
                pbcHighlight.Left = -400
                grdLibEvents.text = "»"
                pbcArrow.Visible = False
            Case BUSNAMEINDEX
                lbcBuses.Clear
                slStr = grdLib.TextMatrix(grdLib.FixedRows, BUSESINDEX)
                gParseCDFields slStr, False, smBuses()
                For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
                    slStr = Trim$(smBuses(ilLoop))
                    lbcBuses.AddItem slStr
                    llRow = gListBoxFind(lbcBDE, slStr)
                    If llRow >= 0 Then
                        lbcBuses.ItemData(lbcBuses.NewIndex) = lbcBDE.ItemData(llRow)
                    End If
                Next ilLoop
'                pbcEDefine.Move grdLibEvents.Left + grdLibEvents.ColPos(grdLibEvents.Col) + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                pbcEDefine.Width = gSetCtrlWidth("BusName", lmCharacterWidth, pbcEDefine.Width, 0)
'                lbcBuses.Move pbcEDefine.Left, pbcEDefine.Top + pbcEDefine.Height, pbcEDefine.Width
'                gSetListBoxHeight lbcBuses, CLng(grdLibEvents.Height / 2)
'                If lbcBuses.Top + lbcBuses.Height > cmcCancel.Top Then
'                    lbcBuses.Top = edcEDropdown.Top - lbcBuses.Height
'                End If
                slStr = grdLibEvents.text
                'ilFieldChgd = imFieldChgd
                If slStr <> "" Then
                    gParseCDFields slStr, False, smBuses()
                    lbcBuses.ListIndex = -1
                    For ilLoop = 0 To lbcBuses.ListCount - 1 Step 1
                        lbcBuses.Selected(ilLoop) = False
                    Next ilLoop
                    For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
                        slStr = Trim$(smBuses(ilLoop))
                        If slStr <> "" Then
                            llRow = gListBoxFind(lbcBuses, slStr)
                            If llRow >= 0 Then
                                lbcBuses.Selected(llRow) = True
                            End If
                        End If
                    Next ilLoop
                Else
                    For ilLoop = 0 To lbcBuses.ListCount - 1 Step 1
                        lbcBuses.Selected(ilLoop) = True
                    Next ilLoop
                End If
                'imFieldChgd = ilFieldChgd
                'mSetCommands
                'edcDropdown.Visible = True
                'cmcDropDown.Visible = True
                'lbcBuses.Visible = True
                'edcDropdown.SetFocus
                lacHelp.Caption = "Select Bus or Buses.  If the one that you want is not shown, then add it to the Default Buses selected in the Header area"
'                pbcEDefine.Visible = True
'                lbcBuses.Visible = True
'                lbcBuses.SetFocus
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
                lacHelp.Caption = "Select Bus Control.  Default value set on Bus Screen.  If multi-buses selected, then the first default is shown"
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
''                edcEvent.MaxLength = gSetMaxChars("Time", 0)
''                imMaxColChars = gGetMaxChars("Time")
''                edcEvent.Text = grdLibEvents.Text
''                lacHelp.Caption = "Enter Time offset of event from the Start Time defined in the Header area.  Time format is hh:mm:ss.t"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
'                ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                ltcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, ltcEvent.Width, 0)
                'ltcEvent.MaxLength = gSetMaxChars("Time", 0)
                'imMaxColChars = gGetMaxChars("Time")
                slStr = grdLibEvents.text
                ltcEvent.CSI_UseHours = False
                ltcEvent.CSI_UseTenths = True
                If Not gIsLengthTenths(slStr) Then
                    ltcEvent.text = ""
                Else
                    ltcEvent.text = ""
                    ltcEvent.text = slStr   'grdLibEvents.Text
                End If
                lacHelp.Caption = "Enter Time offset of event from the Start Time defined in Air Information area.  Time format is mm:ss.t"
'                ltcEvent.Visible = True
'                ltcEvent.SetFocus
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
                    smYN = "Y"  '"N"
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
                lacHelp.Caption = "Select End Time Type parameter."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcTTE_E.Visible = True
'                edcEDropdown.SetFocus
            Case DURATIONINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, edcEvent.Width, 0)
''                edcEvent.MaxLength = gSetMaxChars("Duration", 0)
''                imMaxColChars = gGetMaxChars("Duration")
''                edcEvent.Text = grdLibEvents.Text
''                lacHelp.Caption = "Enter the length of this event.  If entered, then the End Time Type must not be entered. Format is hh:mm:ss.t"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
'                ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                ltcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, ltcEvent.Width, 0)
                slStr = grdLibEvents.text
                ltcEvent.CSI_UseHours = True
                ltcEvent.CSI_UseTenths = True
                If Not gIsLengthTenths(slStr) Then
                    ltcEvent.text = ""
                Else
                    ltcEvent.text = ""
                    ltcEvent.text = slStr 'grdLibEvents.Text
                End If
                lacHelp.Caption = "Enter the length of this event.  Format is hh:mm:ss.t"
'                ltcEvent.Visible = True
'                ltcEvent.SetFocus
            Case AIRHOURSINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.MaxLength = 0
''                imMaxColChars = 0
''                If grdLibEvents.Text = "" Then
''                    edcEvent.Text = grdLib.TextMatrix(grdLib.FixedRows, HOURSINDEX)
''                Else
''                    edcEvent.Text = grdLibEvents.Text
''                End If
''                lacHelp.Caption = "Enter the hours to air.  Enter Hour # separated by commas and pair of numbers separated by dash (6-10, 13, 15)"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
'                hpcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
                slStr = grdLibEvents.text
                'hpcEvent.MaxLength = 0
                hpcEvent.text = ""
                If slStr = "" Then
                    hpcEvent.text = grdLib.TextMatrix(grdLib.FixedRows, HOURSINDEX)
                Else
                    hpcEvent.text = slStr
                End If
                lacHelp.Caption = "Enter the hours to air.  Enter Hour # separated by commas and pair of numbers separated by dash (6-10, 13, 15)"
'                hpcEvent.Visible = True
'                hpcEvent.SetFocus
            Case AIRDAYSINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.MaxLength = 0
''                imMaxColChars = 0
''                If grdLibEvents.Text = "" Then
''                    edcEvent.Text = grdLib.TextMatrix(grdLib.FixedRows, DAYSINDEX)
''                Else
''                    edcEvent.Text = grdLibEvents.Text
''                End If
''                lacHelp.Caption = "Enter the days to air.  The days must be within the define days (start date to end date). Enter day names separated by commas and pair of names separated by dash (Mo-Fr, Su)"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
'                dpcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
                slStr = grdLibEvents.text
                dpcEvent.MaxLength = 0
                dpcEvent.text = ""
                If slStr = "" Then
                    dpcEvent.text = grdLib.TextMatrix(grdLib.FixedRows, DAYSINDEX)
                Else
                    dpcEvent.text = slStr
                End If
                lacHelp.Caption = "Enter the days to air.  The days must be within the define days (start date to end date). Enter day names separated by commas and pair of names separated by dash (Mo-Fr, Su)"
'                dpcEvent.Visible = True
'                dpcEvent.SetFocus
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
                    ilFound = False
                    'If Event Type is Avail, then get default from Bus
                    If Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX)) <> "" Then
                        slStr = Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX))
                        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                                If tgCurrETE(ilETE).sCategory = "A" Then
                                    slBuses = Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, BUSNAMEINDEX))
                                    gParseCDFields slBuses, False, smBuses()
                                    For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
                                        slStr = Trim$(smBuses(ilLoop))
                                        If slStr <> "" Then
                                            For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                                If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) = 0 Then
                                                    If tgCurrBDE(ilBDE).iAseCode > 0 Then
                                                        For ilASE = 0 To lbcASE.ListCount - 1 Step 1
                                                            If lbcASE.ItemData(ilASE) = tgCurrBDE(ilBDE).iAseCode Then
                                                                lbcASE.ListIndex = ilASE
                                                                edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                                                                ilFound = True
                                                                Exit For
                                                            End If
                                                        Next ilASE
                                                    End If
                                                    If ilFound Then
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilBDE
                                            If ilFound Then
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                    If ilFound Then
                                        Exit For
                                    End If
                                ElseIf tgCurrETE(ilETE).sCategory = "P" Then
                                    'No default
                                    lbcASE.ListIndex = -1
                                    edcEDropdown.text = ""
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilETE
                    End If
                    If Not ilFound Then
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
                edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
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
                '9/13/11:  Allow netcue to be the same
                'lacHelp.Caption = "Select Netcue parameter.  Netque 1 and 2 must be different"
                lacHelp.Caption = "Select Netcue parameter."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcNNE.Visible = True
'                edcEDropdown.SetFocus
            Case TITLE1INDEX
                 mLoadCTE_1
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("Title1", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
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
                edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
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
                    edcEDropdown.text = ""
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
Private Sub mSetShow()
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilBusGroup As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilBGECode As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    Dim ilBSE As Integer
    Dim ilBus As Integer
    Dim ilRet As Integer
    
    If (lmEnableRow >= grdLib.FixedRows) And (lmEnableRow < grdLib.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case SUBLIBNAMEINDEX
            Case DESCRIPTIONINDEX
            Case STARTDATEINDEX  'Date
            Case ENDDATEINDEX  'Date
            Case DAYSINDEX  'Date
            'Case STARTTIMEINDEX  'Date
            'Case LENGTHINDEX  'Date
            Case HOURSINDEX  'Date
            Case BUSGROUPSINDEX
                'Remove any Buses associated with a group that is removed
'                ReDim ilBusSel(0 To 0) As Integer
'                For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
'                    If lbcBDE.Selected(ilLoop) Then
'                        ilFound = False
'                        For ilBusGroup = LBound(smBusGroups) To UBound(smBusGroups) Step 1
'                            slStr = Trim$(smBusGroups(ilBusGroup))
'                            If slStr <> "" Then
'                                llRow = gListBoxFind(lbcBGE, slStr)
'                                If llRow >= 0 Then
'                                    ilBGECode = lbcBGE.ItemData(llRow)
'                                    ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
'                                    For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
'                                        If tmCurrBSE(ilBSE).iBdeCode = lbcBDE.ItemData(ilLoop) Then
'                                            ilFound = True
'                                            Exit For
'                                        End If
'                                    Next ilBSE
'                                    If ilFound Then
'                                        Exit For
'                                    End If
'                                End If
'                            End If
'                        Next ilBusGroup
'                        If Not ilFound Then
'                            ilBusSel(UBound(ilBusSel)) = lbcBDE.ItemData(ilLoop)
'                            ReDim Preserve ilBusSel(0 To UBound(ilBusSel) + 1) As Integer
'                        End If
'                    End If
'                Next ilLoop
'                For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
'                    If lbcBGE.Selected(ilLoop) Then
'                        ilFound = False
'                        For ilBusGroup = LBound(smBusGroups) To UBound(smBusGroups) Step 1
'                            slStr = Trim$(smBusGroups(ilBusGroup))
'                            If slStr <> "" Then
'                                If StrComp(Trim$(lbcBGE.List(ilLoop)), slStr, vbTextCompare) = 0 Then
'                                    ilFound = True
'                                    Exit For
'                                End If
'                            End If
'                        Next ilBusGroup
'                        If Not ilFound Then
'                            ilFound = False
'                            ilBGECode = lbcBGE.ItemData(ilLoop)
'                            ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
'                            For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
'                                ilFound = False
'                                For ilBus = 0 To UBound(ilBusSel) - 1 Step 1
'                                    If tmCurrBSE(ilBSE).iBdeCode = ilBusSel(ilBus) Then
'                                        ilFound = True
'                                        Exit For
'                                    End If
'                                Next ilBus
'                                If Not ilFound Then
'                                    ilBusSel(UBound(ilBusSel)) = tmCurrBSE(ilBSE).iBdeCode
'                                    ReDim Preserve ilBusSel(0 To UBound(ilBusSel) + 1) As Integer
'                                End If
'                            Next ilBSE
'                        End If
'                    End If
'                Next ilLoop
''                For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
''                    lbcBDE.Selected(ilLoop) = False
''                Next ilLoop
                ReDim ilNewGroupBus(0 To 0) As Integer
                ReDim ilOldBusSel(0 To 0) As Integer
                ReDim ilOldGroupBus(0 To 0) As Integer
                'Get current selected buses
                For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
                    If lbcBDE.Selected(ilLoop) Then
                        ilOldBusSel(UBound(ilOldBusSel)) = lbcBDE.ItemData(ilLoop)
                        ReDim Preserve ilOldBusSel(0 To UBound(ilOldBusSel) + 1) As Integer
                    End If
                Next ilLoop
                'Get list of buses that could have been highlighted from the Groups that were previously selected
                For ilBusGroup = LBound(smBusGroups) To UBound(smBusGroups) Step 1
                    slStr = Trim$(smBusGroups(ilBusGroup))
                    If slStr <> "" Then
                        llRow = gListBoxFind(lbcBGE, slStr)
                        If llRow >= 0 Then
                            ilBGECode = lbcBGE.ItemData(llRow)
                            ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
                            For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
                                ilFound = False
                                For ilBus = 0 To UBound(ilOldGroupBus) - 1 Step 1
                                    If ilOldGroupBus(ilBus) = tmCurrBSE(ilBSE).iBdeCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilBus
                                If Not ilFound Then
                                    ilOldGroupBus(UBound(ilOldGroupBus)) = tmCurrBSE(ilBSE).iBdeCode
                                    ReDim Preserve ilOldGroupBus(0 To UBound(ilOldGroupBus) + 1) As Integer
                                End If
                            Next ilBSE
                            'If ilFound Then
                            '    Exit For
                            'End If
                        End If
                    End If
                Next ilBusGroup
                'Get Buses from current selected Groups
                For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                    If lbcBGE.Selected(ilLoop) Then
                        ilFound = False
                        ilBGECode = lbcBGE.ItemData(ilLoop)
                        ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
                        For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
                            ilFound = False
                            For ilBus = 0 To UBound(ilNewGroupBus) - 1 Step 1
                                If tmCurrBSE(ilBSE).iBdeCode = ilNewGroupBus(ilBus) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilBus
                            If Not ilFound Then
                                ilNewGroupBus(UBound(ilNewGroupBus)) = tmCurrBSE(ilBSE).iBdeCode
                                ReDim Preserve ilNewGroupBus(0 To UBound(ilNewGroupBus) + 1) As Integer
                            End If
                        Next ilBSE
                    End If
                Next ilLoop
                'De-select items from old bus groups
                imIgnoreBDEChg = True
                For ilLoop = 0 To UBound(ilOldGroupBus) - 1 Step 1
                    For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                        If ilOldGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                            lbcBDE.Selected(ilBDE) = False
                            Exit For
                        End If
                    Next ilBDE
                Next ilLoop
                For ilLoop = 0 To UBound(ilNewGroupBus) - 1 Step 1
                    ilFound = False
                    For ilBus = 0 To UBound(ilOldGroupBus) - 1 Step 1
                        If ilNewGroupBus(ilLoop) = ilOldGroupBus(ilBus) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilBus
                    If Not ilFound Then
                        For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                            If ilNewGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                                lbcBDE.Selected(ilBDE) = True
                                Exit For
                            End If
                        Next ilBDE
                    Else
                        'Was it Selected previously
                        ilFound = False
                        For ilBus = 0 To UBound(ilOldBusSel) - 1 Step 1
                            If ilNewGroupBus(ilLoop) = ilOldBusSel(ilBus) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilBus
                        If ilFound Then
                            For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                                If ilNewGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                                    lbcBDE.Selected(ilBDE) = True
                                    Exit For
                                End If
                            Next ilBDE
                        End If
                    End If
                Next ilLoop
                slStr = ""
                For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                    If lbcBDE.Selected(ilBDE) Then
                        If slStr = "" Then
                            slStr = lbcBDE.List(ilBDE)
                        Else
                            slStr = slStr & ", " & lbcBDE.List(ilBDE)
                        End If
                    End If
                Next ilBDE
                grdLib.TextMatrix(lmEnableRow, BUSESINDEX) = slStr
                imIgnoreBDEChg = False
            Case BUSESINDEX
                lbcBuses.Clear
                For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
                    If lbcBDE.Selected(ilLoop) Then
                        lbcBuses.AddItem lbcBDE.List(ilLoop)
                        lbcBuses.ItemData(lbcBuses.NewIndex) = lbcBDE.ItemData(ilLoop)
                    End If
                Next ilLoop
            Case CONFLICTINDEX
                grdLib.TextMatrix(lmEnableRow, lmEnableCol) = smConflict
                If (Trim$(grdLib.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdLib.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdLib.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdLib.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
    End If
    cccLib.Visible = False
    dpcLib.Visible = False
    hpcLib.Visible = False
    edcLib.Visible = False
    lbcDNE.Visible = False
    lbcDSE.Visible = False
    lbcBGE.Visible = False
    lbcBDE.Visible = False
    pbcDefine.Visible = False
    cmcNone.Visible = False
    cmcDefine.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    pbcConflict.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub mESetShow()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim slOrigValue As String
    
    If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
        Select Case lmEEnableCol
            Case HIGHLIGHTINDEX
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = ""
            Case BUSNAMEINDEX
            Case BUSCTRLINDEX
            Case EVENTTYPEINDEX
                mSetColExportColor lmEEnableRow
            Case TIMEINDEX
            Case STARTTYPEINDEX
            Case FIXEDINDEX
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = smYN
            Case ENDTYPEINDEX
            Case DURATIONINDEX
            Case AIRHOURSINDEX
            Case AIRDAYSINDEX
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
                '12/1/11: Remove question
                'slStr = UCase(Trim$(edcEDropdown.text))
                'If (slStr <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)) <> "") Then
                '    If UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))) <> slStr Then
                '        ilRet = MsgBox("Change all occurrences of this Comment within this Library", vbQuestion + vbYesNo + vbDefaultButton2, "Comment Changed")
                '        If ilRet = vbYes Then
                '            slOrigValue = UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)))
                '            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                '                If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                '                    If UCase(Trim$(grdLibEvents.TextMatrix(llRow, lmEEnableCol))) = slOrigValue Then
                '                        grdLibEvents.TextMatrix(llRow, lmEEnableCol) = Trim$(edcEDropdown.text)
                '                    End If
                '                End If
                '            Next llRow
                '        End If
                '    End If
                'End If
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = Trim$(edcEDropdown.text)
            Case TITLE2INDEX
                '12/1/11: Remove question
                'slStr = UCase(Trim$(edcEDropdown.text))
                'If (slStr <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)) <> "") Then
                '    If UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))) <> slStr Then
                '        ilRet = MsgBox("Change all occurrences of this Comment within this Library", vbQuestion + vbYesNo + vbDefaultButton2, "Comment Changed")
                '        If ilRet = vbYes Then
                '            slOrigValue = UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)))
                '            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                '                If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                '                    If UCase(Trim$(grdLibEvents.TextMatrix(llRow, lmEEnableCol))) = slOrigValue Then
                '                        grdLibEvents.TextMatrix(llRow, lmEEnableCol) = Trim$(edcEDropdown.text)
                '                    End If
                '                End If
                '            Next llRow
                '        End If
                '    End If
                'End If
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
    lbcBuses.Visible = False
    lbcCCE_B.Visible = False
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
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    pbcHighlight.Visible = False
    mHideConflictGrid
    ltcEvent.Visible = False
    hpcEvent.Visible = False
    dpcEvent.Visible = False
    imShowGridBox = False
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
    Dim ilTestHour As Integer
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
    Dim llSTime As Long
    Dim llETime As Long
    Dim llLEndTime As Long
    Dim ilSHour As Integer
    Dim ilCol As Integer
    '5/31/11: Disallow changes in the schedule area
    Dim slLatestAirDate As String
    
    grdLib.Redraw = False
    'Test if fields defined
    ilError = False
    llRow = grdLib.FixedRows
    If ilTestState Then
        smNowDate = Format(gNow(), sgShowDateForm)
        slLatestAirDate = gGetLatestSchdDate(True)
        grdLib.Row = llRow
        For ilCol = NAMEINDEX To STATEINDEX Step 1
            grdLib.Col = ilCol
            grdLib.CellForeColor = vbBlack
        Next ilCol
    End If
    If (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Dormant") Or (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo") Then
        slStr = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            'slStr = grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX)
            'If (slStr <> "") Then
                ilError = True
                grdLib.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = NAMEINDEX
                grdLib.CellForeColor = vbRed
            'End If
        Else
            If ilTestState Then
                slStr = grdLib.TextMatrix(llRow, STARTDATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdLib.TextMatrix(llRow, STARTDATEINDEX) = "Missing"
                    grdLib.Row = llRow
                    grdLib.Col = STARTDATEINDEX
                    grdLib.CellForeColor = vbRed
                Else
                    If Not gIsDate(slStr) Then
                        ilError = True
                        grdLib.Row = llRow
                        grdLib.Col = STARTDATEINDEX
                        grdLib.CellForeColor = vbRed
                    Else
                        If igLibCallType <> 4 Then
                            If (gDateValue(slStr) <= gDateValue(slLatestAirDate)) Or (gDateValue(slStr) <= gDateValue(smNowDate)) Then
                                ilError = True
                                grdLib.Row = llRow
                                grdLib.Col = STARTDATEINDEX
                                grdLib.CellForeColor = vbRed
                            End If
                        End If
                    End If
                End If
                slStr = grdLib.TextMatrix(llRow, ENDDATEINDEX)
                If slStr <> "" Then
                    If Not gIsDate(slStr) Then
                        ilError = True
                        grdLib.Row = llRow
                        grdLib.Col = ENDDATEINDEX
                        grdLib.CellForeColor = vbRed
                    Else
                        If gDateValue(grdLib.TextMatrix(llRow, STARTDATEINDEX)) > gDateValue(slStr) Then
                            ilError = True
                            grdLib.Row = llRow
                            grdLib.Col = ENDDATEINDEX
                            grdLib.CellForeColor = vbRed
                        Else
                            If (gDateValue(slStr) <= gDateValue(slLatestAirDate)) Or (gDateValue(slStr) <= gDateValue(smNowDate)) Then
                                ilError = True
                                grdLib.Row = llRow
                                grdLib.Col = ENDDATEINDEX
                                grdLib.CellForeColor = vbRed
                            End If
                        End If
                    End If
                End If
            End If
        End If
        grdLib.Redraw = True
        If ilError Then
            mCheckFields = False
            Exit Function
        Else
            mCheckFields = True
            Exit Function
        End If
    End If
    slStr = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        'slStr = grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX)
        'If (slStr <> "") Then
            ilError = True
            grdLib.TextMatrix(llRow, NAMEINDEX) = "Missing"
            grdLib.Row = llRow
            grdLib.Col = NAMEINDEX
            grdLib.CellForeColor = vbRed
        'End If
    Else
        If ilTestState Then
            If igLibCallType = 4 Then   'Only check end date
                slStr = grdLib.TextMatrix(llRow, ENDDATEINDEX)
                If slStr <> "" Then
                    If Not gIsDate(slStr) Then
                        ilError = True
                        grdLib.Row = llRow
                        grdLib.Col = ENDDATEINDEX
                        grdLib.CellForeColor = vbRed
                        mCheckFields = False
                        Exit Function
                    End If
                    If (gDateValue(slStr) <= gDateValue(slLatestAirDate)) Or (gDateValue(slStr) <= gDateValue(smNowDate)) Then
                        ilError = True
                        grdLib.Row = llRow
                        grdLib.Col = ENDDATEINDEX
                        grdLib.CellForeColor = vbRed
                        mCheckFields = False
                        Exit Function
                    End If
                End If
                mCheckFields = True
                Exit Function
            End If
            slStr = grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = SUBLIBNAMEINDEX
                grdLib.CellForeColor = vbRed
            End If
            slAllowedDays = ""
            slStr = grdLib.TextMatrix(llRow, STARTDATEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, STARTDATEINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = STARTDATEINDEX
                grdLib.CellForeColor = vbRed
            Else
                If Not gIsDate(slStr) Then
                    ilError = True
                    grdLib.Row = llRow
                    grdLib.Col = STARTDATEINDEX
                    grdLib.CellForeColor = vbRed
                Else
'                    If igLibCallType <> 1 Then
'                        smNowDate = Format(gNow(), sgShowDateForm)
'                        If gDateValue(slStr) < gDateValue(smNowDate) Then
''Jim request- 2/3/05
''                            ilError = True
''                            grdLib.Row = llRow
''                            grdLib.Col = STARTDATEINDEX
''                            grdLib.CellForeColor = vbRed
'                        End If
'                    End If
                    If igLibCallType <> 4 Then
                        If (gDateValue(slStr) <= gDateValue(slLatestAirDate)) Or (gDateValue(slStr) <= gDateValue(smNowDate)) Then
                            ilError = True
                            grdLib.Row = llRow
                            grdLib.Col = STARTDATEINDEX
                            grdLib.CellForeColor = vbRed
                        End If
                    End If
                    slStr = grdLib.TextMatrix(llRow, ENDDATEINDEX)
                    If slStr <> "" Then
                        If gIsDate(slStr) Then
                            If gDateValue(grdLib.TextMatrix(llRow, STARTDATEINDEX)) > gDateValue(grdLib.TextMatrix(llRow, ENDDATEINDEX)) Then
                                ilError = True
                                grdLib.Row = llRow
                                grdLib.Col = ENDDATEINDEX
                                grdLib.CellForeColor = vbRed
                            Else
                                If (gDateValue(slStr) <= gDateValue(slLatestAirDate)) Or (gDateValue(slStr) <= gDateValue(smNowDate)) Then
                                    ilError = True
                                    grdLib.Row = llRow
                                    grdLib.Col = ENDDATEINDEX
                                    grdLib.CellForeColor = vbRed
                                End If
                                slAllowedDays = String(7, "N")
                                If gDateValue(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) + 6 > gDateValue(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                                    ilStartDay = Weekday(grdLib.TextMatrix(llRow, STARTDATEINDEX), vbMonday)
                                    ilEndDay = Weekday(grdLib.TextMatrix(llRow, ENDDATEINDEX), vbMonday)
                                    If ilStartDay <= ilEndDay Then
                                        For ilDay = ilStartDay To ilEndDay Step 1
                                            Mid$(slAllowedDays, ilDay, 1) = "Y"
                                        Next ilDay
                                    Else
                                        For ilDay = ilStartDay To 7 Step 1
                                            Mid$(slAllowedDays, ilDay, 1) = "Y"
                                        Next ilDay
                                        For ilDay = 1 To ilEndDay Step 1
                                            Mid$(slAllowedDays, ilDay, 1) = "Y"
                                        Next ilDay
                                    End If
                                Else
                                    For ilDay = 1 To 7 Step 1
                                        Mid$(slAllowedDays, ilDay, 1) = "Y"
                                    Next ilDay
                                End If
                                slDHEDays = gCreateDayStr(grdLib.TextMatrix(llRow, DAYSINDEX))
                                If slDHEDays = "" Then
                                    ilError = True
                                    grdLib.Row = llRow
                                    grdLib.Col = DAYSINDEX
                                    grdLib.CellForeColor = vbRed
                                Else
'2/10/05, remove checking
'                                    For ilDay = 1 To 7 Step 1
'                                        If (Mid$(slAllowedDays, ilDay, 1) = "N") And (Mid$(slDHEDays, ilDay, 1) = "Y") Then
'                                            ilError = True
'                                            grdLib.Row = llRow
'                                            grdLib.Col = DAYSINDEX
'                                            grdLib.CellForeColor = vbRed
'                                        End If
'                                    Next ilDay
                                End If
                            End If
                        End If
                    Else
                        slAllowedDays = String(7, "Y")
                    End If
                End If
            End If
            slStr = grdLib.TextMatrix(llRow, ENDDATEINDEX)
            If slStr <> "" Then
                If Not gIsDate(slStr) Then
                    ilError = True
                    grdLib.Row = llRow
                    grdLib.Col = ENDDATEINDEX
                    grdLib.CellForeColor = vbRed
                End If
            End If
            slStr = grdLib.TextMatrix(llRow, DAYSINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, DAYSINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = DAYSINDEX
                grdLib.CellForeColor = vbRed
            End If
            llEndTime = -1
            llLength = -1
            llLEndTime = -1
            ilSHour = -1
            'slAllowedHours = ""
            'slStr = grdLib.TextMatrix(llRow, STARTTIMEINDEX)
            'If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            '    ilError = True
            '    grdLib.TextMatrix(llRow, STARTTIMEINDEX) = "Missing"
            '    grdLib.Row = llRow
            '    grdLib.Col = STARTTIMEINDEX
            '    grdLib.CellForeColor = vbRed
            'Else
            '    If Not gIsTime(slStr) Then
            '        ilError = True
            '        grdLib.Row = llRow
            '        grdLib.Col = STARTTIMEINDEX
            '        grdLib.CellForeColor = vbRed
            '    Else
            '        slStr = grdLib.TextMatrix(llRow, LENGTHINDEX)
            '        If slStr <> "" Then
            '            If gIsLength(slStr) Then
            '                'Test that time is not after 12 mid
            '                slStartTime = grdLib.TextMatrix(llRow, STARTTIMEINDEX)
            '                llStartTime = gTimeToLong(slStartTime, False)
            '                llLength = gLengthToLong(grdLib.TextMatrix(llRow, LENGTHINDEX))
            '                llLEndTime = llStartTime + llLength - 1
            '                llEndTime = llLEndTime
            '                If llEndTime >= CLng(24) * 3600 Then
            '                    ilError = True
            '                    grdLib.Row = llRow
            '                    grdLib.Col = LENGTHINDEX
            '                    grdLib.CellForeColor = vbRed
            '                Else
            '                    slStr = grdLib.TextMatrix(llRow, HOURSINDEX)
            '                    If slStr <> "" Then
            '                        slDHEHours = mCreateHourStr(slStr)
            '                    End If
            '                    If slDHEHours = "" Then
            '                        ilError = True
            '                        grdLib.Row = llRow
            '                        grdLib.Col = HOURSINDEX
            '                        grdLib.CellForeColor = vbRed
            '                    Else
            '                        ilStartHour = llStartTime \ 3600
            '                        ilEndHour = llEndTime \ 3600
            '                        slAllowedHours = String(24, "N")
            '                        For ilHour = ilStartHour + 1 To ilEndHour + 1 Step 1
            '                            Mid$(slAllowedHours, ilHour, 1) = "Y"
            '                        Next ilHour
            '                        For ilHour = 1 To 24 Step 1
            '                            If (Mid$(slAllowedHours, ilHour, 1) = "N") And (Mid$(slDHEHours, ilHour, 1) = "Y") Then
            '                                ilError = True
            '                                grdLib.Row = llRow
            '                                grdLib.Col = HOURSINDEX
            '                                grdLib.CellForeColor = vbRed
            '                                Exit For
            '                            End If
            '                        Next ilHour
            '                    End If
            '                End If
            '            End If
            '        End If
            '    End If
            'End If
            'slStr = grdLib.TextMatrix(llRow, LENGTHINDEX)
            'If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            '    ilError = True
            '    grdLib.TextMatrix(llRow, LENGTHINDEX) = "Missing"
            '    grdLib.Row = llRow
            '    grdLib.Col = LENGTHINDEX
            '    grdLib.CellForeColor = vbRed
            'Else
            '    If Not gIsLength(slStr) Then
            '        ilError = True
            '        grdLib.Row = llRow
            '        grdLib.Col = LENGTHINDEX
            '        grdLib.CellForeColor = vbRed
            '    End If
            'End If
            slStr = grdLib.TextMatrix(llRow, HOURSINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, HOURSINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = HOURSINDEX
                grdLib.CellForeColor = vbRed
            Else
                slDHEHours = gCreateHourStr(slStr)
                If slDHEHours = "" Then
                    ilError = True
                    grdLib.Row = llRow
                    grdLib.Col = HOURSINDEX
                    grdLib.CellForeColor = vbRed
                End If
            End If
            slStr = grdLib.TextMatrix(llRow, BUSESINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, BUSESINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = BUSESINDEX
                grdLib.CellForeColor = vbRed
            End If
            slStr = grdLib.TextMatrix(llRow, CONFLICTINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = CONFLICTINDEX
                grdLib.CellForeColor = vbRed
            End If
            slStr = grdLib.TextMatrix(llRow, STATEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdLib.TextMatrix(llRow, STATEINDEX) = "Missing"
                grdLib.Row = llRow
                grdLib.Col = STATEINDEX
                grdLib.CellForeColor = vbRed
            End If
        End If
    End If
    If ilTestState Then
        grdLibEvents.Redraw = False
        If ilTestState Then
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
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
        'ilError = False
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            If slStr = "" Then
                slStr = grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)
                If slStr <> "" Then
                    ilError = True
                    grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "Missing"
                    grdLibEvents.Row = llRow
                    grdLibEvents.LeftCol = HIGHLIGHTINDEX
                    grdLibEvents.Col = EVENTTYPEINDEX
                    grdLibEvents.CellForeColor = vbRed
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
                    ilError = True
                    grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = "Missing"
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BUSNAMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sTime = "Y") And (tlManEPE.sTime = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TIMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TIMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                llSTime = -1
                llETime = -1
                ilSHour = -1
                llELength = -1
                If (slStr <> "") And (tlUsedEPE.sTime = "Y") Then
                    If Not gIsLengthTenths(slStr) Then
                        ilError = True
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = TIMEINDEX
                        grdLibEvents.CellForeColor = vbRed
                    Else
                        llELength = gStrLengthInTenthToLong(slStr)
                        If llELength < CLng(36000) Then
                            'Adjust to last Hour
'                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX))
'                            If (slStr <> "") Then
'                                slDEEHours = gCreateHourStr(slStr)
'                                If slDEEHours <> "" Then
'                                    For ilHour = 24 To 1 Step -1
'                                        If Mid$(slDEEHours, ilHour, 1) = "Y" Then
'                                            llETime = llELength + (ilHour - 1) * CLng(3600) * 10
'                                            ilSHour = ilHour
'                                            Exit For
'                                        End If
'                                    Next ilHour
'                                End If
'                            End If
'                            'If (llLEndTime >= 0) And (llETime >= 0) Then
'                            '    If llETime > llLEndTime * 10 Then
'                            '        ilError = True
'                            '        grdLibEvents.Row = llRow
'                            '        grdLibEvents.Col = TIMEINDEX
'                            '        grdLibEvents.CellForeColor = vbRed
'                            '    End If
'                            'End If
                        Else
                            ilError = True
                            grdLibEvents.Row = llRow
                            grdLibEvents.Col = TIMEINDEX
                            grdLibEvents.CellForeColor = vbRed
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStartType = "Y") And (tlManEPE.sStartType = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = STARTTYPEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, FIXEDINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sFixedTime = "Y") And (tlManEPE.sFixedTime = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, FIXEDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = FIXEDINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sEndType = "Y") And (tlManEPE.sEndType = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = ENDTYPEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                llDuration = -1
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, DURATIONINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sDuration = "Y") And (tlManEPE.sDuration = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = DURATIONINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                If slStr <> "" Then
                    If Not gIsLengthTenths(slStr) Then
                        ilError = True
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = DURATIONINDEX
                        grdLibEvents.CellForeColor = vbRed
                    Else
                        llDuration = gStrLengthInTenthToLong(slStr)
'                        If llELength + llDuration < CLng(36000) Then
'                            'If (llLEndTime >= 0) And (ilSHour >= 0) Then
'                            '    llETime = llELength + llDuration + (ilSHour - 1) * CLng(3600) * 10 - 1
'                            '    If llETime > 10 * llLEndTime Then
'                            '        ilError = True
'                            '        grdLibEvents.Row = llRow
'                            '        grdLibEvents.Col = DURATIONINDEX
'                            '        grdLibEvents.CellForeColor = vbRed
'                            '    End If
'                            'End If
'                        Else
'                            ilError = True
'                            grdLibEvents.Row = llRow
'                            grdLibEvents.Col = DURATIONINDEX
'                            grdLibEvents.CellForeColor = vbRed
'                        End If
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
                        ilError = True
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = DURATIONINDEX
                        grdLibEvents.CellForeColor = vbRed
                    End If
                End If
                slDEEHours = ""
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AIRHOURSINDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    slDEEHours = gCreateHourStr(slStr)
                    If slDEEHours = "" Then
                        ilError = True
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = AIRHOURSINDEX
                        grdLibEvents.CellForeColor = vbRed
                    Else
                        'If slAllowedHours <> "" Then
                        '    For ilHour = 1 To 24 Step 1
                        '        If (Mid$(slAllowedHours, ilHour, 1) = "N") And (Mid$(slDEEHours, ilHour, 1) = "Y") Then
                        '            ilError = True
                        '            grdLibEvents.Row = llRow
                        '            grdLibEvents.Col = AIRHOURSINDEX
                        '            grdLibEvents.CellForeColor = vbRed
                        '        End If
                        '    Next ilHour
                        'End If
                    End If
                End If
                'Check that times don't overlap
                If (llELength <> -1) And (slDEEHours <> "") And (llDuration <> -1) Then
                    For ilHour = 1 To 24 Step 1
                        If Mid$(slDEEHours, ilHour, 1) = "Y" Then
                            llSTime = llELength + 36000 * (ilHour - 1)
                            llETime = llSTime + llDuration  ' - 1
                            'If llETime >= 864000 Then
                            '    ilError = True
                            '    grdLibEvents.Row = llRow
                            '    grdLibEvents.Col = DURATIONINDEX
                            '    grdLibEvents.CellForeColor = vbRed
                            'Else
                                For ilTestHour = ilHour + 1 To 24 Step 1
                                    If Mid$(slDEEHours, ilTestHour, 1) = "Y" Then
                                        If llETime > llELength + 36000 * (ilTestHour - 1) Then
                                            ilError = True
                                            grdLibEvents.Row = llRow
                                            grdLibEvents.Col = DURATIONINDEX
                                            grdLibEvents.CellForeColor = vbRed
                                        End If
                                        Exit For
                                    End If
                                Next ilTestHour
                            'End If
                        End If
                    Next ilHour
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AIRDAYSINDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    slDEEDays = gCreateDayStr(grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX))
                    If slDEEDays = "" Then
                        ilError = True
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = AIRDAYSINDEX
                        grdLibEvents.CellForeColor = vbRed
                    Else
                        If slAllowedDays <> "" Then
'2/10/05- remove day test
'                            For ilDay = 1 To 7 Step 1
'                                If (Mid$(slAllowedDays, ilDay, 1) = "N") And (Mid$(slDEEDays, ilDay, 1) = "Y") Then
'                                    ilError = True
'                                    grdLibEvents.Row = llRow
'                                    grdLibEvents.Col = AIRDAYSINDEX
'                                    grdLibEvents.CellForeColor = vbRed
'                                End If
'                            Next ilDay
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, MATERIALINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sMaterialType = "Y") And (tlManEPE.sMaterialType = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = MATERIALINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sAudioName = "Y") And (tlManEPE.sAudioName = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = BACKUPNAMEINDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    'If (StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)), Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)), vbTextCompare) = 0) And (Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)) <> "") Then
                                        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX)), slStr, vbTextCompare) = 0 Then
                                            ilError = True
                                            grdLibEvents.Row = llRow
                                            grdLibEvents.Col = PROTNAMEINDEX
                                            grdLibEvents.CellForeColor = vbRed
                                        End If
                                    'End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sAudioItemID = "Y") And (tlManEPE.sAudioItemID = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOITEMIDINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sAudioISCI = "Y") And (tlManEPE.sAudioISCI = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOISCIINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sAudioControl = "Y") And (tlManEPE.sAudioControl = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOCTRLINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sBkupAudioName = "Y") And (tlManEPE.sBkupAudioName = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    'If (StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)), Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)), vbTextCompare) = 0) And (Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)) <> "") Then
                                        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX)), slStr, vbTextCompare) = 0 Then
                                            ilError = True
                                            grdLibEvents.Row = llRow
                                            grdLibEvents.Col = PROTNAMEINDEX
                                            grdLibEvents.CellForeColor = vbRed
                                        End If
                                    'End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sBkupAudioControl = "Y") And (tlManEPE.sBkupAudioControl = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BACKUPCTRLINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sProtAudioName = "Y") And (tlManEPE.sProtAudioName = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sProtAudioItemID = "Y") And (tlManEPE.sProtAudioItemID = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTITEMIDINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTISCIINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sProtAudioISCI = "Y") And (tlManEPE.sProtAudioISCI = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTISCIINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTISCIINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sProtAudioControl = "Y") And (tlManEPE.sProtAudioControl = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTCTRLINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sRelay1 = "Y") And (tlManEPE.sRelay1 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = RELAY1INDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = RELAY2INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sRelay2 = "Y") And (tlManEPE.sRelay2 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = RELAY2INDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, FOLLOWINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sFollow = "Y") And (tlManEPE.sFollow = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = FOLLOWINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                llSilence = -1
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sSilenceTime = "Y") And (tlManEPE.sSilenceTime = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCETIMEINDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                If slStr <> "" Then
                    If Not gIsLength(slStr) Then
                        ilError = True
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = SILENCETIMEINDEX
                        grdLibEvents.CellForeColor = vbRed
                    Else
                        llSilence = 10 * gLengthToLong(slStr)  'gStrLengthInTenthToLong(slStr)  'gLengthToLong(slStr)
'                        If llELength >= 0 Then
'                            If llELength + llSilence < CLng(36000) Then
'                                'If (llLEndTime >= 0) And (ilSHour >= 0) Then
'                                '    llETime = llELength + llSilence + (ilSHour - 1) * CLng(3600) * 10 - 1
'                                '    If llETime > 10 * llLEndTime Then
'                                '        ilError = True
'                                '        grdLibEvents.Row = llRow
'                                '        grdLibEvents.Col = SILENCETIMEINDEX
'                                '        grdLibEvents.CellForeColor = vbRed
'                                '    End If
'                                'End If
'                            Else
'                                ilError = True
'                                grdLibEvents.Row = llRow
'                                grdLibEvents.Col = SILENCETIMEINDEX
'                                grdLibEvents.CellForeColor = vbRed
'                            End If
'                        End If
                    End If
                End If
                If (llELength <> -1) And (slDEEHours <> "") And (llSilence <> -1) Then
                    For ilHour = 1 To 24 Step 1
                        If Mid$(slDEEHours, ilHour, 1) = "Y" Then
                            llSTime = llELength + 36000 * (ilHour - 1)
                            llETime = llSTime + llSilence - 1
                            'If llETime >= 864000 Then
                            '    ilError = True
                            '    grdLibEvents.Row = llRow
                            '    grdLibEvents.Col = SILENCETIMEINDEX
                            '    grdLibEvents.CellForeColor = vbRed
                            'Else
                                For ilTestHour = ilHour + 1 To 24 Step 1
                                    If Mid$(slDEEHours, ilTestHour, 1) = "Y" Then
                                        If llETime > llELength + 36000 * (ilTestHour - 1) Then
                                            ilError = True
                                            grdLibEvents.Row = llRow
                                            grdLibEvents.Col = SILENCETIMEINDEX
                                            grdLibEvents.CellForeColor = vbRed
                                        End If
                                        Exit For
                                    End If
                                Next ilTestHour
                            'End If
                        End If
                    Next ilHour
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence1 = "Y") And (tlManEPE.sSilence1 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE1INDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE2INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE3INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence2 = "Y") And (tlManEPE.sSilence2 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE2INDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE3INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence3 = "Y") And (tlManEPE.sSilence3 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE3INDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX)), slStr, vbTextCompare) = 0 Then
                                        ilError = True
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = vbRed
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence4 = "Y") And (tlManEPE.sSilence4 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE4INDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStartNetcue = "Y") And (tlManEPE.sStartNetcue = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = NETCUE1INDEX
                    grdLibEvents.CellForeColor = vbRed
                Else
                    '9/13/11:  Allow netcue to be the same
                    'If slStr <> "" Then
                    '    If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                    '        slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
                    '        If slStr <> "" Then
                    '            If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                    '                If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX)), slStr, vbTextCompare) = 0 Then
                    '                    ilError = True
                    '                    grdLibEvents.Row = llRow
                    '                    grdLibEvents.Col = NETCUE2INDEX
                    '                    grdLibEvents.CellForeColor = vbRed
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStopNetcue = "Y") And (tlManEPE.sStopNetcue = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = NETCUE2INDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sTitle1 = "Y") And (tlManEPE.sTitle1 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TITLE1INDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sTitle2 = "Y") And (tlManEPE.sTitle2 = "Y") Then
                    ilError = True
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TITLE2INDEX
                    grdLibEvents.CellForeColor = vbRed
                End If
                If sgClientFields = "A" Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCFormat = "Y") And (tlManEPE.sABCFormat = "Y") Then
                        ilError = True
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCFORMATINDEX
                        grdLibEvents.CellForeColor = vbRed
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCPgmCode = "Y") And (tlManEPE.sABCPgmCode = "Y") Then
                        ilError = True
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCPGMCODEINDEX
                        grdLibEvents.CellForeColor = vbRed
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCXDSMode = "Y") And (tlManEPE.sABCXDSMode = "Y") Then
                        ilError = True
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCXDSMODEINDEX
                        grdLibEvents.CellForeColor = vbRed
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCRecordItem = "Y") And (tlManEPE.sABCRecordItem = "Y") Then
                        ilError = True
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCRECORDITEMINDEX
                        grdLibEvents.CellForeColor = vbRed
                    End If
                End If
            End If
        Next llRow
    End If
    grdLibEvents.Redraw = True
    grdLib.Redraw = True
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
    
    
    gGrid_AlignAllColsLeft grdLib
    mGridColumnWidth
    'Set Titles
    grdLib.TextMatrix(0, NAMEINDEX) = "Library"
    grdLib.TextMatrix(1, NAMEINDEX) = "Name"
    grdLib.TextMatrix(0, SUBLIBNAMEINDEX) = "Subname"
    grdLib.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdLib.TextMatrix(0, STARTDATEINDEX) = "Start"
    grdLib.TextMatrix(1, STARTDATEINDEX) = "Date"
    grdLib.TextMatrix(0, ENDDATEINDEX) = "End"
    grdLib.TextMatrix(1, ENDDATEINDEX) = "Date"
    grdLib.TextMatrix(0, DAYSINDEX) = "Days"
    'grdLib.TextMatrix(0, STARTTIMEINDEX) = "Start"
    'grdLib.TextMatrix(1, STARTTIMEINDEX) = "Hour"
    'grdLib.TextMatrix(0, LENGTHINDEX) = "Length"
    grdLib.TextMatrix(0, HOURSINDEX) = "Clock"
    grdLib.TextMatrix(1, HOURSINDEX) = "Hours"
    grdLib.TextMatrix(0, BUSGROUPSINDEX) = "Bus"
    grdLib.TextMatrix(1, BUSGROUPSINDEX) = "Groups"
    grdLib.TextMatrix(0, BUSESINDEX) = "Buses"
    grdLib.TextMatrix(0, CONFLICTINDEX) = "Ignore"
    grdLib.TextMatrix(1, CONFLICTINDEX) = "Conflict"
    grdLib.TextMatrix(0, STATEINDEX) = "State"
    grdLib.Row = 1
    For ilCol = 0 To grdLib.Cols - 1 Step 1
        grdLib.Col = ilCol
        grdLib.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdLib.Height = 3 * grdLib.RowHeight(0) + 30    '15
    gGrid_IntegralHeight grdLib
    gGrid_Clear grdLib, True
    
    
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
        grdLibEvents.TextMatrix(0, ilCol) = "Sil."
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
    grdLibEvents.TextMatrix(0, TIMEINDEX) = "Offset"
    grdLibEvents.TextMatrix(1, TIMEINDEX) = "Time"
    grdLibEvents.TextMatrix(0, STARTTYPEINDEX) = "Start "
    grdLibEvents.TextMatrix(1, STARTTYPEINDEX) = "Type"
    grdLibEvents.TextMatrix(0, FIXEDINDEX) = "Fix"
    grdLibEvents.TextMatrix(0, ENDTYPEINDEX) = "End"
    grdLibEvents.TextMatrix(1, ENDTYPEINDEX) = "Type"
    grdLibEvents.TextMatrix(0, DURATIONINDEX) = "Duration"
    grdLibEvents.TextMatrix(0, AIRHOURSINDEX) = "Clock"
    grdLibEvents.TextMatrix(1, AIRHOURSINDEX) = "Hours"
    grdLibEvents.TextMatrix(0, AIRDAYSINDEX) = "Days"
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
    grdLibEvents.Height = cmcCancel.Top - grdLibEvents.Top - 240    '4 * grdLibEvents.RowHeight(0) + 15
    gGrid_IntegralHeight grdLibEvents
    gGrid_Clear grdLibEvents, True
    
    gGrid_AlignAllColsLeft grdConflicts
    mGridColumnWidth
    'Set Titles
    grdConflicts.TextMatrix(0, CONFLICTNAMEINDEX) = "Name"
    grdConflicts.TextMatrix(0, CONFLICTSUBNAMEINDEX) = "Subname"
    grdConflicts.TextMatrix(0, CONFLICTSTARTDATEINDEX) = "Start"
    grdConflicts.TextMatrix(0, CONFLICTENDDATEINDEX) = "End"
    grdConflicts.TextMatrix(0, CONFLICTDAYSINDEX) = "Days"
    grdConflicts.TextMatrix(0, CONFLICTOFFSETINDEX) = "Offset"
    grdConflicts.TextMatrix(0, CONFLICTHOURSINDEX) = "Hours"
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
    grdConflicts.Move grdLib.Left, grdLib.Top

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim ilPass As Integer
    
    grdLib.Width = EngrLibDef.Width - 2 * grdLib.Left
    grdLib.ColWidth(CODEINDEX) = 0
    grdLib.ColWidth(USEDFLAGINDEX) = 0
    grdLib.ColWidth(NAMEINDEX) = grdLib.Width / 10
    grdLib.ColWidth(SUBLIBNAMEINDEX) = grdLib.Width / 10
    grdLib.ColWidth(STARTDATEINDEX) = grdLib.Width / 18
    grdLib.ColWidth(ENDDATEINDEX) = grdLib.Width / 18
    grdLib.ColWidth(DAYSINDEX) = grdLib.Width / 14
    'grdLib.ColWidth(STARTTIMEINDEX) = grdLib.Width / 18
    'grdLib.ColWidth(LENGTHINDEX) = grdLib.Width / 18
    grdLib.ColWidth(HOURSINDEX) = grdLib.Width / 10
    grdLib.ColWidth(BUSGROUPSINDEX) = grdLib.Width / 10
    grdLib.ColWidth(BUSESINDEX) = grdLib.Width / 10
    grdLib.ColWidth(CONFLICTINDEX) = grdLib.Width / 20
    grdLib.ColWidth(STATEINDEX) = grdLib.Width / 20
    grdLib.ColWidth(DESCRIPTIONINDEX) = grdLib.Width
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdLib.ColWidth(DESCRIPTIONINDEX) > grdLib.ColWidth(ilCol) Then
                grdLib.ColWidth(DESCRIPTIONINDEX) = grdLib.ColWidth(DESCRIPTIONINDEX) - grdLib.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    
    grdLibEvents.Width = grdLib.Width
    grdLibEvents.ColWidth(PCODEINDEX) = 0
    grdLibEvents.ColWidth(SORTTIMEINDEX) = 0
    grdLibEvents.ColWidth(ERRORFLAGINDEX) = 0
    grdLibEvents.ColWidth(CHGSTATUSINDEX) = 0
    grdLibEvents.ColWidth(EVTCONFLICTINDEX) = 0
    imUnusedCount = 0
    fmUsedWidth = 0
    fmUnusedWidth = 0
    grdLibEvents.ColWidth(HIGHLIGHTINDEX) = (3 * pbcHighlight.TextWidth("»")) / 2
    For ilPass = 0 To 1 Step 1
'        grdLibEvents.ColWidth(BUSNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSNAMEINDEX), 25, tgUsedSumEPE.sBus)
'        grdLibEvents.ColWidth(BUSCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSCTRLINDEX), 57, tgUsedSumEPE.sBusControl)
'        grdLibEvents.ColWidth(TIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TIMEINDEX), 17, tgUsedSumEPE.sTime)  '21
'        grdLibEvents.ColWidth(STARTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(STARTTYPEINDEX), 31, tgUsedSumEPE.sStartType)   '27
'        grdLibEvents.ColWidth(FIXEDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FIXEDINDEX), 38, tgUsedSumEPE.sFixedTime)
'        grdLibEvents.ColWidth(ENDTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(ENDTYPEINDEX), 31, tgUsedSumEPE.sEndType) '27
'        grdLibEvents.ColWidth(DURATIONINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(DURATIONINDEX), 17, tgUsedSumEPE.sDuration)  '25
'        grdLibEvents.ColWidth(AIRHOURSINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AIRHOURSINDEX), 25, "Y")
'        grdLibEvents.ColWidth(AIRDAYSINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AIRDAYSINDEX), 25, "Y")
'        grdLibEvents.ColWidth(MATERIALINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(MATERIALINDEX), 29, tgUsedSumEPE.sMaterialType)
'        grdLibEvents.ColWidth(AUDIONAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIONAMEINDEX), 23, tgUsedSumEPE.sAudioName)
'        grdLibEvents.ColWidth(AUDIOITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOITEMIDINDEX), 24, tgUsedSumEPE.sAudioItemID)
'        grdLibEvents.ColWidth(AUDIOCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOCTRLINDEX), 58, tgUsedSumEPE.sAudioControl)
'        grdLibEvents.ColWidth(BACKUPNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPNAMEINDEX), 23, tgUsedSumEPE.sBkupAudioName)
'        grdLibEvents.ColWidth(BACKUPCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPCTRLINDEX), 58, tgUsedSumEPE.sBkupAudioControl)
'        grdLibEvents.ColWidth(PROTNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTNAMEINDEX), 23, tgUsedSumEPE.sProtAudioName)
'        grdLibEvents.ColWidth(PROTITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTITEMIDINDEX), 24, tgUsedSumEPE.sProtAudioItemID)
'        grdLibEvents.ColWidth(PROTCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTCTRLINDEX), 58, tgUsedSumEPE.sProtAudioControl)
'        grdLibEvents.ColWidth(RELAY1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY1INDEX), 30, tgUsedSumEPE.sRelay1)
'        grdLibEvents.ColWidth(RELAY2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY2INDEX), 30, tgUsedSumEPE.sRelay2)
'        grdLibEvents.ColWidth(FOLLOWINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FOLLOWINDEX), 35, tgUsedSumEPE.sFollow)
'        grdLibEvents.ColWidth(SILENCETIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCETIMEINDEX), 25, tgUsedSumEPE.sSilenceTime)
'        grdLibEvents.ColWidth(SILENCE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE1INDEX), 58, tgUsedSumEPE.sSilence1)
'        grdLibEvents.ColWidth(SILENCE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE2INDEX), 58, tgUsedSumEPE.sSilence2)
'        grdLibEvents.ColWidth(SILENCE3INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE3INDEX), 58, tgUsedSumEPE.sSilence3)
'        grdLibEvents.ColWidth(SILENCE4INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE4INDEX), 58, tgUsedSumEPE.sSilence4)
'        grdLibEvents.ColWidth(NETCUE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE1INDEX), 31, tgUsedSumEPE.sStartNetcue)
'        grdLibEvents.ColWidth(NETCUE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE2INDEX), 31, tgUsedSumEPE.sStopNetcue)
'        grdLibEvents.ColWidth(TITLE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE1INDEX), 53, tgUsedSumEPE.sTitle1)
'        grdLibEvents.ColWidth(TITLE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE2INDEX), 53, tgUsedSumEPE.sTitle2)
        
        grdLibEvents.ColWidth(EVENTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(EVENTTYPEINDEX), 32, tgUsedSumEPE.sBus)
        grdLibEvents.ColWidth(BUSNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSNAMEINDEX), 32, tgUsedSumEPE.sBus)
        grdLibEvents.ColWidth(BUSCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSCTRLINDEX), 65, tgUsedSumEPE.sBusControl)
        grdLibEvents.ColWidth(TIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TIMEINDEX), 27, tgUsedSumEPE.sTime)  '30
        grdLibEvents.ColWidth(STARTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(STARTTYPEINDEX), 40, tgUsedSumEPE.sStartType)   '27
        grdLibEvents.ColWidth(FIXEDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FIXEDINDEX), 50, tgUsedSumEPE.sFixedTime)
        grdLibEvents.ColWidth(ENDTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(ENDTYPEINDEX), 40, tgUsedSumEPE.sEndType) '27
        grdLibEvents.ColWidth(DURATIONINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(DURATIONINDEX), 20, tgUsedSumEPE.sDuration)  '25
        grdLibEvents.ColWidth(AIRHOURSINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AIRHOURSINDEX), 25, "Y")
        grdLibEvents.ColWidth(AIRDAYSINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AIRDAYSINDEX), 25, "Y")
        grdLibEvents.ColWidth(MATERIALINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(MATERIALINDEX), 40, tgUsedSumEPE.sMaterialType)
        grdLibEvents.ColWidth(AUDIONAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIONAMEINDEX), 27, tgUsedSumEPE.sAudioName)
        grdLibEvents.ColWidth(AUDIOITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOITEMIDINDEX), 27, tgUsedSumEPE.sAudioItemID)
        grdLibEvents.ColWidth(AUDIOISCIINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOISCIINDEX), 40, tgUsedSumEPE.sAudioISCI)
        grdLibEvents.ColWidth(AUDIOCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOCTRLINDEX), 65, tgUsedSumEPE.sAudioControl)
        grdLibEvents.ColWidth(BACKUPNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPNAMEINDEX), 27, tgUsedSumEPE.sBkupAudioName)
        grdLibEvents.ColWidth(BACKUPCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPCTRLINDEX), 65, tgUsedSumEPE.sBkupAudioControl)
        grdLibEvents.ColWidth(PROTNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTNAMEINDEX), 27, tgUsedSumEPE.sProtAudioName)
        grdLibEvents.ColWidth(PROTITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTITEMIDINDEX), 27, tgUsedSumEPE.sProtAudioItemID)
        grdLibEvents.ColWidth(PROTISCIINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTISCIINDEX), 40, tgUsedSumEPE.sProtAudioISCI)
        grdLibEvents.ColWidth(PROTCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTCTRLINDEX), 65, tgUsedSumEPE.sProtAudioControl)
        grdLibEvents.ColWidth(RELAY1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY1INDEX), 50, tgUsedSumEPE.sRelay1)
        grdLibEvents.ColWidth(RELAY2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY2INDEX), 50, tgUsedSumEPE.sRelay2)
        grdLibEvents.ColWidth(FOLLOWINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FOLLOWINDEX), 40, tgUsedSumEPE.sFollow)
        grdLibEvents.ColWidth(SILENCETIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCETIMEINDEX), 40, tgUsedSumEPE.sSilenceTime)
        grdLibEvents.ColWidth(SILENCE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE1INDEX), 65, tgUsedSumEPE.sSilence1)
        grdLibEvents.ColWidth(SILENCE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE2INDEX), 65, tgUsedSumEPE.sSilence2)
        grdLibEvents.ColWidth(SILENCE3INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE3INDEX), 65, tgUsedSumEPE.sSilence3)
        grdLibEvents.ColWidth(SILENCE4INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE4INDEX), 65, tgUsedSumEPE.sSilence4)
        grdLibEvents.ColWidth(NETCUE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE1INDEX), 40, tgUsedSumEPE.sStartNetcue)
        grdLibEvents.ColWidth(NETCUE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE2INDEX), 40, tgUsedSumEPE.sStopNetcue)
'        grdLibEvents.ColWidth(TITLE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE1INDEX), 53, tgUsedSumEPE.sTitle1)
'        grdLibEvents.ColWidth(TITLE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE2INDEX), 53, tgUsedSumEPE.sTitle2)
        If sgClientFields = "A" Then
            If tgUsedSumEPE.sABCFormat <> "Y" Then
                grdLibEvents.ColWidth(ABCFORMATINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCFORMATINDEX) = grdLibEvents.Width / 28
            End If
            If tgUsedSumEPE.sABCPgmCode <> "Y" Then
                grdLibEvents.ColWidth(ABCPGMCODEINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCPGMCODEINDEX) = grdLibEvents.Width / 28
            End If
            If tgUsedSumEPE.sABCXDSMode <> "Y" Then
                grdLibEvents.ColWidth(ABCXDSMODEINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCXDSMODEINDEX) = grdLibEvents.Width / 28
            End If
            If tgUsedSumEPE.sABCRecordItem <> "Y" Then
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
    For ilCol = HIGHLIGHTINDEX To TITLE2INDEX Step 1
        If ilCol <> TITLE1INDEX And ilCol <> TITLE2INDEX Then
            If grdLibEvents.ColWidth(TITLE1INDEX) > grdLibEvents.ColWidth(ilCol) Then
                grdLibEvents.ColWidth(TITLE1INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) - grdLibEvents.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    grdLibEvents.ColWidth(TITLE2INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) / 8
    grdLibEvents.ColWidth(TITLE1INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) - grdLibEvents.ColWidth(TITLE2INDEX)
    '8/26/11: Move here
    gGrid_IntegralHeight grdLibEvents


    grdConflicts.Width = EngrLibDef.Width - 2 * grdConflicts.Left
    imUnusedCount = 0
    fmUsedWidth = 0
    fmUnusedWidth = 0
    For ilPass = 0 To 1 Step 1
        'grdConflicts.ColWidth(CONFLICTNAMEINDEX) = grdConflicts.Width / 10
        'grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX), 10, tgUsedSumEPE.sBus)
        grdConflicts.ColWidth(CONFLICTSTARTDATEINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTSTARTDATEINDEX), 18, "Y")
        grdConflicts.ColWidth(CONFLICTENDDATEINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTENDDATEINDEX), 18, "Y")
        grdConflicts.ColWidth(CONFLICTDAYSINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTDAYSINDEX), 14, "Y")
        grdConflicts.ColWidth(CONFLICTOFFSETINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTOFFSETINDEX), 12, tgUsedSumEPE.sTime)
        grdConflicts.ColWidth(CONFLICTHOURSINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTHOURSINDEX), 12, "Y")
        grdConflicts.ColWidth(CONFLICTDURATIONINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTDURATIONINDEX), 12, tgUsedSumEPE.sDuration)
        grdConflicts.ColWidth(CONFLICTBUSESINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTBUSESINDEX), 10, tgUsedSumEPE.sBus)
        grdConflicts.ColWidth(CONFLICTAUDIOINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTAUDIOINDEX), 20, tgUsedSumEPE.sAudioName)
        grdConflicts.ColWidth(CONFLICTBACKUPINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTBACKUPINDEX), 20, tgUsedSumEPE.sBkupAudioName)
        grdConflicts.ColWidth(CONFLICTPROTINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTPROTINDEX), 20, tgUsedSumEPE.sProtAudioName)
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
    gGrid_Clear grdLib, True
    gGrid_Clear grdLibEvents, True
    'Can't be 0 to 0 because index stored into grid
    smState = ""
    smConflict = ""
    Dim lmDeleteCodes(0 To 0) As Long
    ReDim tmCurrDEE(1 To 1) As DEE
    imFieldChgd = False
    imLimboAllowed = False
    ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
    ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
End Sub
Private Sub mMoveCtrlsToRec()
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilDays As Integer
    Dim ilHours As Integer
    Dim ilSet As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilPos As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim ilSHour As Integer
    Dim ilEHour As Integer
    
'    smNowDate = Format(gNow(), sgShowDateForm)
'    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    llRow = grdLib.FixedRows
    If Trim$(grdLib.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdLib.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmDHE.lCode = Val(grdLib.TextMatrix(llRow, CODEINDEX))
    tmDHE.lDneCode = 0
    tmDHE.sType = "L"
    slStr = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
    For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
        If StrComp(Trim$(tgCurrLibDNE(ilLoop).sName), slStr, vbTextCompare) = 0 Then
            tmDHE.lDneCode = tgCurrLibDNE(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    tmDHE.lDseCode = 0
    slStr = Trim$(grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX))
    For ilLoop = 0 To UBound(tgCurrDSE) - 1 Step 1
        If StrComp(Trim$(tgCurrDSE(ilLoop).sName), slStr, vbTextCompare) = 0 Then
            tmDHE.lDseCode = tgCurrDSE(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    tmDHE.lCteCode = 0  'Set later
    smDHEComment = Trim$(grdLib.TextMatrix(llRow, DESCRIPTIONINDEX))
    tmDHE.sStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
    If Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX)) <> "" Then
        tmDHE.sEndDate = Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX))
    Else
        tmDHE.sEndDate = "12/31/2069"
    End If
    slStr = Trim$(grdLib.TextMatrix(llRow, DAYSINDEX))
    tmDHE.sDays = gCreateDayStr(slStr)
'    tmDHE.sStartTime = gFormatTime(grdLib.TextMatrix(llRow, STARTTIMEINDEX))
'    tmDHE.lLength = gLengthToLong(grdLib.TextMatrix(llRow, LENGTHINDEX))
    slStr = Trim$(grdLib.TextMatrix(llRow, HOURSINDEX))
    tmDHE.sHours = gCreateHourStr(slStr)
    slStr = tmDHE.sHours
    For ilSHour = 0 To 23 Step 1
        If Mid$(slStr, ilSHour + 1, 1) = "Y" Then
            If ilSHour <= 9 Then
                tmDHE.sStartTime = "0" & Trim$(Str(ilSHour)) & ":00:00"
            Else
                tmDHE.sStartTime = Trim$(Str(ilSHour)) & ":00:00"
            End If
            For ilEHour = 23 To 0 Step -1
                If Mid$(slStr, ilEHour + 1, 1) = "Y" Then
                    ilHours = ilEHour - ilSHour + 1
                    tmDHE.lLength = CLng(3600) * ilHours
                    Exit For
                End If
            Next ilEHour
            Exit For
        End If
    Next ilSHour
    'Bus Groups
    smDHEBusGroups = grdLib.TextMatrix(llRow, BUSGROUPSINDEX)
    'Buses
    smDHEBuses = grdLib.TextMatrix(llRow, BUSESINDEX)
    tmDHE.sBusNames = smDHEBuses
    If grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Both" Then
        tmDHE.sIgnoreConflicts = "I"
    ElseIf grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Bus" Then
        tmDHE.sIgnoreConflicts = "B"
    ElseIf grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Audio" Then
        tmDHE.sIgnoreConflicts = "A"
    Else
        tmDHE.sIgnoreConflicts = "N"
    End If
    If grdLib.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmDHE.sState = "D"
    ElseIf grdLib.TextMatrix(llRow, STATEINDEX) = "Limbo" Then
        tmDHE.sState = "L"
    Else
        tmDHE.sState = "A"
    End If
    If tmDHE.lCode <= 0 Then
        tmDHE.sUsedFlag = "N"
    Else
        tmDHE.sUsedFlag = grdLib.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmDHE.iVersion = 0
    tmDHE.lOrigDHECode = tmDHE.lCode
    tmDHE.sCurrent = "Y"
    tmDHE.sEnteredDate = Format(Now, sgShowDateForm) 'smNowDate
    tmDHE.sEnteredTime = Format(Now, sgShowTimeWSecForm) 'smNowTime
    tmDHE.iUieCode = tgUIE.iCode
    tmDHE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slHours As String
    Dim ilDBE As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    Dim llRet As Long
    Dim llListIndex As Long
    
    
    'gGrid_Clear grdLib, True
    llRow = grdLib.FixedRows
    For ilDNE = 0 To UBound(tgCurrLibDNE) - 1 Step 1
        If tmDHE.lDneCode = tgCurrLibDNE(ilDNE).lCode Then
            grdLib.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrLibDNE(ilDNE).sName)
            llRet = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrLibDNE(ilDNE).sName))
            If llRet < 0 Then
                lbcDNE.AddItem Trim$(tgCurrLibDNE(ilDNE).sName)
                lbcDNE.ItemData(lbcDNE.NewIndex) = tgCurrLibDNE(ilDNE).lCode
            End If
            Exit For
        End If
    Next ilDNE
    For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
        If tmDHE.lDseCode = tgCurrDSE(ilDSE).lCode Then
            grdLib.TextMatrix(llRow, SUBLIBNAMEINDEX) = Trim$(tgCurrDSE(ilDSE).sName)
            llRet = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrDSE(ilDSE).sName))
            If llRet < 0 Then
                lbcDSE.AddItem Trim$(tgCurrDSE(ilDSE).sName)
                lbcDSE.ItemData(lbcDSE.NewIndex) = tgCurrDSE(ilDSE).lCode
            End If
            Exit For
        End If
    Next ilDSE
    ilRet = gGetRec_CTE_CommtsTitle(tmDHE.lCteCode, "EngrLibDef- mMoveRecToCtrl for CTE", tmCTE)
    grdLib.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tmCTE.sComment)
    grdLib.TextMatrix(llRow, STARTDATEINDEX) = Trim$(tmDHE.sStartDate)
    If gDateValue(Trim$(tmDHE.sEndDate)) <> gDateValue("12/31/2069") Then
        grdLib.TextMatrix(llRow, ENDDATEINDEX) = Trim$(tmDHE.sEndDate)
    Else
        grdLib.TextMatrix(llRow, ENDDATEINDEX) = ""
    End If
    slStr = gDayMap(tmDHE.sDays)
    grdLib.TextMatrix(llRow, DAYSINDEX) = Trim$(slStr)
    'grdLib.TextMatrix(llRow, STARTTIMEINDEX) = Trim$(tmDHE.sStartTime)
    'grdLib.TextMatrix(llRow, LENGTHINDEX) = gLongToLength(tmDHE.lLength)
    slHours = Trim$(tmDHE.sHours)
    slStr = gHourMap(slHours)
    grdLib.TextMatrix(llRow, HOURSINDEX) = slStr
    
    slStr = ""
    smCurrDBEStamp = ""
    Erase tmCurrDBE
    ilRet = gGetRecs_DBE_DayBusSel(smCurrDBEStamp, tmDHE.lCode, "Bus Definition-mMoveRecToCtrls", tmCurrDBE())
    For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
        If tmCurrDBE(ilDBE).sType = "G" Then
            For ilBGE = 0 To UBound(tgCurrBGE) - 1 Step 1
                If tmCurrDBE(ilDBE).iBgeCode = tgCurrBGE(ilBGE).iCode Then
                    slStr = slStr & Trim$(tgCurrBGE(ilBGE).sName) & ","
                    llRet = SendMessageByString(lbcBGE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrBGE(ilBGE).sName))
                    If llRet < 0 Then
                        lbcBGE.AddItem Trim$(tgCurrBGE(ilBGE).sName)
                        lbcBGE.ItemData(lbcBGE.NewIndex) = tgCurrBGE(ilBGE).iCode
                    End If
                    Exit For
                End If
            Next ilBGE
        End If
    Next ilDBE
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    grdLib.TextMatrix(llRow, BUSGROUPSINDEX) = slStr

    slStr = ""
    For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
        If tmCurrDBE(ilDBE).sType = "B" Then
            'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '    If tmCurrDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                ilBDE = gBinarySearchBDE(tmCurrDBE(ilDBE).iBdeCode, tgCurrBDE())
                If ilBDE <> -1 Then
                    slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName) & ","
                    llRet = SendMessageByString(lbcBDE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrBDE(ilBDE).sName))
                    If llRet < 0 Then
                        lbcBDE.AddItem Trim$(tgCurrBDE(ilBDE).sName)
                        lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilBDE).iCode
                    End If
            '        Exit For
                End If
            'Next ilBDE
        End If
    Next ilDBE
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    grdLib.TextMatrix(llRow, BUSESINDEX) = slStr
    ''Set selection to handle the case when tab passed group to get original selected values
    'gParseCDFields slStr, False, smBuses()
    'lbcBDE.ListIndex = -1
    'For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
    '    lbcBDE.Selected(ilLoop) = False
    'Next ilLoop
    'For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
    '    slStr = Trim$(smBuses(ilLoop))
    '    If slStr <> "" Then
    '        'llRow = SendMessageByString(lbcBGE.hwnd, LB_FINDSTRING, -1, slStr)
    '        llListIndex = gListBoxFind(lbcBDE, slStr)
    '        If llListIndex >= 0 Then
    '            lbcBDE.Selected(llListIndex) = True
    '        End If
    '    End If
    'Next ilLoop
    
    If tmDHE.sIgnoreConflicts = "I" Then
        grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Both"
    ElseIf tmDHE.sIgnoreConflicts = "B" Then
        grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Bus"
    ElseIf tmDHE.sIgnoreConflicts = "A" Then
        grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Audio"
    Else
        grdLib.TextMatrix(llRow, CONFLICTINDEX) = "No"
    End If
    If tmDHE.sState = "A" Then
        grdLib.TextMatrix(llRow, STATEINDEX) = "Active"
    ElseIf tmDHE.sState = "L" Then
        grdLib.TextMatrix(llRow, STATEINDEX) = "Limbo"
    Else
        grdLib.TextMatrix(llRow, STATEINDEX) = "Dormant"
    End If
    grdLib.TextMatrix(llRow, CODEINDEX) = tmDHE.lCode
    grdLib.TextMatrix(llRow, USEDFLAGINDEX) = tmDHE.sUsedFlag
    
    
    grdLib.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetRec_DHE_DayHeaderInfoAPI(hmDHE, lgLibCallCode, "EngrLibDef-mPopulation", tmDHE)
    ilRet = gGetRecs_DEE_DayEventAPI(hmDEE, sgCurrDEEStamp, lgLibCallCode, "EngrLibDef-mPopulate", tgCurrDEE())
    If lgLibCallCode <= 0 Then
        tmDHE.lCode = 0
    End If
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim ilNew As Integer
    Dim llOldDEECode As Long
    Dim ilDEECompare As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llNewAgedDHECode As Long
    Dim ilNameOk As Integer
    Dim slDHEStartDate As String
    Dim tlDHE As DHE
    Dim tlSvDHE As DHE
    Dim ilSpotRomoved As Integer
    Dim ilCTE As Integer
    
    bmInSave = True
    gSetMousePointer grdLib, grdLibEvents, vbHourglass
    
    'ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    mPopATE
    gConflictPop
    If Not mCheckFields(True) Then
        gSetMousePointer grdLib, grdLibEvents, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Library Save"
        mSave = False
        Exit Function
    End If
    '9/6/11: Test in mNameOk
    'If (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) <> "Dormant") And (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) <> "Limbo") Then
        If Not mNameOk() Then
            gSetMousePointer grdLib, grdLibEvents, vbDefault
            mSave = False
            Exit Function
        End If
    'End If
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    mInitConflictTest
    '10/9/09:  Remove conflict checking and hide cmcConflict button.
    '          Conflict checking only performed when schedule created and on the schedule screen
    '          The Conflict table does not break out which audio error so more code would have to be added
    '          to bypass site option test
    'If (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) <> "Dormant") And (grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) <> "Limbo") Then
    '    If mCheckLibConflicts() Then
    '        gSetMousePointer grdLib, grdLibEvents, vbDefault
    '        MsgBox "Dates/Days/Times/Buses/Audio in Conflict with other Libraries", vbCritical + vbOKOnly, "Library Save"
    '        mSave = False
    '        Exit Function
    '    End If
    '    If mCheckEventConflicts() Then
    '        gSetMousePointer grdLib, grdLibEvents, vbDefault
    '        MsgBox "Dates/Days/Times/Buses/Audio in Conflict within this Library", vbCritical + vbOKOnly, "Library Save"
    '        mSave = False
    '        Exit Function
    '    End If
    'End If
    'smOverlapMsg set in mNameOk
    ilRet = MsgBox(smOverlapMsg, vbOKCancel + vbQuestion, "Library Add")
    If ilRet = vbCancel Then
        gSetMousePointer grdLib, grdLibEvents, vbDefault
        mSave = False
        Exit Function
    End If
    llRow = grdLib.FixedRows
    If imGapExist Then
        slStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
        slEndDate = DateAdd("d", -1, slStartDate)
        gSetMousePointer grdLib, grdLibEvents, vbDefault
        ilRet = MsgBox("Expand Library ending on " & Format(lmGapEndDate, "ddddd") & " to end on " & slEndDate, vbYesNoCancel + vbQuestion, "Library Change")
        If ilRet = vbCancel Then
            mSave = False
            Exit Function
        End If
        gSetMousePointer grdLib, grdLibEvents, vbHourglass
        If ilRet = vbNo Then
            imGapExist = False
        End If
    End If
    grdLib.Redraw = False
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    slStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
    If Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX)) <> "" Then
        slEndDate = Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX))
    Else
        slEndDate = "12/31/2069"
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    tmSchdChgInfo.lNewChgDHE = tmDHE.lCode
    tmSchdChgInfo.lCheckDHE = 0
    tmSchdChgInfo.lSplitDHE = 0
    tmSchdChgInfo.lExpandDHE = 0
    tmSchdChgInfo.lDEEDHE = 0
    If (imOverlapCase > 0) And (lmCurrentDHE > 0) Then
        tmSchdChgInfo.lCheckDHE = lmCurrentDHE
        llRow = grdLib.FixedRows
        ilRet = gGetRec_DHE_DayHeaderInfo(lmCurrentDHE, "EngrLibDef-mPopulation", tlDHE)
        'tlDHE.sEnteredDate = smNowDate
        'tlDHE.sEnteredTime = smNowTime
        tlDHE.sEnteredDate = Format(Now, sgShowDateForm)
        tlDHE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tlDHE.iVersion = gGetLatestVersion_DHE(tlDHE.lOrigDHECode, "EngrLibDef-mSave: Get Latest Version") + 1
        tlDHE.iUieCode = tgUIE.iCode
        If imOverlapCase = 1 Then   'Replace
            If tlDHE.sUsedFlag = "N" Then
                ilRet = gPutDelete_DHE_DayHeaderInfo(tlDHE.lCode, "Library Definition-mSave: Delete DHE")
            Else
                '5/31/11: Disallow changes in the schedule area
                'tlDHE.sState = "D"
                'ilRet = gPutUpdate_Library(tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
                'ilRet = gPutDelete_CME_Conflict_Master("L", tlDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
                'ilRet = gPutUpdate_DHE_DayHeaderInfo(4, tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
                ilRet = gPutDelete_DHE_DayHeaderInfo(tlDHE.lCode, "Library Definition-mSave: Delete DHE")
            End If
        ElseIf imOverlapCase = 2 Then   'Terminate Current
            tlDHE.sEndDate = DateAdd("d", -1, slStartDate)
            '5/31/11: Disallow changes in the schedule area
            'ilRet = gPutUpdate_Library(tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
            'ilRet = gPutDelete_CME_Conflict_Master("L", tlDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
            'ilRet = gCreateCMEForLib(tlDHE, tlDHE.sStartDate, hmCME)
            ilRet = gPutUpdate_DHE_DayHeaderInfo(5, tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
        ElseIf imOverlapCase = 3 Then   'Change start date to be after new end date
            '5/31/11: Disallow changes in the schedule area
            'tlDHE.sStartDate = DateAdd("d", 1, slEndDate)
            'ilRet = gPutUpdate_Library(tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
            'ilRet = gPutDelete_CME_Conflict_Master("L", tlDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
            'ilRet = gCreateCMEForLib(tlDHE, tlDHE.sStartDate, hmCME)
            mSave = False
            Exit Function
        ElseIf imOverlapCase = 4 Then   'Split
            '5/31/11: Disallow changes in the schedule area
            'LSet tlSvDHE = tlDHE
            'tlDHE.sEndDate = DateAdd("d", -1, slStartDate)
            'ilRet = gPutUpdate_Library(tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
            'ilRet = gPutDelete_CME_Conflict_Master("L", tlDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
            'ilRet = gCreateCMEForLib(tlDHE, tlDHE.sStartDate, hmCME)
            'tmSchdChgInfo.lSplitDHE = llNewAgedDHECode
            'tlSvDHE.lCode = llNewAgedDHECode
            'tlSvDHE.sStartDate = DateAdd("d", 1, slEndDate)
            'tlSvDHE.iVersion = gGetLatestVersion_DHE(tlDHE.lOrigDHECode, "EngrLibDef-mSave: Get Latest Version") + 1
            'ilRet = gPutUpdate_Library(tlSvDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
            'ilRet = gCreateCMEForLib(tlSvDHE, tlSvDHE.sStartDate, hmCME)
            mSave = False
            Exit Function
        End If
    End If
    'If (imGapExist) And (lmGapDHE > 0) Then
    '    tmSchdChgInfo.lExpandDHE = lmGapDHE
    '    llRow = grdLib.FixedRows
    '    slStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
    '    If Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX)) <> "" Then
    '        slEndDate = Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX))
    '    Else
    '        slEndDate = "12/31/2069"
    '    End If
    '    llStartDate = gDateValue(slStartDate)
    '    llEndDate = gDateValue(slEndDate)
    '    ilRet = gGetRec_DHE_DayHeaderInfo(lmGapDHE, "EngrLibDef-mPopulation", tlDHE)
    '    'tlDHE.sEnteredDate = smNowDate
    '    'tlDHE.sEnteredTime = smNowTime
    '    tlDHE.sEnteredDate = Format(Now, sgShowDateForm)
    '    tlDHE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    '    tlDHE.iVersion = gGetLatestVersion_DHE(tlDHE.lOrigDHECode, "EngrLibDef-mSave: Get Latest Version") + 1
    '    tlDHE.iUieCode = tgUIE.iCode
    '    tlDHE.sEndDate = DateAdd("d", -1, slStartDate)
    '    ilRet = gPutUpdate_Library(tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
    'End If
    
    mMoveCtrlsToRec
    mMoveDEECtrlsToRec
    
    If Trim$(smDHEComment) <> "" Then
        mSetCTE smDHEComment, "DH"
        ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Library Definition-mSave: Insert CTE", hmCTE)
        If ilRet Then
            tmDHE.lCteCode = tmCTE.lCode
        Else
            tmDHE.lCteCode = 0
        End If
    Else
        tmDHE.lCteCode = 0
    End If
    If tmDHE.lCode <= 0 Then
        ilNew = True
        ilRet = gPutInsert_DHE_DayHeaderInfo(0, tmDHE, "Library Definition-mSave: DHE")
    Else
        '5/31/11: Disallow changes in the schedule area
        'ilNew = False
        'ilRet = gGetRec_DHE_DayHeaderInfo(lgLibCallCode, "EngrLibDef-mPopulation", tlDHE)
        'If ilRet Then
        '    tmDHE.iVersion = gGetLatestVersion_DHE(tlDHE.lOrigDHECode, "EngrLibDef-mSave: Get Latest Version") + 1
        '    If mCompare(tmDHE, tlDHE) Then
        '        ilRet = gPutUpdate_DHE_DayHeaderInfo(0, tmDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
        '    Else
        '        ilRet = gPutUpdate_DHE_DayHeaderInfo(1, tmDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
        '    End If
        '    '5/31/11: Disallow changes in the schedule area
        '    'If tmSchdChgInfo.lCheckDHE = 0 Then
        '    '    tmSchdChgInfo.lDEEDHE = tmDHE.lCode 'llNewAgedDHECode
        '    '    tmSchdChgInfo.lCheckDHE = llNewAgedDHECode
        '    'End If
        'End If
        ilNew = True
        ilRet = gGetRec_DHE_DayHeaderInfo(lgLibCallCode, "EngrLibDef-mPopulation", tlDHE)
        If ilRet Then
            tmDHE.iVersion = gGetLatestVersion_DHE(tlDHE.lOrigDHECode, "EngrLibDef-mSave: Get Latest Version") + 1
            If tlDHE.sUsedFlag = "N" Then
                ilRet = gPutDelete_DHE_DayHeaderInfo(tlDHE.lCode, "Library Definition-mSave: Delete DHE")
                '1/8/12: Retain as comments are not removed when library deleted in case comments in schedule
                'For ilCTE = 0 To UBound(tmCurr1CTE_Name) - 1 Step 1
                '    tmCurr1CTE_Name(ilCTE).lCteCode = 0
                'Next ilCTE
                'For ilCTE = 0 To UBound(tmCurr2CTE_Name) - 1 Step 1
                '    tmCurr2CTE_Name(ilCTE).lCteCode = 0
                'Next ilCTE
            Else
                'This case should not happen any longer as start date must be after latest scheduled date
                tlDHE.sState = "D"
                tlDHE.sCurrent = "N"
                ilRet = gPutUpdate_DHE_DayHeaderInfo(4, tlDHE, "Library Definition-mSave: Update DHE", llNewAgedDHECode)
            End If
            'ilRet = gPutDelete_CME_Conflict_Master("L", tlDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
            tmDHE.lCode = 0
            ilRet = gPutInsert_DHE_DayHeaderInfo(0, tmDHE, "Library Definition-mSave: DHE")
        End If
    End If
    '5/31/11: Disallow changes in the schedule area
    'If tmSchdChgInfo.lNewChgDHE = 0 Then
    '    tmSchdChgInfo.lNewChgDHE = tmDHE.lCode
    '    If (tmSchdChgInfo.lCheckDHE <> 0) Then
    '        tmSchdChgInfo.lDEEDHE = tmDHE.lCode
    '    End If
    'End If
    gParseCDFields smDHEBusGroups, False, smBusGroups()
    For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
        mSetDBE Trim$(smBusGroups(ilLoop)), "G"
        ilRet = gPutInsert_DBE_DayBusSel(tmDBE, "Library Definition-mSave: DBE")
        ilRet = gPutUpdate_BGE_UsedFlag(tmDBE.iBgeCode, tgCurrBGE())
    Next ilLoop
    gParseCDFields smDHEBuses, False, smBuses()
    For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
        mSetDBE Trim$(smBuses(ilLoop)), "B"
        ilRet = gPutInsert_DBE_DayBusSel(tmDBE, "Library Definition-mSave: DBE")
        ilRet = gPutUpdate_BDE_UsedFlag(tmDBE.iBdeCode, tgCurrBDE())
    Next ilLoop
    For llRow = LBound(tmCurrDEE) To UBound(tmCurrDEE) - 1 Step 1
        llOldDEECode = tmCurrDEE(llRow).lCode
        'If tmCurrDEE(llRow).lCode > 0 Then
        '    ilDEECompare = mCompareDEE(tmCurrDEE(llRow).lCode, smEBuses(llRow), Trim$(smT1Comment(llRow)))
        'End If
        tmCurrDEE(llRow).l1CteCode = 0
        If Trim$(smT1Comment(llRow)) <> "" Then
            For ilCTE = 0 To UBound(tmCurr1CTE_Name) - 1 Step 1
                If StrComp(UCase(Trim$(tmCurr1CTE_Name(ilCTE).sComment)), UCase(Trim$(smT1Comment(llRow))), vbBinaryCompare) = 0 Then
                    tmCurrDEE(llRow).l1CteCode = tmCurr1CTE_Name(ilCTE).lCteCode
                    Exit For
                End If
            Next ilCTE
            If tmCurrDEE(llRow).l1CteCode = 0 Then
                mSetCTE smT1Comment(llRow), "T1"
                ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Library Definition-mSave: Insert CTE", hmCTE)
                If ilRet Then
                    tmCurrDEE(llRow).l1CteCode = tmCTE.lCode
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).sComment = tmCTE.sComment
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lCteCode = tmCTE.lCode
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDheCode = tmDHE.lCode
                    ReDim Preserve tmCurr1CTE_Name(0 To UBound(tmCurr1CTE_Name) + 1) As DEECTE
                End If
            End If
        End If
        '7/8/11: Change T2 to work like T1
        tmCurrDEE(llRow).l2CteCode = 0
        If Trim$(smT2Comment(llRow)) <> "" Then
            For ilCTE = 0 To UBound(tmCurr2CTE_Name) - 1 Step 1
                If StrComp(UCase(Trim$(tmCurr2CTE_Name(ilCTE).sComment)), UCase(Trim$(smT2Comment(llRow))), vbBinaryCompare) = 0 Then
                    tmCurrDEE(llRow).l2CteCode = tmCurr2CTE_Name(ilCTE).lCteCode
                    Exit For
                End If
            Next ilCTE
            If tmCurrDEE(llRow).l2CteCode = 0 Then
                mSetCTE smT2Comment(llRow), "T2"
                ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Library Definition-mSave: Insert CTE", hmCTE)
                If ilRet Then
                    tmCurrDEE(llRow).l2CteCode = tmCTE.lCode
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).sComment = tmCTE.sComment
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lCteCode = tmCTE.lCode
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDheCode = tmDHE.lCode
                    ReDim Preserve tmCurr2CTE_Name(0 To UBound(tmCurr2CTE_Name) + 1) As DEECTE
                End If
            End If
        End If
        
        
        tmCurrDEE(llRow).lCode = 0
        tmCurrDEE(llRow).lDheCode = tmDHE.lCode
        ilRet = gPutInsert_DEE_DayEvent(tmCurrDEE(llRow), "Library Definition-mSave: DEE")
        'If llOldDEECode > 0 Then
        '    If Not ilDEECompare Then
        '        ilRet = gUpdateAIE(1, tmDHE.iVersion, "DEE", llOldDEECode, tmCurrDEE(llRow).lCode, tmDHE.lOrigDHECode, "Library Definition- mSave: Insert DEE:AIE")
        '        mSetUsedFlags tmCurrDEE(llRow)
        '    End If
        'Else
            mSetUsedFlags tmCurrDEE(llRow)
        'End If
        gParseCDFields smEBuses(llRow), False, smBuses()
        For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
            mSetEBE Trim$(smBuses(ilLoop)), tmCurrDEE(llRow).lCode
            ilRet = gPutInsert_EBE_EventBusSel(tmEBE, "Library Definition-mSave: EBE")
            ilRet = gPutUpdate_BDE_UsedFlag(tmEBE.iBdeCode, tgCurrBDE())
        Next ilLoop
    Next llRow
'    For ilLoop = LBound(lmDeleteCodes) To UBound(lmDeleteCodes) - 1 Step 1
'        ilRet = gPutDelete_ETE_EventType(lmDeleteCodes(ilLoop), "EngrLibDef- Delete")
'    Next ilLoop
'    ReDim lmDeleteCodes(0 To 0) As Integer
'    grdLib.Redraw = True
'    sgCurrETEStamp = ""
'    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibDef-mSave", tgCurrETE())
'    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrLibDef-mSave", tgCurrDEE())
    '5/31/11: Disallow changes in the schedule area
    'ilRet = gPutDelete_CME_Conflict_Master("L", tmDHE.lCode, 0, 0, "Library Definition-mSave: Delete CME", hmCME)
    'slDHEStartDate = tmDHE.sStartDate
    'ilRet = gCreateCMEForLib(tmDHE, slDHEStartDate, hmCME)
    'ilRet = gAdjustSEE(tmSchdChgInfo, hmSEE, hmSOE, ilSpotRomoved)
    'If ilSpotRomoved = 1 Then
    '    MsgBox "Spots deleted from schedule as no matching avail found. The Load-Automation file will be re-created automatically without the deleted spots", vbInformation + vbOK
    'ElseIf ilSpotRomoved = 2 Then
    '    MsgBox "Spots removed from schedule as no matching avail found", vbInformation + vbOK
    'End If
    lmLibCallCode = 0
    imLibCallType = 0
    imFieldChgd = False
    mSetCommands
    sgCurrLibDHEStamp = ""
    mSave = True
End Function
Private Sub cmcCancel_Click()
    If bmInSave Then
        Exit Sub
    End If
    igReturnCallStatus = CALLCANCELLED
    Unload EngrLibDef
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If bmInSave Then
        Exit Sub
    End If
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrLibDef
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdLib, grdLibEvents, vbHourglass
        ilRet = mSave()
        bmInSave = False
        gSetMousePointer grdLib, grdLibEvents, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdLib, grdLibEvents, vbDefault
    Unload EngrLibDef
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilCol As Integer
    Dim llTopRow As Long
    
    If bmInSave Then
        Exit Sub
    End If
    If imFieldChgd = True Then
        gSetMousePointer grdLib, grdLibEvents, vbHourglass
        llTopRow = grdLibEvents.TopRow
        ilRet = mSave()
        bmInSave = False
        If Not ilRet Then
            grdLib.Redraw = True
            grdLibEvents.Redraw = True
            gSetMousePointer grdLib, grdLibEvents, vbDefault
            Exit Sub
        End If
        DoEvents
        grdLib.Redraw = False
        grdLibEvents.Redraw = False
        DoEvents
        mClearControls
        lgLibCallCode = tmDHE.lCode
        DoEvents
        mPopulate
        DoEvents
        grdLibEvents.Visible = False
        'mMove---- will reset Redraw to True
        mMoveRecToCtrls
        grdLibEvents.Redraw = False
        mMoveDEERecToCtrls
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
        grdLibEvents.TopRow = llTopRow
        mSetBackgroundColor
        lmEEnableRow = -1
        lmEEnableCol = -1
        lmEnableRow = -1
        lmEnableCol = -1
        imFieldChgd = False
        imLimboAllowed = False
        If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
            imLimboAllowed = True
        End If
        mSetCommands
        grdLibEvents.Visible = True
        grdLib.Redraw = True
        grdLibEvents.Redraw = True
        gSetMousePointer grdLib, grdLibEvents, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub


Private Sub dpcEvent_OnChange()
    If StrComp(Trim$(grdLibEvents.text), Trim$(dpcEvent.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        grdLibEvents.text = dpcEvent.text
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub dpcLib_OnChange()
    If StrComp(Trim$(grdLib.text), Trim$(dpcLib.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        grdLib.text = dpcLib.text
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    Dim ilANE As Integer
    Dim ilCCE As Integer
    Dim llCode As Long
    Dim ilLoop As Integer
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdLib.Col
        Case NAMEINDEX
            'llRow = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, 1, slStr)
            llRow = gListBoxFind(lbcDNE, slStr)
            If llRow >= 0 Then
                lbcDNE.ListIndex = llRow
                edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case SUBLIBNAMEINDEX
            llRow = gListBoxFind(lbcDSE, slStr)
            If llRow >= 0 Then
                lbcDSE.ListIndex = llRow
                edcDropdown.text = lbcDSE.List(lbcDSE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case BUSGROUPSINDEX
            llRow = gListBoxFind(lbcBGE, slStr)
            If llRow >= 0 Then
                lbcBGE.ListIndex = llRow
                edcDropdown.text = lbcBGE.List(lbcBGE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case BUSESINDEX
            llRow = gListBoxFind(lbcBDE, slStr)
            If llRow >= 0 Then
                lbcBDE.ListIndex = llRow
                edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If StrComp(Trim$(grdLib.text), Trim$(edcDropdown.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        Select Case grdLib.Col
            Case NAMEINDEX
                If lbcDNE.ListIndex > 0 Then
                    llCode = lbcDNE.ItemData(lbcDNE.ListIndex)
                    For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
                        If llCode = tgCurrLibDNE(ilLoop).lCode Then
                            grdLib.TextMatrix(grdLib.Row, DESCRIPTIONINDEX) = Trim$(tgCurrLibDNE(ilLoop).sDescription)
                            Exit For
                        End If
                    Next ilLoop
                Else
                    grdLib.TextMatrix(grdLib.Row, DESCRIPTIONINDEX) = ""
                End If
                mPopNNE
        End Select
        grdLib.text = edcDropdown.text
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands

End Sub

Private Sub edcDropdown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdLib.Col
            Case NAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcDNE, True
            Case SUBLIBNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcDSE, True
            Case BUSGROUPSINDEX
                gProcessArrowKey Shift, KeyCode, lbcBGE, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mBranch()
    End If
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
        Case BUSCTRLINDEX
            llRow = gListBoxFind(lbcCCE_B, slStr)
            If llRow >= 0 Then
                lbcCCE_B.ListIndex = llRow
                edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case EVENTTYPEINDEX
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
                slStr = Trim$(edcEDropdown.text)
                For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                                slStr = ""
                                If tmCurrASE(ilASE).iPriCceCode > 0 Then
                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                                        If tmCurrASE(ilASE).iPriCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                                            grdLibEvents.TextMatrix(grdLibEvents.Row, AUDIOCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                                            Exit For
                                        End If
                                    Next ilCCE
                                Else
                                    grdLibEvents.TextMatrix(grdLibEvents.Row, AUDIOCTRLINDEX) = ""
                                End If
                                If tmCurrASE(ilASE).iBkupAneCode > 0 Then
                                    'For ilANE2 = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tmCurrASE(ilASE).iBkupAneCode = tgCurrANE(ilANE2).iCode Then
                                        ilANE2 = gBinarySearchANE(tmCurrASE(ilASE).iBkupAneCode, tgCurrANE())
                                        If ilANE2 <> -1 Then
                                            grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPNAMEINDEX) = Trim$(tgCurrANE(ilANE2).sName)
                                    '        Exit For
                                        End If
                                    'Next ilANE2
                                Else
                                    grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPNAMEINDEX) = ""
                                End If
                                If tmCurrASE(ilASE).iBkupCceCode > 0 Then
                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                                        If tmCurrASE(ilASE).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                                            grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                                            Exit For
                                        End If
                                    Next ilCCE
                                Else
                                    grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPCTRLINDEX) = ""
                                End If
                                If tmCurrASE(ilASE).iProtAneCode > 0 Then
                                    'For ilANE2 = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tmCurrASE(ilASE).iProtAneCode = tgCurrANE(ilANE2).iCode Then
                                        ilANE2 = gBinarySearchANE(tmCurrASE(ilASE).iProtAneCode, tgCurrANE())
                                        If ilANE2 <> -1 Then
                                            grdLibEvents.TextMatrix(grdLibEvents.Row, PROTNAMEINDEX) = Trim$(tgCurrANE(ilANE2).sName)
                                    '        Exit For
                                        End If
                                    'Next ilANE2
                                Else
                                    grdLibEvents.TextMatrix(grdLibEvents.Row, PROTNAMEINDEX) = ""
                                End If
                                If tmCurrASE(ilASE).iProtCceCode > 0 Then
                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                                        If tmCurrASE(ilASE).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                                            grdLibEvents.TextMatrix(grdLibEvents.Row, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                                            Exit For
                                        End If
                                    Next ilCCE
                                Else
                                    grdLibEvents.TextMatrix(grdLibEvents.Row, PROTCTRLINDEX) = ""
                                End If
                            End If
                    '        Exit For
                        End If
                    'Next ilANE
                    If slStr = "" Then
                        Exit For
                    End If
                Next ilASE
            Case ENDTYPEINDEX
                '11/24/04- Allow end type and Duration to co-exist
                'If lbcTTE_E.ListIndex > 1 Then
                '    grdLibEvents.TextMatrix(grdLibEvents.Row, DURATIONINDEX) = ""
                'End If
        End Select
        '7/8/11: Make T2 work like T1
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
        Select Case grdLib.Col
            Case NAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcDNE, True
            Case SUBLIBNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcDSE, True
            Case BUSGROUPSINDEX
                gProcessArrowKey Shift, KeyCode, lbcBGE, True
        End Select
        Select Case grdLibEvents.Col
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
    Select Case grdLibEvents.Col
        Case TIMEINDEX
        Case DURATIONINDEX
        Case AIRHOURSINDEX
        Case AIRDAYSINDEX
        Case AUDIOITEMIDINDEX
        Case AUDIOISCIINDEX
        Case PROTITEMIDINDEX
        Case PROTISCIINDEX
        Case SILENCETIMEINDEX
    End Select
    If grdLibEvents.text <> edcEvent.text Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        Select Case grdLibEvents.Col
            Case TIMEINDEX
            Case DURATIONINDEX
                '11/24/04- Allow end type and Duration to co-exist
                'If Trim$(edcEvent.Text) <> "" Then
                '    grdLibEvents.TextMatrix(grdLibEvents.Row, ENDTYPEINDEX) = ""
                'End If
            Case AIRHOURSINDEX
            Case AIRDAYSINDEX
            Case AUDIOITEMIDINDEX
            Case AUDIOISCIINDEX
            Case PROTITEMIDINDEX
            Case PROTISCIINDEX
            Case SILENCETIMEINDEX
        End Select
        grdLibEvents.text = edcEvent.text
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub edcEvent_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcLib_Change()
    Dim slStr As String
    
    Select Case grdLib.Col
        Case DESCRIPTIONINDEX
        Case STARTDATEINDEX  'Date
        Case ENDDATEINDEX  'Date
        Case DAYSINDEX  'Date
        'Case STARTTIMEINDEX  'Date
        'Case LENGTHINDEX  'Date
        Case HOURSINDEX  'Date
    End Select
    If StrComp(Trim$(grdLib.text), Trim$(edcLib.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        Select Case grdLib.Col
            Case STARTDATEINDEX
                If grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX) <> "" Then
                    If gIsDate(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                        If gIsDate(edcDropdown.text) Then
                            If gDateValue(edcDropdown.text) + 6 > gDateValue(grdLib.TextMatrix(grdLib.Row, ENDDATEINDEX)) Then
                                grdLib.TextMatrix(grdLib.Row, DAYSINDEX) = ""
                            End If
                        End If
                    End If
                End If
            Case ENDDATEINDEX
                If grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX) <> "" Then
                    If gIsDate(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) Then
                        If gIsDate(edcDropdown.text) Then
                            If gDateValue(grdLib.TextMatrix(grdLib.Row, STARTDATEINDEX)) + 6 > gDateValue(edcDropdown.text) Then
                                grdLib.TextMatrix(grdLib.Row, DAYSINDEX) = ""
                            End If
                        End If
                    End If
                End If
            'Case STARTTIMEINDEX
            '    grdLib.TextMatrix(grdLib.Row, HOURSINDEX) = ""
            'Case LENGTHINDEX
            '    grdLib.TextMatrix(grdLib.Row, HOURSINDEX) = ""
        End Select
        grdLib.text = edcLib.text
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub edcLib_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSearch_GotFocus()
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        mFindMatch True
        If lgLibCallCode > 0 Then
            gGrid_FillWithRows grdLibEvents
            grdLib.Height = 3 * grdLib.RowHeight(0) + 15
            grdConflicts.Height = 4 * grdConflicts.RowHeight(0) + 15
            'edcSearch.SetFocus
            cmcCancel.SetFocus
        Else
            gGrid_FillWithRows grdLibEvents
            grdLib.Height = 3 * grdLib.RowHeight(0) + 15
            grdConflicts.Height = 4 * grdConflicts.RowHeight(0) + 15
            cmcCancel.SetFocus
        End If
    End If
    imFirstActivate = False
    Me.KeyPreview = True
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    bmIntegralSet = False
    'Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    'Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    Me.Move Me.Left, Me.Top, 0.97 * Screen.Width, 0.82 * Screen.Height
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrLibDef
    'gCenterFormModal EngrLibDef
    gCenterForm EngrLibDef
'    Unload EngrLib
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdLib.FixedRows) And (lmEnableRow < grdLib.Rows) Then
            If (lmEnableCol >= grdLib.FixedCols) And (lmEnableCol < grdLib.Cols) Then
                If (lmEnableCol <> BUSGROUPSINDEX) And (lmEnableCol <> BUSESINDEX) Then
                    If lmEnableCol = CONFLICTINDEX Then
                        smConflict = smESCValue
                    ElseIf lmEnableCol = STATEINDEX Then
                        smState = smESCValue
                    Else
                        grdLib.text = smESCValue
                    End If
                    mSetShow
                Else
                    grdLib.text = smESCValue
                End If
                mEnableBox
            End If
        End If
        If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
            If (lmEEnableCol >= grdLibEvents.FixedCols) And (lmEEnableCol < grdLibEvents.Cols) Then
                If lmEEnableCol <> FIXEDINDEX Then
                    grdLibEvents.text = smESCValue
                Else
                    smYN = smESCValue
                End If
                mESetShow
                mEEnableBox
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(LIBRARYJOB) = 2
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    'cmcCancel.Top = Me.Height - 2 * cmcCancel.Height - 225
    grdLibEvents.Height = cmcCancel.Top - grdLibEvents.Top - 240     '4 * grdLibEvents.RowHeight(0) + 15
    '8/26/11: Moved
    'gGrid_IntegralHeight grdLibEvents
    gGrid_FillWithRows grdLibEvents
    grdLib.Height = 3 * grdLib.RowHeight(0) + 15
    gGrid_IntegralHeight grdLib
    grdConflicts.Height = 4 * grdConflicts.RowHeight(0) + 15
    gGrid_IntegralHeight grdConflicts
    lacHelp.Top = grdLibEvents.Top + grdLibEvents.Height
    'cmcCancel.Top = lacHelp.Top + lacHelp.Height - 75
    'cmcDone.Top = cmcCancel.Top
    'cmcSave.Top = cmcCancel.Top
    'cmcReplace.Top = cmcCancel.Top + cmcCancel.Height + 60
    'cmcConflict.Top = cmcReplace.Top
    'cmcImport.Top = cmcReplace.Top
    imcInsert.Top = lacHelp.Top + lacHelp.Height + 120
    imcTrash.Top = imcInsert.Top
    imcPrint.Top = imcInsert.Top
    lmCharacterWidth = CLng(pbcTab.TextWidth("n"))
    'Adjust height so that the line under the scroll bar is not visible with IsRowVisible acll
    '8/26/11: Removed
    'grdLibEvents.Height = grdLibEvents.Height - 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    btrDestroy hmSEE
    btrDestroy hmSOE
    btrDestroy hmCTE
    btrDestroy hmCME
    btrDestroy hmDHE
    btrDestroy hmDEE
    
    Erase smHours
    Erase smDays
    Erase lmDeleteCodes
    Erase tmCurrBSE
    Erase smBuses
    Erase tmCurrDBE
    Erase tmCurrEBE
    Erase smT1Comment
    Erase tmCurr1CTE_Name
    Erase smT2Comment
    Erase tmCurr2CTE_Name
    Erase tmCurrDEE
    Erase smEBuses
    
    Erase smGridValues
    Erase smReplaceValues
       
    Erase tmCurrLibDBE
    Erase tmCurrLibDEE
    Erase tmCurrLibEBE
    
    Erase tmConflictList
    Erase tmConflictTest
    
    Set EngrLibDef = Nothing
    EngrLib.Show vbModeless
End Sub





Private Sub mInit()
    Dim llRow As Long
    Dim llCol As Long
    Dim slDate As String
    Dim ilRet As Integer
    'Dim slEDate As String
    'Dim slLDate As String
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdLib, grdLibEvents, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim lmDeleteCodes(0 To 0) As Long
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
    ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
    'Can't be 0 to 0 because of index in grid
    ReDim tmCurrDEE(1 To 1) As DEE
'    cmcSearch.Top = 30
'    edcSearch.Top = cmcSearch.Top
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    lmEEnableRow = -1
    lmConflictRow = -1
    imFirstActivate = True
    imInChg = True
    imIgnoreBDEChg = False
    imLimboAllowed = False
    imDefaultProgIndex = -1
    bmInBranch = False
    bmPrinting = False
    bmInSave = False
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmSOE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSOE, "", sgDBPath & "SOE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCME = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCME, "", sgDBPath & "CME.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmDHE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmDHE, "", sgDBPath & "DHE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmDEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmDEE, "", sgDBPath & "DEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    mPopATE
    mPopANE
    mPopASE
    mPopBGE
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
    mPopulate
    gConflictPop
    If sgClientFields = "A" Then
        '8/26/: Retained adding horizontal scroll bar
        grdLibEvents.ScrollBars = flexScrollBarBoth
        imMaxCols = ABCRECORDITEMINDEX
    Else
        imMaxCols = TITLE2INDEX
    End If
    imStartDateChangable = True
    imLibCallType = igLibCallType
    lmLibCallCode = lgLibCallCode
    If igLibCallType = 1 Then
        lacScreen.Caption = "Change Library Definition"
    ElseIf igLibCallType = 2 Then
        lacScreen.Caption = "Create New Library Definition by Modelling"
    ElseIf igLibCallType = 3 Then
        lacScreen.Caption = "View Library Definition"
        cmcReplace.Enabled = False
        cmcImport.Enabled = False
    ElseIf igLibCallType = 4 Then
        lacScreen.Caption = "Terminate Library Definition"
        cmcReplace.Enabled = False
        cmcImport.Enabled = False
    Else
        lacScreen.Caption = "Create New Library Definition from Scratch"
    End If
    'slEDate = gGetEarlestSchdDate(False)
    'If slEDate <> "" Then
    '    slLDate = gGetLatestSchdDate(False)
    '    lacScreen.Caption = lacScreen.Caption & ", Scheduled " & slEDate & "-" & slLDate
    'End If
'    If lgLibCallCode > 0 Then
'        mMoveRecToCtrls
'        mMoveDEERecToCtrls
'        mSortCol TIMEINDEX
'        'If (igLibCallType = 1) Then
'        '    slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
'        '    smNowDate = Format(gNow(), sgShowDateForm)
'        '    If gDateValue(slDate) < gDateValue(smNowDate) Then
'        '        grdLib.Col = STARTDATEINDEX
'        '        grdLib.Row = grdLib.FixedRows
'        '        grdLib.CellBackColor = LIGHTYELLOW
'        '        imStartDateChangable = False
'        '    End If
'        '    If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
'        '        imLimboAllowed = True
'        '    End If
'        'ElseIf igLibCallType = 2 Then
'        '    grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX) = ""
'        '    grdLib.TextMatrix(grdLib.FixedRows, ENDDATEINDEX) = ""
'        '    grdLib.TextMatrix(grdLib.FixedRows, CODEINDEX) = "0"
'        '    igLibCallType = 0
'        '    lgLibCallCode = 0
'        '    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
'        '        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
'        '            grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
'        '            grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
'        '        End If
'        '    Next llRow
'        '    tmDHE.lCode = 0
'        '    imLimboAllowed = True
'        'ElseIf (igLibCallType = 4) Then
'        '    slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
'        '    smNowDate = Format(gNow(), sgShowDateForm)
'        '    If gDateValue(slDate) < gDateValue(smNowDate) Then
'        '        grdLib.Col = STARTDATEINDEX
'        '        grdLib.Row = grdLib.FixedRows
'        '        grdLib.CellBackColor = LIGHTYELLOW
'        '        imStartDateChangable = False
'        '    End If
'        '    grdLib.Row = grdLib.FixedRows
'        '    For llCol = NAMEINDEX To STATEINDEX Step 1
'        '        If (llCol <> ENDDATEINDEX) Then
'        '            If (llCol <> SUBLIBNAMEINDEX) Or (Trim$(grdLib.TextMatrix(grdLib.FixedRows, llCol)) <> "") Then
'        '                grdLib.Col = llCol
'        '                grdLib.CellBackColor = LIGHTYELLOW
'        '            End If
'        '        End If
'        '    Next llCol
'        '    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
'        '        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
'        '            For llCol = EVENTTYPEINDEX To imMaxCols Step 1
'        '                grdLibEvents.Col = llCol
'        '                grdLibEvents.Row = llRow
'        '                grdLibEvents.CellBackColor = LIGHTYELLOW
'        '            Next llCol
'        '        End If
'        '    Next llRow
'        '    If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
'        '        imLimboAllowed = True
'        '    End If
'        'End If
'        mSetBackgroundColor
'    Else
'        imLimboAllowed = True
'    End If
    imInChg = False
    imFieldChgd = False
    mSetCommands
    If igLibCallType <> 3 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(LIBRARYJOB) = 2) Then
            cmcDone.Enabled = True
            If igLibCallType = 4 Then
                imcInsert.Enabled = False
                imcTrash.Enabled = False
            End If
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
    tmcStart.Enabled = True
    gSetMousePointer grdLib, grdLibEvents, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdLib, grdLibEvents, vbDefault
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
    gHandleError "EngrErrors.Txt", "Library Definition-Form Load"
    Resume Next
End Sub

Private Sub grdLibEvents_Click()
    If grdLibEvents.Col >= grdLibEvents.Cols - 1 Then
        Exit Sub
    End If

End Sub

Private Sub grdLibEvents_EnterCell()
    mESetShow
    mSetShow
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
    
    grdLibEvents.ToolTipText = ""
    If (y > grdLibEvents.RowHeight(0)) And (y < grdLibEvents.RowHeight(0) + grdLibEvents.RowHeight(1)) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdLibEvents, x, y, llRow, llCol)
    grdLibEvents.ToolTipText = Trim$(grdLibEvents.TextMatrix(llRow, llCol))
End Sub

Private Sub grdLibEvents_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    If bmInSave Then
        grdLibEvents.Redraw = True
        Exit Sub
    End If
    'If same cell entered after clicking some other place, a enter cell event does not happen
    mSetShow
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
        grdLibEvents.LeftCol = HIGHLIGHTINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        grdLibEvents.Redraw = True
    End If
    grdLibEvents.Redraw = True
    '8/26/11: Check that row is not behind scroll bar
    If grdLibEvents.RowPos(grdLibEvents.Row) + grdLibEvents.RowHeight(grdLibEvents.Row) + 60 >= grdLibEvents.Height Then
        imIgnoreScroll = True
        grdLibEvents.TopRow = grdLibEvents.TopRow + 1
    End If
    If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
        mEEnableBox
    Else
        Beep
        pbcClickFocus.SetFocus
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
            lacHelp.Visible = True
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            dpcEvent.Visible = False
            hpcEvent.Visible = False
            ltcEvent.Visible = False
            pbcEDefine.Visible = False
            edcEDropdown.Visible = False
            cmcEDropDown.Visible = False
            lbcBuses.Visible = False
            lbcCCE_B.Visible = False
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
            lbcCTE_2.Visible = False
            lbcCTE_1.Visible = False
            pbcArrow.Visible = False
            mHideConflictGrid
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        mHideConflictGrid
        imFromArrow = False
    End If

End Sub

Private Sub hpcEvent_OnChange()
    If StrComp(Trim$(grdLibEvents.text), Trim$(hpcEvent.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        grdLibEvents.text = hpcEvent.text
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub hpcLib_OnChange()
    If StrComp(Trim$(grdLib.text), Trim$(hpcLib.text), vbTextCompare) <> 0 Then
        imFieldChgd = True
        grdLib.text = hpcLib.text
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub imcInsert_Click()
    If bmInSave Then
        Exit Sub
    End If
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    'igRptIndex = LIBRARYEVENT_RPT
    'igRptSource = vbModal
    'EngrLibEvtRpt.Show vbModal
    
    
    Dim ilRptDest As Integer            'disply, print, save as file
    Dim slRptName As String
    Dim slExportName As String
    Dim slRptType As String
    Dim llResult As Long
    Dim ilExportType As Integer
    Dim llGridRow As Long
    Dim slStr As String
    Dim llTime As Long
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim llEndDate As Long
    Dim slEndDate As String
    Dim slFilter As String              'filters selected by user
    Dim ilLoop As Integer
    Dim slOperator As String * 2        'operator for filter
    Dim slIgnoreConflicts As String
    Dim slLibStartDate As String
    Dim slLibEndDate As String
    Dim slState As String
    

    If bmInSave Then
        Exit Sub
    End If
    If bmPrinting Then
        Exit Sub
    End If
    bmPrinting = True
    igRptIndex = LIBRARYEVENT_RPT
    igRptSource = vbModal
    'slRptName = "LibEvtSS.rpt"      'concatenate the crystal report name plus extension

    slExportName = ""               'no export for now
    
    Set rstLibEvts = New Recordset
    gGeneraterstLibEvts       'create Library Events snapshot
    rstLibEvts.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstLibEvts, sgDBPath & "\Libdet.ttx", True)
    
    slStartDate = Trim$(grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX))
    llStartDate = gDateValue(slStartDate)
    slStartDate = gAdjYear(Format$(llStartDate, "ddddd")) 'insure year is appended
    slEndDate = Trim$(grdLib.TextMatrix(grdLib.FixedRows, ENDDATEINDEX))
    If slEndDate = "" Then
        slEndDate = "12/31/2069"
    Else
        slEndDate = gAdjYear(slEndDate)
    End If
    llEndDate = gDateValue(slEndDate)
    slEndDate = Format$(llEndDate, "ddddd") 'insure year is appended
    slStr = Trim$(grdLib.TextMatrix(grdLib.FixedRows, CONFLICTINDEX))
    If slStr = "Audio" Then
        slIgnoreConflicts = "A"
    ElseIf slStr = "Bus" Then
        slIgnoreConflicts = "B"
    ElseIf slStr = "Both" Then      'ignore both
        slIgnoreConflicts = "I"
    Else
        slIgnoreConflicts = "N"
    End If
    
    slLibStartDate = Trim$(grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX))
    slLibEndDate = Trim$(grdLib.TextMatrix(grdLib.FixedRows, ENDDATEINDEX))
    slStr = Trim$(grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX))

    If slStr = "Active" Then
        slState = "A"
    ElseIf slStr = "Dormant" Then
        slState = "D"
    Else
        slState = "L"
    End If

    'loop thru the ItemID grid to print whats shown on the screen
    For llGridRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1
        slStr = Trim$(grdLibEvents.TextMatrix(llGridRow, EVENTTYPEINDEX))
        If slStr = "" Then
            Exit For
        Else
            rstLibEvts.AddNew
            rstLibEvts.Fields("Name") = Trim$(grdLib.TextMatrix(grdLib.FixedRows, NAMEINDEX))
            rstLibEvts.Fields("SubName") = Trim$(grdLib.TextMatrix(grdLib.FixedRows, SUBLIBNAMEINDEX))
            rstLibEvts.Fields("EventType") = Left(grdLibEvents.TextMatrix(llGridRow, EVENTTYPEINDEX), 1)   'program, spot, avail
            rstLibEvts.Fields("EvBusDesc") = Trim$(grdLibEvents.TextMatrix(llGridRow, BUSNAMEINDEX))              'Bus name
            rstLibEvts.Fields("EvBusCtl") = grdLibEvents.TextMatrix(llGridRow, BUSCTRLINDEX)                'Bus Control index
            rstLibEvts.Fields("EvStarttime") = grdLibEvents.TextMatrix(llGridRow, TIMEINDEX)               'Event start time
            slStr = Trim$(grdLibEvents.TextMatrix(llGridRow, TIMEINDEX))
            llTime = gStrTimeInTenthToLong(slStr, False)                'convert the start time of event to long for sorting
            rstLibEvts.Fields("EvStartTimeSort") = llTime
            rstLibEvts.Fields("EvStartType") = Trim$(grdLibEvents.TextMatrix(llGridRow, STARTTYPEINDEX))          'start type
            rstLibEvts.Fields("EvFix") = grdLibEvents.TextMatrix(llGridRow, FIXEDINDEX)                    'Fixed type
            rstLibEvts.Fields("EvEndType") = Trim$(grdLibEvents.TextMatrix(llGridRow, ENDTYPEINDEX))              'end type
            rstLibEvts.Fields("EvDur") = Trim$(grdLibEvents.TextMatrix(llGridRow, DURATIONINDEX))                 'duration
            rstLibEvts.Fields("EvHours") = Trim$(grdLibEvents.TextMatrix(llGridRow, AIRHOURSINDEX))
            rstLibEvts.Fields("EvDays") = Trim$(grdLibEvents.TextMatrix(llGridRow, AIRDAYSINDEX))
            rstLibEvts.Fields("EvMatType") = Trim$(grdLibEvents.TextMatrix(llGridRow, MATERIALINDEX))             'material type
            rstLibEvts.Fields("EvAudName1") = Trim$(grdLibEvents.TextMatrix(llGridRow, AUDIONAMEINDEX))           'primary audio name
            rstLibEvts.Fields("EvItem1") = Trim$(grdLibEvents.TextMatrix(llGridRow, AUDIOITEMIDINDEX))            'primary audio item id
            'rstLibEvts.Fields("EvISCI1") = Trim$(grdLibEvents.TextMatrix(llGridRow, AUDIOISCIINDEX))            'primary audio item id
            rstLibEvts.Fields("EvCtl1") = Trim$(grdLibEvents.TextMatrix(llGridRow, AUDIOCTRLINDEX))               'primary audio control
            rstLibEvts.Fields("EvAudName2") = Trim$(grdLibEvents.TextMatrix(llGridRow, BACKUPNAMEINDEX))         'backup audio name
            rstLibEvts.Fields("EvCtl2") = grdLibEvents.TextMatrix(llGridRow, BACKUPCTRLINDEX)              'back control char
            rstLibEvts.Fields("EvAudName3") = Trim$(grdLibEvents.TextMatrix(llGridRow, PROTNAMEINDEX))            'protection audio name
            rstLibEvts.Fields("EvItem3") = Trim$(grdLibEvents.TextMatrix(llGridRow, PROTITEMIDINDEX))             'protection item id
            'rstLibEvts.Fields("EvISCI3") = Trim$(grdLibEvents.TextMatrix(llGridRow, PROTISCIINDEX))             'protection item id
            rstLibEvts.Fields("EvCtl3") = grdLibEvents.TextMatrix(llGridRow, PROTCTRLINDEX)                'protection control
            rstLibEvts.Fields("EvRelay1") = Trim$(grdLibEvents.TextMatrix(llGridRow, RELAY1INDEX))                'relay 1 of 2
            rstLibEvts.Fields("EvRelay2") = Trim$(grdLibEvents.TextMatrix(llGridRow, RELAY2INDEX))                'relay 2 of 2
            rstLibEvts.Fields("EvFollow") = Trim$(grdLibEvents.TextMatrix(llGridRow, FOLLOWINDEX))                'follow name
            rstLibEvts.Fields("EvTime") = Trim$(grdLibEvents.TextMatrix(llGridRow, SILENCETIMEINDEX))      'silence time
            rstLibEvts.Fields("EvSilence1") = Trim$(grdLibEvents.TextMatrix(llGridRow, SILENCE1INDEX))            'silence name 1 of 4
            rstLibEvts.Fields("EvSilence2") = Trim$(grdLibEvents.TextMatrix(llGridRow, SILENCE2INDEX))            'silence name 2 of 4
            rstLibEvts.Fields("EvSilence3") = Trim$(grdLibEvents.TextMatrix(llGridRow, SILENCE3INDEX))            'silence name 3 of 4
            rstLibEvts.Fields("EvSilence4") = Trim$(grdLibEvents.TextMatrix(llGridRow, SILENCE4INDEX))            'silence name 4 of 4
            rstLibEvts.Fields("EvNetCue1") = Trim$(grdLibEvents.TextMatrix(llGridRow, NETCUE1INDEX))              'netcue name 1 of 2
            rstLibEvts.Fields("EvNetCue2") = Trim$(grdLibEvents.TextMatrix(llGridRow, NETCUE2INDEX))              'netcue name 2 of 2
            rstLibEvts.Fields("EvTitle1Desc") = Trim$(grdLibEvents.TextMatrix(llGridRow, TITLE1INDEX))               'title 1 of 2
            rstLibEvts.Fields("EvTitle2Desc") = Trim$(grdLibEvents.TextMatrix(llGridRow, TITLE2INDEX))               'title 2 of 2
            rstLibEvts.Fields("EvABCCustomFields") = ""
            If sgClientFields = "A" Then            'abc client
                If Trim$(grdLibEvents.TextMatrix(llGridRow, ABCFORMATINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCPGMCODEINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCXDSMODEINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCRECORDITEMINDEX)) <> "" Then
                    rstLibEvts.Fields("EvABCCustomFields") = "ABC: Format-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCFORMATINDEX)) & ", Pgm Code-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCPGMCODEINDEX)) & ", XDS Mode-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCXDSMODEINDEX)) & ", Record Item-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCRECORDITEMINDEX))
                End If
            End If
            rstLibEvts.Fields("IgnoreConflicts") = slIgnoreConflicts
            rstLibEvts.Fields("StartDate") = Trim$(slLibStartDate)
            rstLibEvts.Fields("EndDate") = Trim$(slLibEndDate)
            rstLibEvts.Fields("State") = slState
        End If
    Next llGridRow
    
    sgCrystlFormula2 = "Date(" + Format$(gAdjYear(slStartDate), "yyyy") + "," + Format$(slStartDate, "mm") + "," + Format$(slStartDate, "dd") + ")"
    sgCrystlFormula3 = "Date(" + Format$(gAdjYear(slEndDate), "yyyy") + "," + Format$(slEndDate, "mm") + "," + Format$(slEndDate, "dd") + ")"
    
    sgCrystlFormula4 = " "                                     'blank, unused.  retain for common code in calling crystal events report
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    igRptSource = vbModal
        
    EngrSchedRpt.Show vbModal
    
    'determine how the user responded, either cancel or produce output
    If igReturnCallStatus = CALLDONE Then           'produce the report flag
    
        slExportName = sgReturnCallName     'if exporting path and filename, this is filename; otherwise blank
        slRptType = ""
        'determine which version (condensed # of fields or all fields)
        If sgReturnOption = "ALL" Then
            slRptName = "LibEvtSSAll.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstLibEvts
        Else
            slRptName = "LibEvtSS.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstLibEvts
        End If
    End If
    gSetMousePointer grdLibEvents, grdLibEvents, vbDefault
    Set rstLibEvts = Nothing
    cmcCancel.SetFocus
    bmPrinting = False
    Exit Sub
End Sub

Private Sub imcTrash_Click()
    If bmInSave Then
        Exit Sub
    End If
    mESetShow
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    mDeleteRow
End Sub



Private Sub grdLib_Click()
    If grdLib.Col >= grdLib.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdLib_EnterCell()
    mESetShow
    mSetShow
End Sub

Private Sub grdLib_GotFocus()
    If grdLib.Col >= grdLib.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdLib_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdLib.TopRow
    grdLib.Redraw = False
End Sub

Private Sub grdLib_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    If bmInSave Then
        grdLib.Redraw = True
        Exit Sub
    End If
    'If same cell entered after clicking some other place, a enter cell event does not happen
    mESetShow
    'Determine if in header
    If y < grdLib.RowHeight(0) Then
'        mSortCol grdLib.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdLib, x, y)
    If Not ilFound Then
        grdLib.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdLib.Col >= grdLib.Cols - 1 Then
        grdLib.Redraw = True
        Exit Sub
    End If
    If (grdLib.Col = STARTDATEINDEX) And (Not imStartDateChangable) Then
        grdLib.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    lmTopRow = grdLib.TopRow
    
'    llRow = grdLib.Row
'    If grdLib.TextMatrix(llRow, CATEGORYINDEX) = "" Then
'        grdLib.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdLib.TextMatrix(llRow, CATEGORYINDEX) = ""
'        grdLib.Row = llRow + 1
'        grdLib.Col = CATEGORYINDEX
'        grdLib.Redraw = True
'    End If
    grdLib.Redraw = True
    'If Library name with focus and then State pressed, then the entercell event does not happen (I don't know why)
    mSetShow
    mEnableBox
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
        'For ilLoop = 0 To UBound(tmCurrASE) - 1 Step 1
        '    If ilCode = tmCurrASE(ilLoop).iCode Then
            ilLoop = gBinarySearchASE(ilCode, tmCurrASE())
            If ilLoop <> -1 Then
                lbcASE.ToolTipText = Trim$(tmCurrASE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcBDE_Click()
    Dim slStr As String
    Dim ilLoop As Integer
    If imIgnoreBDEChg Then
        Exit Sub
    End If
    slStr = ""
    For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
        If lbcBDE.Selected(ilLoop) Then
            slStr = slStr & lbcBDE.List(ilLoop) & ","
        End If
    Next ilLoop
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    grdLib.text = slStr
    grdLib.CellForeColor = vbBlack
    imFieldChgd = True
    mSetCommands

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

Private Sub lbcBGE_Click()
    Dim slStr As String
    Dim ilLoop As Integer
    slStr = ""
    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
        If lbcBGE.Selected(ilLoop) Then
            slStr = slStr & lbcBGE.List(ilLoop) & ","
        End If
    Next ilLoop
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    grdLib.text = slStr
    grdLib.CellForeColor = vbBlack
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub lbcBGE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBGE, y)
    If (llRow < lbcBGE.ListCount) And (lbcBGE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBGE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
            If ilCode = tgCurrBGE(ilLoop).iCode Then
                lbcBGE.ToolTipText = Trim$(tgCurrBGE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcBuses_Click()
    Dim slStr As String
    Dim ilLoop As Integer
    slStr = ""
    For ilLoop = 0 To lbcBuses.ListCount - 1 Step 1
        If lbcBuses.Selected(ilLoop) Then
            slStr = slStr & lbcBuses.List(ilLoop) & ","
        End If
    Next ilLoop
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    If grdLibEvents.text <> slStr Then
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
    End If
    grdLibEvents.text = slStr
    If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
        grdLibEvents.CellForeColor = vbBlue
    Else
        grdLibEvents.CellForeColor = vbBlack
    End If
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub lbcBuses_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBuses, y)
    If (llRow < lbcBuses.ListCount) And (lbcBuses.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBuses.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If ilCode = tgCurrBDE(ilLoop).iCode Then
            ilLoop = gBinarySearchBDE(ilCode, tgCurrBDE())
            If ilLoop <> -1 Then
                lbcBuses.ToolTipText = Trim$(tgCurrBDE(ilLoop).sDescription)
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
    '7/8/11: Make T2 work like T1
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

Private Sub lbcDNE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcDNE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcDNE.Visible = False
End Sub

Private Sub lbcDNE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llCode As Long
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcDNE, y)
    If (llRow < lbcDNE.ListCount) And (lbcDNE.ListCount > 0) And (llRow <> -1) Then
        llCode = lbcDNE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
            If llCode = tgCurrLibDNE(ilLoop).lCode Then
                lbcDNE.ToolTipText = Trim$(tgCurrLibDNE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If

End Sub

Private Sub lbcDSE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcDSE.List(lbcDSE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcDSE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcDSE.Visible = False
End Sub

Private Sub lbcDSE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llCode As Long
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcDSE, y)
    If (llRow < lbcDSE.ListCount) And (lbcDSE.ListCount > 0) And (llRow <> -1) Then
        llCode = lbcDSE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrDSE) - 1 Step 1
            If llCode = tgCurrDSE(ilLoop).lCode Then
                lbcDSE.ToolTipText = Trim$(tgCurrDSE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
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
    Dim ilPos As Integer
    
    slStr = ltcEvent.text
'    ilPos = InStr(1, slStr, ":", vbTextCompare)
'    If ilPos > 0 Then
'        slStr = Mid$(slStr, ilPos + 1)
'    End If
    If grdLibEvents.text <> slStr Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
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
    mSetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcConflict_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smConflict <> "Audio" Then
            imFieldChgd = True
        End If
        smConflict = "Audio"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
        If (smConflict = "Bus") Then
            imFieldChgd = True
            smConflict = "Both"
        ElseIf (smConflict = "Both") Then
            imFieldChgd = True
            smConflict = "Bus"
        Else
            imFieldChgd = True
            smConflict = "Bus"
        End If
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) And (imLimboAllowed) Then
        If smConflict <> "No" Then
            imFieldChgd = True
        End If
        smConflict = "No"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smConflict = "Audio" Then
            imFieldChgd = True
            smConflict = "Bus"
            pbcConflict_Paint
            grdLib.CellForeColor = vbBlack
        ElseIf smConflict = "Bus" Then
            imFieldChgd = True
            smConflict = "Both"
            pbcConflict_Paint
            grdLib.CellForeColor = vbBlack
        ElseIf smConflict = "Both" Then
            imFieldChgd = True
            smConflict = "No"
            pbcConflict_Paint
            grdLib.CellForeColor = vbBlack
        ElseIf smConflict = "No" Then
            imFieldChgd = True
            smConflict = "Audio"
            pbcConflict_Paint
            grdLib.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcConflict_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smConflict = "Audio" Then
        imFieldChgd = True
        smConflict = "Bus"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf smConflict = "Bus" Then
        imFieldChgd = True
        smConflict = "Both"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf smConflict = "Both" Then
        imFieldChgd = True
        smConflict = "No"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf smConflict = "No" Then
        imFieldChgd = True
        smConflict = "Audio"
        pbcConflict_Paint
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcConflict_Paint()
    pbcConflict.Cls
    pbcConflict.CurrentX = 30  'fgBoxInsetX
    pbcConflict.CurrentY = 0 'fgBoxInsetY
    pbcConflict.Print smConflict
End Sub

Private Sub pbcDefine_Click()
'    Dim ilRet As Integer
'    ilRet = mBranch()
'    pbcDefine.SetFocus
End Sub

Private Sub pbcDefine_Paint()
    pbcDefine.CurrentX = 30
    pbcDefine.CurrentY = 0
    pbcDefine.Print "Multi-Select"
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
    
    If bmInBranch Then
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
    If pbcHighlight.Visible Or edcEvent.Visible Or edcEDropdown.Visible Or pbcYN.Visible Or pbcEDefine.Visible Or ltcEvent.Visible Or hpcEvent.Visible Or dpcEvent.Visible Then
        If Not lbcBuses.Visible Then
            If Not mEBranch() Then
                mEEnableBox
                bmInBranch = False
                Exit Sub
            End If
        End If
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
                    mEEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdLibEvents.Col = grdLibEvents.Col - 1
                If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
                    mEEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdLibEvents.TopRow = grdLibEvents.FixedRows
        grdLibEvents.LeftCol = HIGHLIGHTINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        grdLibEvents.Row = grdLibEvents.FixedRows
        If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
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
    If GetFocus() <> pbcETab.hwnd Then
        Exit Sub
    End If
    If pbcHighlight.Visible Or edcEvent.Visible Or edcEDropdown.Visible Or pbcYN.Visible Or pbcEDefine.Visible Or ltcEvent.Visible Or hpcEvent.Visible Or dpcEvent.Visible Then
        If Not lbcBuses.Visible Then
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
                    grdLibEvents.LeftCol = HIGHLIGHTINDEX
                    grdLibEvents.Col = EVENTTYPEINDEX
                    DoEvents
                    'grdLibEvents.TextMatrix(grdLibEvents.Row, CODEINDEX) = 0
                    If Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX)) <> "" Then
                        If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
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
                        grdLibEvents.LeftCol = HIGHLIGHTINDEX
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
                If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
                    mEEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdLibEvents.TopRow = grdLibEvents.FixedRows
        grdLibEvents.LeftCol = HIGHLIGHTINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        DoEvents
        grdLibEvents.Row = grdLibEvents.FixedRows
        If mColOk(grdLibEvents.Row, grdLibEvents.Col) Then
            mEEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer
    
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If edcLib.Visible Or pbcDefine.Visible Or pbcConflict.Visible Or pbcState.Visible Or edcDropdown.Visible Or hpcLib.Visible Or dpcLib.Visible Or cccLib.Visible Then
        If Not pbcDefine.Visible Then
            If Not mBranch() Then
                mEnableBox
                Exit Sub
            End If
        End If
        If grdLib.Col = STARTDATEINDEX Then
'            If Not gIsDate(edcLib.Text) Then
'                Beep
'                edcLib.SetFocus
'                Exit Sub
'            End If
            If Not gIsDate(cccLib.text) Then
                Beep
                cccLib.SetFocus
                Exit Sub
            End If
        ElseIf grdLib.Col = ENDDATEINDEX Then
'            If Trim$(edcLib.Text) <> "" Then
'                If Not gIsDate(edcLib.Text) Then
'                    Beep
'                    edcLib.SetFocus
'                    Exit Sub
'                End If
'            End If
            If Trim$(cccLib.text) <> "" Then
                If Not gIsDate(cccLib.text) Then
                    Beep
                    cccLib.SetFocus
                    Exit Sub
                End If
            End If
        'ElseIf grdLib.Col = STARTTIMEINDEX Then
        '    If Not gIsTime(edcLib.Text) Then
        '        Beep
        '        edcLib.SetFocus
        '        Exit Sub
        '    End If
        'ElseIf grdLib.Col = LENGTHINDEX Then
        '    If Not gIsLength(edcLib.Text) Then
        '        Beep
        '        edcLib.SetFocus
        '        Exit Sub
        '    End If
        End If
        mSetShow
        Do
            ilPrev = False
            If grdLib.Col = NAMEINDEX Then
                cmcCancel.SetFocus
            Else
                grdLib.Col = grdLib.Col - 1
                If grdLib.CellBackColor <> LIGHTYELLOW Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        grdLib.Col = NAMEINDEX
        grdLib.Row = grdLib.FixedRows
        mEnableBox
    End If
End Sub

Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smState <> "Active" Then
            imFieldChgd = True
        End If
        smState = "Active"
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("L") Or (KeyAscii = Asc("l")) And (imLimboAllowed) Then
        If smState <> "Limbo" Then
            imFieldChgd = True
        End If
        smState = "Limbo"
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            If imLimboAllowed Then
                smState = "Limbo"
            Else
                smState = "Dormant"
            End If
            pbcState_Paint
            grdLib.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdLib.CellForeColor = vbBlack
        ElseIf smState = "Limbo" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdLib.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        If imLimboAllowed Then
            smState = "Limbo"
        Else
            smState = "Dormant"
        End If
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    ElseIf smState = "Limbo" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdLib.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = 30  'fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    pbcState.Print smState
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If edcLib.Visible Or pbcDefine.Visible Or edcDropdown.Visible Or pbcConflict.Visible Or pbcState.Visible Or hpcLib.Visible Or dpcLib.Visible Or cccLib.Visible Then
        If Not pbcDefine.Visible Then
            If Not mBranch() Then
                mEnableBox
                Exit Sub
            End If
        End If
        If grdLib.Col = STARTDATEINDEX Then
'            If Not gIsDate(edcLib.Text) Then
'                Beep
'                edcLib.SetFocus
'                Exit Sub
'            End If
            If Not gIsDate(cccLib.text) Then
                Beep
                cccLib.SetFocus
                Exit Sub
            End If
        ElseIf grdLib.Col = ENDDATEINDEX Then
'            If Trim$(edcLib.text) <> "" Then
'                If Not gIsDate(edcLib.text) Then
'                    Beep
'                    edcLib.SetFocus
'                    Exit Sub
'                End If
'            End If
            If Trim$(cccLib.text) <> "" Then
                If Not gIsDate(cccLib.text) Then
                    Beep
                    cccLib.SetFocus
                    Exit Sub
                End If
            End If
        'ElseIf grdLib.Col = STARTTIMEINDEX Then
        '    If Not gIsTime(edcLib.Text) Then
        '        Beep
        '        edcLib.SetFocus
        '        Exit Sub
        '    End If
        'ElseIf grdLib.Col = LENGTHINDEX Then
        '    If Not gIsLength(edcLib.Text) Then
        '        Beep
        '        edcLib.SetFocus
        '        Exit Sub
        '    End If
        End If
        mSetShow
        Do
            ilNext = False
            If grdLib.Col = STATEINDEX Then
                pbcESTab.SetFocus
            Else
                grdLib.Col = grdLib.Col + 1
                If grdLib.CellBackColor <> LIGHTYELLOW Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        grdLib.Col = NAMEINDEX
        grdLib.Row = grdLib.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    Dim llFirstInsertRow As Long
    Dim ilInsert As Integer
    Dim ilCol As Integer
    
    llTRow = grdLibEvents.TopRow
    llRow = grdLibEvents.Row
'    slMsg = "Insert above selected Row"
'    If MsgBox(slMsg, vbYesNo) = vbNo Then
'        mInsertRow = False
'        Exit Function
'    End If
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
                grdLibEvents.TextMatrix(llRow, ilCol) = grdLibEvents.TextMatrix(llFirstInsertRow - 1, ilCol)
            Next ilCol
        End If
        grdLibEvents.Redraw = False
    Next ilInsert
    DoEvents
    grdLibEvents.Row = llFirstInsertRow
    grdLibEvents.TopRow = llTRow
    grdLibEvents.Redraw = True
    DoEvents
    grdLibEvents.LeftCol = HIGHLIGHTINDEX
    grdLibEvents.Col = EVENTTYPEINDEX
    mEEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdLibEvents.TopRow
    llRow = grdLibEvents.Row
    If (Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) <> 0) And (grdLib.TextMatrix(grdLib.FixedRows, USEDFLAGINDEX) = "Y") Then
        MsgBox "Row used or was used, unable to delete", vbInformation + vbOK
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete selected Row"
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdLibEvents.Redraw = False
    If (Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) <> 0) Then
        lmDeleteCodes(UBound(lmDeleteCodes)) = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
        ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
    End If
    grdLibEvents.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdLibEvents.AddItem ""
    grdLibEvents.Redraw = False
    grdLibEvents.TopRow = llTRow
    grdLibEvents.Redraw = True
    DoEvents
    'grdLibEvents.Col = CATEGORYINDEX
    'mEnableBox
    cmcCancel.SetFocus
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As DHE, tlOld As DHE) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilDBE As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    
    If tlNew.lDneCode <> tlOld.lDneCode Then
        mCompare = False
        Exit Function
    End If
    If tlNew.lDseCode <> tlOld.lDseCode Then
        mCompare = False
        Exit Function
    End If
    If gTimeToLong(tlNew.sStartTime, False) <> gTimeToLong(tlOld.sStartTime, False) Then
        mCompare = False
        Exit Function
    End If
    If tlNew.lLength <> tlOld.lLength Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sHours, tlOld.sHours, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If gDateValue(tlNew.sStartDate) <> gDateValue(tlOld.sStartDate) Then
        mCompare = False
        Exit Function
    End If
    If gDateValue(tlNew.sEndDate) <> gDateValue(tlOld.sEndDate) Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDays, tlOld.sDays, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    ilRet = gGetRec_CTE_CommtsTitle(tlOld.lCteCode, "EngrLibDef- mCompare for CTE", tmCTE)
    If ilRet Then
        If StrComp(Trim$(smDHEComment), Trim$(tmCTE.sComment), vbTextCompare) <> 0 Then
            mCompare = False
            Exit Function
        End If
    Else
        If Trim$(smDHEComment) <> "" Then
            mCompare = False
            Exit Function
        End If
    End If
    
    gParseCDFields smDHEBusGroups, False, smBusGroups()
    Erase tmCurrDBE
    smCurrDBEStamp = ""
    ilRet = gGetRecs_DBE_DayBusSel(smCurrDBEStamp, tlOld.lCode, "Bus Definition-mMoveRecToCtrls", tmCurrDBE())
    For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
        slStr = Trim$(smBusGroups(ilLoop))
        If slStr <> "" Then
            ilFound = False
            For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
                If tmCurrDBE(ilDBE).sType = "G" Then
                    For ilBGE = 0 To UBound(tgCurrBGE) - 1 Step 1
                        If tmCurrDBE(ilDBE).iBgeCode = tgCurrBGE(ilBGE).iCode Then
                            If StrComp(Trim$(tgCurrBGE(ilBGE).sName), slStr, vbTextCompare) = 0 Then
                                ilFound = True
                                tmCurrDBE(ilDBE).sType = ""
                            End If
                            Exit For
                        End If
                    Next ilBGE
                End If
            Next ilDBE
            If Not ilFound Then
                mCompare = False
                Exit Function
            End If
        End If
    Next ilLoop
    For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
        If tmCurrDBE(ilDBE).sType = "G" Then
            mCompare = False
            Exit Function
        End If
    Next ilDBE
    gParseCDFields smDHEBuses, False, smBuses()
    For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
        slStr = Trim$(smBuses(ilLoop))
        If slStr <> "" Then
            ilFound = False
            For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
                If tmCurrDBE(ilDBE).sType = "B" Then
                    'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                    '    If tmCurrDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                        ilBDE = gBinarySearchBDE(tmCurrDBE(ilDBE).iBdeCode, tgCurrBDE())
                        If ilBDE <> -1 Then
                            If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) = 0 Then
                                ilFound = True
                                tmCurrDBE(ilDBE).sType = ""
                            End If
                    '        Exit For
                        End If
                    'Next ilBDE
                End If
            Next ilDBE
            If Not ilFound Then
                mCompare = False
                Exit Function
            End If
        End If
    Next ilLoop
    For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
        If tmCurrDBE(ilDBE).sType = "B" Then
            mCompare = False
            Exit Function
        End If
    Next ilDBE
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
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


Private Sub mMoveDEERecToCtrls()
    Dim llRow As Long
    Dim slStr As String
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
    Dim ilCTE As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slHours As String
    Dim llRet As Long
    
    mPopNNE
    ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
    ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
    llRow = grdLibEvents.FixedRows
    For ilLoop = 0 To UBound(tgCurrDEE) - 1 Step 1
        If llRow + 1 > grdLibEvents.Rows Then
            grdLibEvents.AddItem ""
        End If
        grdLibEvents.Row = llRow
        slStr = ""
        smCurrEBEStamp = ""
        Erase tmCurrEBE
        ilRet = gGetRecs_EBE_EventBusSel(smCurrEBEStamp, tgCurrDEE(ilLoop).lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrEBE())
        For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
            'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '    If tmCurrEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                ilBDE = gBinarySearchBDE(tmCurrEBE(ilEBE).iBdeCode, tgCurrBDE())
                If ilBDE <> -1 Then
                    slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName) & ","
            '        Exit For
                End If
            'Next ilBDE
        Next ilEBE
        If slStr <> "" Then
            slStr = Left$(slStr, Len(slStr) - 1)
        End If
        grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = slStr
        grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX) = ""
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If tgCurrDEE(ilLoop).iCceCode = tgCurrBusCCE(ilCCE).iCode Then
                grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX) = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_B.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrBusCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_B.AddItem Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
                    lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = tgCurrBusCCE(ilCCE).iCode
                End If
                Exit For
            End If
        Next ilCCE
        grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrDEE(ilLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = Trim$(tgCurrETE(ilETE).sName)
                Exit For
            End If
        Next ilETE
        grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrLengthInTenth(tgCurrDEE(ilLoop).lTime, False)
        grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = ""
        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            If tgCurrDEE(ilLoop).iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
                grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = Trim$(tgCurrStartTTE(ilTTE).sName)
                llRet = SendMessageByString(lbcTTE_S.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrStartTTE(ilTTE).sName))
                If llRet < 0 Then
                    lbcTTE_S.AddItem Trim$(tgCurrStartTTE(ilTTE).sName)
                    lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = tgCurrStartTTE(ilTTE).iCode
                End If
                Exit For
            End If
        Next ilTTE
        grdLibEvents.TextMatrix(llRow, FIXEDINDEX) = Trim$(tgCurrDEE(ilLoop).sFixedTime)
        grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = ""
        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            If tgCurrDEE(ilLoop).iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
                grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = Trim$(tgCurrEndTTE(ilTTE).sName)
                llRet = SendMessageByString(lbcTTE_E.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrEndTTE(ilTTE).sName))
                If llRet < 0 Then
                    lbcTTE_E.AddItem Trim$(tgCurrEndTTE(ilTTE).sName)
                    lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = tgCurrEndTTE(ilTTE).iCode
                End If
                Exit For
            End If
        Next ilTTE
        '11/24/04- Allow end type and Duration to co-exist
        'If (tgCurrDEE(ilLoop).lDuration > 0) Or (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) = "") Then
        If (tgCurrDEE(ilLoop).lDuration > 0) Then
            grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tgCurrDEE(ilLoop).lDuration, True)
        Else
            grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tgCurrDEE(ilLoop).lDuration, True)
        End If
        slHours = Trim$(tgCurrDEE(ilLoop).sHours)
        slStr = gHourMap(slHours)
        grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX) = slStr
        slStr = gDayMap(tgCurrDEE(ilLoop).sDays)
        grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX) = Trim$(slStr)
        grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = ""
        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            If tgCurrDEE(ilLoop).iMteCode = tgCurrMTE(ilMTE).iCode Then
                grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = Trim$(tgCurrMTE(ilMTE).sName)
                llRet = SendMessageByString(lbcMTE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrMTE(ilMTE).sName))
                If llRet < 0 Then
                    lbcMTE.AddItem Trim$(tgCurrMTE(ilMTE).sName)
                    lbcMTE.ItemData(lbcMTE.NewIndex) = tgCurrMTE(ilMTE).iCode
                End If
                Exit For
            End If
        Next ilMTE
        grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = ""
        'For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).iAudioAseCode = tmCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tgCurrDEE(ilLoop).iAudioAseCode, tmCurrASE())
            If ilASE <> -1 Then
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                        llRet = SendMessageByString(lbcASE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                        If llRet < 0 Then
                            lbcASE.AddItem Trim$(tgCurrANE(ilANE).sName)
                            lbcASE.ItemData(lbcASE.NewIndex) = tmCurrASE(ilASE).iCode
                        End If
                    End If
                'Next ilANE
        '        Exit For
            End If
        'Next ilASE
        grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = Trim$(tgCurrDEE(ilLoop).sAudioItemID)
        grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX) = Trim$(tgCurrDEE(ilLoop).sAudioISCI)
        grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = ""
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrDEE(ilLoop).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
                Exit For
            End If
        Next ilCCE
        grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = ""
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tgCurrDEE(ilLoop).iBkupAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                If llRet < 0 Then
                    lbcANE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilANE).iCode
                End If
        '        Exit For
            End If
        'Next ilANE
        grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = ""
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrDEE(ilLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
                Exit For
            End If
        Next ilCCE
        grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = ""
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tgCurrDEE(ilLoop).iProtAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                If llRet < 0 Then
                    lbcANE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilANE).iCode
                End If
        '        Exit For
            End If
        'Next ilANE
        grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX) = Trim$(tgCurrDEE(ilLoop).sProtItemID)
        grdLibEvents.TextMatrix(llRow, PROTISCIINDEX) = Trim$(tgCurrDEE(ilLoop).sProtISCI)
        grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = ""
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrDEE(ilLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
                Exit For
            End If
        Next ilCCE
        grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = ""
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).i1RneCode = tgCurrRNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tgCurrDEE(ilLoop).i1RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
                llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrRNE(ilRNE).sName))
                If llRet < 0 Then
                    lbcRNE.AddItem Trim$(tgCurrRNE(ilRNE).sName)
                    lbcRNE.ItemData(lbcRNE.NewIndex) = tgCurrRNE(ilRNE).iCode
                End If
        '        Exit For
            End If
        'Next ilRNE
        grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = ""
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).i2RneCode = tgCurrRNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tgCurrDEE(ilLoop).i2RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
                llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrRNE(ilRNE).sName))
                If llRet < 0 Then
                    lbcRNE.AddItem Trim$(tgCurrRNE(ilRNE).sName)
                    lbcRNE.ItemData(lbcRNE.NewIndex) = tgCurrRNE(ilRNE).iCode
                End If
        '        Exit For
            End If
        'Next ilRNE
        grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = ""
        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            If tgCurrDEE(ilLoop).iFneCode = tgCurrFNE(ilFNE).iCode Then
                grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = Trim$(tgCurrFNE(ilFNE).sName)
                llRet = SendMessageByString(lbcFNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrFNE(ilFNE).sName))
                If llRet < 0 Then
                    lbcFNE.AddItem Trim$(tgCurrFNE(ilFNE).sName)
                    lbcFNE.ItemData(lbcFNE.NewIndex) = tgCurrFNE(ilFNE).iCode
                End If
                Exit For
            End If
        Next ilFNE
        If tgCurrDEE(ilLoop).lSilenceTime > 0 Then
            grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = gLongToLength(tgCurrDEE(ilLoop).lSilenceTime, False)   'gLongToStrLengthInTenth(tgCurrDEE(ilLoop).lSilenceTime, False)
        Else
            grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = ""
        End If
        grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = ""
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tgCurrDEE(ilLoop).i1SceCode = tgCurrSCE(ilSCE).iCode Then
                grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
                Exit For
            End If
        Next ilSCE
        grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = ""
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tgCurrDEE(ilLoop).i2SceCode = tgCurrSCE(ilSCE).iCode Then
                grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
                Exit For
            End If
        Next ilSCE
        grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = ""
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tgCurrDEE(ilLoop).i3SceCode = tgCurrSCE(ilSCE).iCode Then
                grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
                Exit For
            End If
        Next ilSCE
        grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = ""
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tgCurrDEE(ilLoop).i4SceCode = tgCurrSCE(ilSCE).iCode Then
                grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
                Exit For
            End If
        Next ilSCE
        grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = ""
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tgCurrDEE(ilLoop).iStartNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
                llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrNNE(ilNNE).sName))
                If llRet < 0 Then
                    lbcNNE.AddItem Trim$(tgCurrNNE(ilNNE).sName)
                    lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilNNE).iCode
                End If
        '        Exit For
            End If
        'Next ilNNE
        grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = ""
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tgCurrDEE(ilLoop).iEndNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
                llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrNNE(ilNNE).sName))
                If llRet < 0 Then
                    lbcNNE.AddItem Trim$(tgCurrNNE(ilNNE).sName)
                    lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilNNE).iCode
                End If
        '        Exit For
            End If
        'Next ilNNE
        grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = ""
        If tgCurrDEE(ilLoop).l1CteCode > 0 Then
            For ilCTE = 0 To UBound(tmCurr1CTE_Name) - 1 Step 1
                If tmCurr1CTE_Name(ilCTE).lCteCode = tgCurrDEE(ilLoop).l1CteCode Then
                    grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = Trim$(tmCurr1CTE_Name(ilCTE).sComment)
                    Exit For
                End If
            Next ilCTE
            If grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = "" Then
                ilRet = gGetRec_CTE_CommtsTitle(tgCurrDEE(ilLoop).l1CteCode, "EngrLibDef- mMoveDEERecToCtrl for CTE", tmCTE)
                grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = Trim$(tmCTE.sComment)
                tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).sComment = tmCTE.sComment
                tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lCteCode = tmCTE.lCode
                tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDheCode = tgCurrDEE(ilLoop).lDheCode
                ReDim Preserve tmCurr1CTE_Name(0 To UBound(tmCurr1CTE_Name) + 1) As DEECTE
            End If
        End If
        grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = ""
        '7/8/11: Make T2 work like T1
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    If tgCurrDEE(ilLoop).l2CteCode = tgCurrCTE(ilCTE).lCode Then
        '        grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tgCurrCTE(ilCTE).sName)
        '        llRet = SendMessageByString(lbcCTE_2.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrCTE(ilCTE).sName))
        '        If llRet < 0 Then
        '            lbcCTE_2.AddItem Trim$(tgCurrCTE(ilCTE).sName)
        '            lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = tgCurrCTE(ilCTE).lCode
        '        End If
        '        Exit For
        '    End If
        'Next ilCTE
        If tgCurrDEE(ilLoop).l2CteCode > 0 Then
            For ilCTE = 0 To UBound(tmCurr2CTE_Name) - 1 Step 1
                If tmCurr2CTE_Name(ilCTE).lCteCode = tgCurrDEE(ilLoop).l2CteCode Then
                    grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tmCurr2CTE_Name(ilCTE).sComment)
                    Exit For
                End If
            Next ilCTE
            If grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = "" Then
                ilRet = gGetRec_CTE_CommtsTitle(tgCurrDEE(ilLoop).l2CteCode, "EngrLibDef- mMoveDEERecToCtrl for CTE", tmCTE)
                grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tmCTE.sComment)
                tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).sComment = tmCTE.sComment
                tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lCteCode = tmCTE.lCode
                tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDheCode = tgCurrDEE(ilLoop).lDheCode
                ReDim Preserve tmCurr2CTE_Name(0 To UBound(tmCurr2CTE_Name) + 1) As DEECTE
            End If
        End If
        
        If sgClientFields = "A" Then
            grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = Trim$(tgCurrDEE(ilLoop).sABCFormat)
            grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = Trim$(tgCurrDEE(ilLoop).sABCPgmCode)
            grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = Trim$(tgCurrDEE(ilLoop).sABCXDSMode)
            grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = Trim$(tgCurrDEE(ilLoop).sABCRecordItem)
        Else
            grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = ""
            grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = ""
            grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = ""
            grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = ""
        End If
        grdLibEvents.TextMatrix(llRow, PCODEINDEX) = tgCurrDEE(ilLoop).lCode
        grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "N"
        mSetColExportColor llRow
        llRow = llRow + 1
    Next ilLoop
    'If Modelling, add comments as if new
    If igLibCallType = 2 Then
        ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
        ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
    End If
    If llRow >= grdLibEvents.Rows Then
        grdLibEvents.AddItem ""
    End If
    'grdLibEvents.Redraw = True
    '8/26/11:  Moved Integral here in addition to ColumnWidth
    If Not bmIntegralSet Then
        bmIntegralSet = True
        gGrid_IntegralHeight grdLibEvents
        gGrid_FillWithRows grdLibEvents
        '8/26/11: Remove one row so row is not behind scroll bar
        grdLibEvents.Height = grdLibEvents.Height - grdLibEvents.RowHeight(0) '+ 30
    End If
End Sub

Private Sub mMoveDEECtrlsToRec()
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
    
    ReDim tmCurrDEE(1 To grdLibEvents.Rows - grdLibEvents.FixedRows) As DEE
    ReDim smT1Comment(1 To UBound(tmCurrDEE)) As String
    ReDim smT2Comment(1 To UBound(tmCurrDEE)) As String
    ReDim smEBuses(1 To UBound(tmCurrDEE)) As String
    llIndex = LBound(tmCurrDEE)
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            'Set Later- Bus selected
            smEBuses(llIndex) = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
            
            tmCurrDEE(llIndex).iCceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX))
            For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
                If StrComp(Trim$(tgCurrBusCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iCceCode = tgCurrBusCCE(ilCCE).iCode
                    Exit For
                End If
            Next ilCCE
            tmCurrDEE(llIndex).iEteCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iEteCode = tgCurrETE(ilETE).iCode
                    Exit For
                End If
            Next ilETE
            slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
            tmCurrDEE(llIndex).lTime = gStrLengthInTenthToLong(slStr)
            tmCurrDEE(llIndex).iStartTteCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX))
            For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
                If StrComp(Trim$(tgCurrStartTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iStartTteCode = tgCurrStartTTE(ilTTE).iCode
                    Exit For
                End If
            Next ilTTE
            tmCurrDEE(llIndex).sFixedTime = grdLibEvents.TextMatrix(llRow, FIXEDINDEX)
            tmCurrDEE(llIndex).iEndTteCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
            For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
                If StrComp(Trim$(tgCurrEndTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iEndTteCode = tgCurrEndTTE(ilTTE).iCode
                    Exit For
                End If
            Next ilTTE
            slStr = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
            tmCurrDEE(llIndex).lDuration = gStrLengthInTenthToLong(slStr)
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX))
            tmCurrDEE(llIndex).sHours = gCreateHourStr(slStr)
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AIRDAYSINDEX))
            tmCurrDEE(llIndex).sDays = gCreateDayStr(slStr)
            tmCurrDEE(llIndex).iMteCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, MATERIALINDEX))
            For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
                If StrComp(Trim$(tgCurrMTE(ilMTE).sName), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iMteCode = tgCurrMTE(ilMTE).iCode
                    Exit For
                End If
            Next ilMTE
            tmCurrDEE(llIndex).iAudioAseCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
            For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                            tmCurrDEE(llIndex).iAudioAseCode = tmCurrASE(ilASE).iCode
                        End If
                '        Exit For
                    End If
                'Next ilANE
                If tmCurrDEE(llIndex).iAudioAseCode <> 0 Then
                    Exit For
                End If
            Next ilASE
            tmCurrDEE(llIndex).sAudioItemID = grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)
            tmCurrDEE(llIndex).sAudioISCI = grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX)
            tmCurrDEE(llIndex).iAudioCceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX))
            For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode
                    Exit For
                End If
            Next ilCCE
            tmCurrDEE(llIndex).iBkupAneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                ilANE = gBinarySearchName(slStr, tgCurrANE_Name())
                If ilANE <> -1 Then
                    tmCurrDEE(llIndex).iBkupAneCode = tgCurrANE_Name(ilANE).iCode   'tgCurrANE(ilANE).iCode
            '        Exit For
                End If
            'Next ilANE
            tmCurrDEE(llIndex).iBkupCceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX))
            For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode
                    Exit For
                End If
            Next ilCCE
            tmCurrDEE(llIndex).iProtAneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                ilANE = gBinarySearchName(slStr, tgCurrANE_Name())
                If ilANE <> -1 Then
                    tmCurrDEE(llIndex).iProtAneCode = tgCurrANE_Name(ilANE).iCode   'tgCurrANE(ilANE).iCode
            '        Exit For
                End If
            'Next ilANE
            tmCurrDEE(llIndex).sProtItemID = grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)
            tmCurrDEE(llIndex).sProtISCI = grdLibEvents.TextMatrix(llRow, PROTISCIINDEX)
            tmCurrDEE(llIndex).iProtCceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX))
            For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode
                    Exit For
                End If
            Next ilCCE
            tmCurrDEE(llIndex).i1RneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX))
            'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                ilRNE = gBinarySearchName(slStr, tgCurrRNE_Name())
                If ilRNE <> -1 Then
                    tmCurrDEE(llIndex).i1RneCode = tgCurrRNE_Name(ilRNE).iCode  'tgCurrRNE(ilRNE).iCode
            '        Exit For
                End If
            'Next ilRNE
            tmCurrDEE(llIndex).i2RneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
            'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                ilRNE = gBinarySearchName(slStr, tgCurrRNE_Name())
                If ilRNE <> -1 Then
                    tmCurrDEE(llIndex).i2RneCode = tgCurrRNE_Name(ilRNE).iCode 'tgCurrRNE(ilRNE).iCode
            '        Exit For
                End If
            'Next ilRNE
            tmCurrDEE(llIndex).iFneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, FOLLOWINDEX))
            For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
                If StrComp(Trim$(tgCurrFNE(ilFNE).sName), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).iFneCode = tgCurrFNE(ilFNE).iCode
                    Exit For
                End If
            Next ilFNE
            slStr = grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX)
            tmCurrDEE(llIndex).lSilenceTime = gLengthToLong(slStr)  'gStrLengthInTenthToLong(slStr)
            tmCurrDEE(llIndex).i1SceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX))
            For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).i1SceCode = tgCurrSCE(ilSCE).iCode
                    Exit For
                End If
            Next ilSCE
             tmCurrDEE(llIndex).i2SceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
            For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).i2SceCode = tgCurrSCE(ilSCE).iCode
                    Exit For
                End If
            Next ilSCE
            tmCurrDEE(llIndex).i3SceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
            For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).i3SceCode = tgCurrSCE(ilSCE).iCode
                    Exit For
                End If
            Next ilSCE
            tmCurrDEE(llIndex).i4SceCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
            For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                    tmCurrDEE(llIndex).i4SceCode = tgCurrSCE(ilSCE).iCode
                    Exit For
                End If
            Next ilSCE
            tmCurrDEE(llIndex).iStartNneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX))
            'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                ilNNE = gBinarySearchName(slStr, tgCurrNNE_Name())
                If ilNNE <> -1 Then
                    tmCurrDEE(llIndex).iStartNneCode = tgCurrNNE_Name(ilNNE).iCode  'tgCurrNNE(ilNNE).iCode
            '        Exit For
                End If
            'Next ilNNE
            tmCurrDEE(llIndex).iEndNneCode = 0
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
            'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                ilNNE = gBinarySearchName(slStr, tgCurrNNE_Name())
                If ilNNE <> -1 Then
                    tmCurrDEE(llIndex).iEndNneCode = tgCurrNNE_Name(ilNNE).iCode  'tgCurrNNE(ilNNE).iCode
            '        Exit For
                End If
            'Next ilNNE
            'Set later
            smT1Comment(llIndex) = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
            '7/8/11: Make T2 work like T1
            'tmCurrDEE(llIndex).l2CteCode = 0
            'slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
            ''For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
            ''    If StrComp(Trim$(tgCurrCTE(ilCTE).sName), slStr, vbTextCompare) = 0 Then
            '    llCTE = gBinarySearchCTEName(slStr, tgCurr2CTE_Name())
            '    If llCTE <> -1 Then
            '        tmCurrDEE(llIndex).l2CteCode = tgCurr2CTE_Name(llCTE).lCode
            ''        Exit For
            '    End If
            ''Next ilCTE
            smT2Comment(llIndex) = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
            If sgClientFields = "A" Then
                tmCurrDEE(llIndex).sABCFormat = grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX)
                tmCurrDEE(llIndex).sABCPgmCode = grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX)
                tmCurrDEE(llIndex).sABCXDSMode = grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX)
                tmCurrDEE(llIndex).sABCRecordItem = grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX)
            Else
                tmCurrDEE(llIndex).sABCFormat = ""
                tmCurrDEE(llIndex).sABCPgmCode = ""
                tmCurrDEE(llIndex).sABCXDSMode = ""
                tmCurrDEE(llIndex).sABCRecordItem = ""
            End If
            tmCurrDEE(llIndex).sIgnoreConflicts = "N"
            tmCurrDEE(llIndex).sUnused = ""
            If Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) = "" Then
                grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
            End If
            tmCurrDEE(llIndex).lCode = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
            llIndex = llIndex + 1
        End If
    Next llRow
    ReDim Preserve tmCurrDEE(1 To llIndex) As DEE
    ReDim Preserve smT1Comment(1 To UBound(tmCurrDEE)) As String
    ReDim Preserve smT2Comment(1 To UBound(tmCurrDEE)) As String
    ReDim Preserve smEBuses(1 To UBound(tmCurrDEE)) As String
    
End Sub




Private Function mCompareDEE(llCode As Long, slBuses As String, slT1Comment As String, slT2Comment As String) As Integer
    Dim ilDEENew As Integer
    Dim ilDEEOld As Integer
    Dim ilEBE As Integer
    Dim slStr As String
    Dim ilBDE As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    If llCode > 0 Then
        For ilDEENew = LBound(tmCurrDEE) To UBound(tmCurrDEE) - 1 Step 1
            If llCode = tmCurrDEE(ilDEENew).lCode Then
                For ilDEEOld = 1 To UBound(tgCurrDEE) - 1 Step 1
                    If llCode = tgCurrDEE(ilDEEOld).lCode Then
                        'Compare fields
                        'Buses
                        smCurrEBEStamp = ""
                        Erase tmCurrEBE
                        ilRet = gGetRecs_EBE_EventBusSel(smCurrEBEStamp, llCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrEBE())
                        gParseCDFields slBuses, False, smBuses()
                        For ilLoop = LBound(smBuses) To UBound(smBuses) Step 1
                            slStr = Trim$(smBuses(ilLoop))
                            If slStr <> "" Then
                                For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
                                    'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                    '    If tmCurrEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                        ilBDE = gBinarySearchBDE(tmCurrEBE(ilEBE).iBdeCode, tgCurrBDE())
                                        If ilBDE <> -1 Then
                                            If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) <> 0 Then
                                                mCompareDEE = False
                                                Exit Function
                                            End If
                                            tmCurrEBE(ilEBE).lCode = -1
                                    '        Exit For
                                        End If
                                    'Next ilBDE
                                Next ilEBE
                            End If
                        Next ilLoop
                        For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
                            If tmCurrEBE(ilEBE).lCode > 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        Next ilEBE
                        
                        If tmCurrDEE(ilDEENew).iCceCode <> tgCurrDEE(ilDEEOld).iCceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iEteCode <> tgCurrDEE(ilDEEOld).iEteCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).lTime <> tgCurrDEE(ilDEEOld).lTime Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iStartTteCode <> tgCurrDEE(ilDEEOld).iStartTteCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).sFixedTime <> tgCurrDEE(ilDEEOld).sFixedTime Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iEndTteCode <> tgCurrDEE(ilDEEOld).iEndTteCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).lDuration <> tgCurrDEE(ilDEEOld).lDuration Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).sHours <> tgCurrDEE(ilDEEOld).sHours Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).sDays <> tgCurrDEE(ilDEEOld).sDays Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iMteCode <> tgCurrDEE(ilDEEOld).iMteCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iAudioAseCode <> tgCurrDEE(ilDEEOld).iAudioAseCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrDEE(ilDEENew).sAudioItemID, tgCurrDEE(ilDEEOld).sAudioItemID, vbTextCompare) <> 0 Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrDEE(ilDEENew).sAudioISCI, tgCurrDEE(ilDEEOld).sAudioISCI, vbTextCompare) <> 0 Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iAudioCceCode <> tgCurrDEE(ilDEEOld).iAudioCceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iBkupAneCode <> tgCurrDEE(ilDEEOld).iBkupAneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iBkupCceCode <> tgCurrDEE(ilDEEOld).iBkupCceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iProtAneCode <> tgCurrDEE(ilDEEOld).iProtAneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrDEE(ilDEENew).sProtItemID, tgCurrDEE(ilDEEOld).sProtItemID, vbTextCompare) <> 0 Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrDEE(ilDEENew).sProtISCI, tgCurrDEE(ilDEEOld).sProtISCI, vbTextCompare) <> 0 Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iProtCceCode <> tgCurrDEE(ilDEEOld).iProtCceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i1RneCode <> tgCurrDEE(ilDEEOld).i1RneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i2RneCode <> tgCurrDEE(ilDEEOld).i2RneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iFneCode <> tgCurrDEE(ilDEEOld).iFneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).lSilenceTime <> tgCurrDEE(ilDEEOld).lSilenceTime Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i1SceCode <> tgCurrDEE(ilDEEOld).i1SceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i2SceCode <> tgCurrDEE(ilDEEOld).i2SceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i3SceCode <> tgCurrDEE(ilDEEOld).i3SceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).i4SceCode <> tgCurrDEE(ilDEEOld).i4SceCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iStartNneCode <> tgCurrDEE(ilDEEOld).iStartNneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        If tmCurrDEE(ilDEENew).iEndNneCode <> tgCurrDEE(ilDEEOld).iEndNneCode Then
                            mCompareDEE = False
                            Exit Function
                        End If
                        'Comment
                        ilRet = gGetRec_CTE_CommtsTitle(tgCurrDEE(ilDEEOld).l1CteCode, "EngrLibDef- mCompaerDEE for CTE", tmCTE)
                        If ilRet Then
                            If StrComp(Trim$(slT1Comment), Trim$(tmCTE.sComment), vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        Else
                            If Trim$(slT1Comment) <> "" Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        End If
                        '7/8/11: Make t2 work like T1
                        'If tmCurrDEE(ilDEENew).l2CteCode <> tgCurrDEE(ilDEEOld).l2CteCode Then
                        '    mCompareDEE = False
                        '    Exit Function
                        'End If
                        ilRet = gGetRec_CTE_CommtsTitle(tgCurrDEE(ilDEEOld).l2CteCode, "EngrLibDef- mCompaerDEE for CTE", tmCTE)
                        If ilRet Then
                            If StrComp(Trim$(slT2Comment), Trim$(tmCTE.sComment), vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        Else
                            If Trim$(slT2Comment) <> "" Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        End If
                        If sgClientFields = "A" Then
                            If StrComp(tmCurrDEE(ilDEENew).sABCFormat, tgCurrDEE(ilDEEOld).sABCFormat, vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrDEE(ilDEENew).sABCPgmCode, tgCurrDEE(ilDEEOld).sABCPgmCode, vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrDEE(ilDEENew).sABCXDSMode, tgCurrDEE(ilDEEOld).sABCXDSMode, vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrDEE(ilDEENew).sABCRecordItem, tgCurrDEE(ilDEEOld).sABCRecordItem, vbTextCompare) <> 0 Then
                                mCompareDEE = False
                                Exit Function
                            End If
                        End If
                        mCompareDEE = True
                        Exit Function
                    End If
                Next ilDEEOld
                mCompareDEE = True
                Exit Function
            End If
        Next ilDEENew
    Else
        mCompareDEE = True
    End If
    
    
    
End Function

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    Dim llCode As Long
    
    mBranch = True
    If (lmEnableRow >= grdLib.FixedRows) And (lmEnableRow < grdLib.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdLib.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case NAMEINDEX
                    'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcDNE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrDayName.Show vbModal
                        sgCurrLibDNEStamp = ""
                        mPopDNE
                        lbcDNE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcDNE, CLng(grdLib.Height / 2)
                        If lbcDNE.Top + lbcDNE.Height > cmcCancel.Top Then
                            lbcDNE.Top = edcDropdown.Top - lbcDNE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                            grdLib.TextMatrix(grdLib.Row, DESCRIPTIONINDEX) = ""
                            llRow = gListBoxFind(lbcDNE, slStr)
                            If llRow > 0 Then
                                lbcDNE.ListIndex = llRow
                                edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                                llCode = lbcDNE.ItemData(lbcDNE.ListIndex)
                                For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
                                    If llCode = tgCurrLibDNE(ilLoop).lCode Then
                                        grdLib.TextMatrix(grdLib.Row, DESCRIPTIONINDEX) = Trim$(tgCurrLibDNE(ilLoop).sDescription)
                                        Exit For
                                    End If
                                Next ilLoop
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case SUBLIBNAMEINDEX
                    llRow = gListBoxFind(lbcDSE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrDaySubName.Show vbModal
                        sgCurrDSEStamp = ""
                        mPopDSE
                        lbcDSE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcDSE, CLng(grdLib.Height / 2)
                        If lbcDSE.Top + lbcDSE.Height > cmcCancel.Top Then
                            lbcDSE.Top = edcDropdown.Top - lbcDSE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcDSE, slStr)
                            If llRow > 0 Then
                                lbcDSE.ListIndex = llRow
                                edcDropdown.text = lbcDSE.List(lbcDSE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case BUSGROUPSINDEX
                    ReDim ilBusGroupSel(0 To 0) As Integer
                    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                        If lbcBGE.Selected(ilLoop) Then
                            ilBusGroupSel(UBound(ilBusGroupSel)) = lbcBGE.ItemData(ilLoop)
                            ReDim Preserve ilBusGroupSel(0 To UBound(ilBusGroupSel) + 1) As Integer
                        End If
                    Next ilLoop
                    igInitCallInfo = 1
                    sgInitCallName = ""
                    EngrBusGroup.Show vbModal
                    sgCurrBGEStamp = ""
                    mPopBGE
                    For ilLoop = 0 To UBound(ilBusGroupSel) - 1 Step 1
                        For ilBGE = 0 To lbcBGE.ListCount - 1 Step 1
                            If ilBusGroupSel(ilLoop) = lbcBGE.ItemData(ilBGE) Then
                                lbcBGE.Selected(ilBGE) = True
                                Exit For
                            End If
                        Next ilBGE
                    Next ilLoop
                    lbcBGE.Move pbcDefine.Left, cmcNone.Top + cmcNone.Height, pbcDefine.Width
                    gSetListBoxHeight lbcBGE, CLng(grdLibEvents.Height / 2)
                    If igReturnCallStatus = CALLDONE Then
                        mBranch = True
                    ElseIf igReturnCallStatus = CALLCANCELLED Then
                        mBranch = False
                    ElseIf igReturnCallStatus = CALLTERMINATED Then
                        mBranch = False
                    End If
                Case BUSESINDEX
                    ReDim ilBusGroupSel(0 To 0) As Integer
                    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                        If lbcBGE.Selected(ilLoop) Then
                            ilBusGroupSel(UBound(ilBusGroupSel)) = lbcBGE.ItemData(ilLoop)
                            ReDim Preserve ilBusGroupSel(0 To UBound(ilBusGroupSel) + 1) As Integer
                        End If
                    Next ilLoop
                    ReDim ilBusSel(0 To 0) As Integer
                    For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
                        If lbcBDE.Selected(ilLoop) Then
                            ilBusSel(UBound(ilBusSel)) = lbcBDE.ItemData(ilLoop)
                            ReDim Preserve ilBusSel(0 To UBound(ilBusSel) + 1) As Integer
                        End If
                    Next ilLoop
                    igInitCallInfo = 1
                    sgInitCallName = ""
                    EngrBus.Show vbModal
                    sgCurrBGEStamp = ""
                    mPopBGE
                    For ilLoop = 0 To UBound(ilBusGroupSel) - 1 Step 1
                        For ilBGE = 0 To lbcBGE.ListCount - 1 Step 1
                            If ilBusGroupSel(ilLoop) = lbcBGE.ItemData(ilBGE) Then
                                lbcBGE.Selected(ilBGE) = True
                                Exit For
                            End If
                        Next ilBGE
                    Next ilLoop
                    sgCurrBDEStamp = ""
                    mPopBDE
                    For ilLoop = 0 To UBound(ilBusSel) - 1 Step 1
                        For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                            If ilBusSel(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                                lbcBDE.Selected(ilBDE) = True
                                Exit For
                            End If
                        Next ilBDE
                    Next ilLoop
                    lbcBDE.Move pbcDefine.Left, cmcDefine.Top + cmcDefine.Height, pbcDefine.Width
                    gSetListBoxHeight lbcBDE, CLng(grdLibEvents.Height / 2)
                    If igReturnCallStatus = CALLDONE Then
                        mBranch = True
                    ElseIf igReturnCallStatus = CALLCANCELLED Then
                        mBranch = False
                    ElseIf igReturnCallStatus = CALLTERMINATED Then
                        mBranch = False
                    End If
                Case CONFLICTINDEX
                Case STATEINDEX
            End Select
        End If
    End If
    imDoubleClickName = False
End Function

Private Function mEBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    mEBranch = True
    bmInBranch = True
    If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEEnableCol
                Case BUSNAMEINDEX
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
                    llRow = gListBoxFind(lbcETE, slStr)
                    If (llRow < 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrEventType.Show vbModal
                        sgCurrETEStamp = ""
                        mPopETE
                        lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
                        If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
                            lbcETE.Top = edcEDropdown.Top - lbcETE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
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
                        smCurrASEStamp = ""
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

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrLibDef-mPopDNE", tgCurrLibDNE())
    lbcDNE.Clear
    For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
        If tgCurrLibDNE(ilLoop).sState = "A" Then
            lbcDNE.AddItem Trim$(tgCurrLibDNE(ilLoop).sName)
            lbcDNE.ItemData(lbcDNE.NewIndex) = tgCurrLibDNE(ilLoop).lCode
        End If
    Next ilLoop
    If igLibCallType <> 3 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(LIBRARYJOB) = 2) Then
            lbcDNE.AddItem "[New]", 0
            lbcDNE.ItemData(lbcDNE.NewIndex) = 0
        Else
            lbcDNE.AddItem "[View]", 0
            lbcDNE.ItemData(lbcDNE.NewIndex) = 0
        End If
    Else
        lbcDNE.AddItem "[View]", 0
        lbcDNE.ItemData(lbcDNE.NewIndex) = 0
    End If
End Sub

Private Sub mPopDSE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrLibDef-mPopDSE Day Subname", tgCurrDSE())
    lbcDSE.Clear
    For ilLoop = 0 To UBound(tgCurrDSE) - 1 Step 1
        If tgCurrDSE(ilLoop).sState = "A" Then
            lbcDSE.AddItem Trim$(tgCurrDSE(ilLoop).sName)
            lbcDSE.ItemData(lbcDSE.NewIndex) = tgCurrDSE(ilLoop).lCode
        End If
    Next ilLoop
    If igLibCallType <> 3 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(LIBRARYJOB) = 2) Then
            lbcDSE.AddItem "[New]", 0
            lbcDSE.ItemData(lbcDSE.NewIndex) = 0
        Else
            lbcDSE.AddItem "[View]", 0
            lbcDSE.ItemData(lbcDSE.NewIndex) = 0
        End If
    Else
        lbcDSE.AddItem "[View]", 0
        lbcDSE.ItemData(lbcDSE.NewIndex) = 0
    End If
End Sub

Private Sub mPopBGE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrLibDef-mPopBGE Bus Groups", tgCurrBGE())
    lbcBGE.Clear
    For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
        If tgCurrBGE(ilLoop).sState = "A" Then
            lbcBGE.AddItem Trim$(tgCurrBGE(ilLoop).sName)
            lbcBGE.ItemData(lbcBGE.NewIndex) = tgCurrBGE(ilLoop).iCode
        End If
    Next ilLoop
'    lbcBGE.AddItem "[None]", 0
'    lbcBGE.ItemData(lbcBGE.NewIndex) = 0
'    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
'        lbcBGE.AddItem "[New]", 0
'        lbcBGE.ItemData(lbcBGE.NewIndex) = 0
'    Else
'        lbcBGE.AddItem "[View]", 0
'        lbcBGE.ItemData(lbcBGE.NewIndex) = 0
'    End If
End Sub

Private Sub mPopBDE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrLibDef-mPopBDE Bus Definition", tgCurrBDE())
    lbcBDE.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        If tgCurrBDE(ilLoop).sState = "A" Then
            lbcBDE.AddItem Trim$(tgCurrBDE(ilLoop).sName)
            lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilLoop).iCode
        End If
    Next ilLoop
'    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
'        lbcBDE.AddItem "[New]", 0
'        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
'    Else
'        lbcBDE.AddItem "[View]", 0
'        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
'    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If edcLib.Visible Or pbcDefine.Visible Or edcDropdown.Visible Or pbcConflict.Visible Or pbcState.Visible Then
        Select Case grdLib.Col
            Case NAMEINDEX
                lbcDNE.Visible = False
            Case SUBLIBNAMEINDEX
                lbcDSE.Visible = False
            Case BUSGROUPSINDEX
                lbcBGE.Visible = False
            Case BUSESINDEX
                lbcBDE.Visible = False
        End Select
    End If
    If edcEvent.Visible Or pbcEDefine.Visible Or edcEDropdown.Visible Or pbcYN.Visible Then
        Select Case grdLibEvents.Col
            Case BUSCTRLINDEX
                lbcCCE_B.Visible = False
            Case EVENTTYPEINDEX
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

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrLibDef-mPopCCE_Audio Control Character", tgCurrAudioCCE())
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

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrLibDef-mPopCCE_Bus Control Character", tgCurrBusCCE())
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

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrLibDef-mPopTTE_StartType Start Type", tgCurrStartTTE())
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

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrLibDef-mPopTTE_EndType End Type", tgCurrEndTTE())
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
    Dim ilATE As Integer

    mPopANE
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", smCurrASEStamp, "EngrLibDef-mPopASE Audio Source", tmCurrASE())
    lbcASE.Clear
    For ilLoop = 0 To UBound(tmCurrASE) - 1 Step 1
        If tmCurrASE(ilLoop).sState = "A" Then
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tmCurrASE(ilLoop).iPriAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tmCurrASE(ilLoop).iPriAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    lbcASE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcASE.ItemData(lbcASE.NewIndex) = tmCurrASE(ilLoop).iCode
                    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                        If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                            tmCurrASE(ilLoop).sDescription = Trim$(tmCurrASE(ilLoop).sDescription) & "/" & Trim$(tgCurrATE(ilATE).sName)
                            Exit For
                        End If
                    Next ilATE
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

    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrLibDef-mPopSCE Silence Character", tgCurrSCE())
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

    Dim llDNECode As Long
    Dim slStr As String

    llDNECode = -1
    slStr = Trim$(grdLib.TextMatrix(grdLib.FixedRows, NAMEINDEX))
    For ilLoop = 0 To UBound(tgCurrLibDNE) - 1 Step 1
        If StrComp(Trim$(tgCurrLibDNE(ilLoop).sName), slStr, vbTextCompare) = 0 Then
            llDNECode = tgCurrLibDNE(ilLoop).lCode
            Exit For
        End If
    Next ilLoop

    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrLibDef-mPopNNE Netcue", tgCurrNNE())
    lbcNNE.Clear
    For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tgCurrNNE(ilLoop).sState = "A" Then
            If (tgCurrNNE(ilLoop).lDneCode = 0) Or (tgCurrNNE(ilLoop).lDneCode = llDNECode) Then
                lbcNNE.AddItem Trim$(tgCurrNNE(ilLoop).sName)
                lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilLoop).iCode
            End If
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

    'ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrLibDef-mPopCTE Title 2", tgCurrCTE())
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

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrLibDef-mPopANE Audio Audio Names", tgCurrANE())
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

Private Sub mPopETE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    imDefaultProgIndex = -1
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibDEF-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrLibDef-mPopETE Event Properties", tgCurrEPE())
    lbcETE.Clear
    For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilLoop).sState = "A" Then
            If tgCurrETE(ilLoop).sCategory <> "S" Then
                lbcETE.AddItem Trim$(tgCurrETE(ilLoop).sName)
                lbcETE.ItemData(lbcETE.NewIndex) = tgCurrETE(ilLoop).iCode
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcETE.ListCount - 1 Step 1
        If Trim$(lbcETE.List(ilLoop)) = "Program" Then
            imDefaultProgIndex = ilLoop
            Exit For
        End If
    Next ilLoop
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

    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrLibDef-mPopMTE Material Type", tgCurrMTE())
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

    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrLibDef-mPopRNE Relay", tgCurrRNE())
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

    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrLibDef-mPopFNE Follow", tgCurrFNE())
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




Private Sub mSetCTE(slComment As String, slType As String)
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    tmCTE.lCode = 0
    tmCTE.sComment = slComment
    tmCTE.sState = "A"
    tmCTE.sType = slType    '"DH" or "T1"
    tmCTE.sUsedFlag = "Y"
    tmCTE.iVersion = 0
    tmCTE.lOrigCteCode = 0
    tmCTE.sCurrent = "Y"
    'tmCTE.sEnteredDate = smNowDate
    'tmCTE.sEnteredTime = smNowTime
    tmCTE.sEnteredDate = Format(Now, sgShowDateForm) 'smNowDate
    tmCTE.sEnteredTime = Format(Now, sgShowTimeWSecForm) 'smNowTime
    tmCTE.iUieCode = tgUIE.iCode
    tmCTE.sUnused = ""

End Sub

Private Sub mSetDBE(slName As String, slType As String)
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    
    tmDBE.lCode = 0
    tmDBE.sType = slType
    tmDBE.lDheCode = tmDHE.lCode
    If slType = "G" Then
        tmDBE.iBdeCode = 0
        tmDBE.iBgeCode = 0
        For ilBGE = 0 To UBound(tgCurrBGE) - 1 Step 1
            If StrComp(Trim$(tgCurrBGE(ilBGE).sName), slName, vbTextCompare) = 0 Then
                tmDBE.iBgeCode = tgCurrBGE(ilBGE).iCode
                Exit For
            End If
        Next ilBGE
    Else
        tmDBE.iBgeCode = 0
        tmDBE.iBdeCode = 0
        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slName, vbTextCompare) = 0 Then
            ilBDE = gBinarySearchName(slName, tgCurrBDE_Name())
            If ilBDE <> -1 Then
                tmDBE.iBdeCode = tgCurrBDE_Name(ilBDE).iCode    'tgCurrBDE(ilBDE).iCode
        '        Exit For
            End If
        'Next ilBDE
    End If
    tmDBE.sUnused = ""
End Sub

Private Sub mSetEBE(slName As String, llDeeCode As Long)
    Dim ilBDE As Integer
    
    tmEBE.lCode = 0
    tmEBE.lDeeCode = llDeeCode
    'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
    '    If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slName, vbTextCompare) = 0 Then
        ilBDE = gBinarySearchName(slName, tgCurrBDE_Name())
        If ilBDE <> -1 Then
            tmEBE.iBdeCode = tgCurrBDE_Name(ilBDE).iCode    'tgCurrBDE(ilBDE).iCode
    '        Exit For
        End If
    'Next ilBDE
    tmEBE.sUnused = ""
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

Private Function mColOk(llRow As Long, llCol As Long) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    
    mColOk = True
    If grdLibEvents.ColWidth(grdLibEvents.Col) <= 0 Then
        mColOk = False
        Exit Function
    End If
    grdLibEvents.Row = llRow
    grdLibEvents.Col = llCol
    If grdLibEvents.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "U" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case grdLibEvents.Col
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mColOk = False
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mColOk = False
                                    End If
                                Case EVENTTYPEINDEX
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
                                Case AIRHOURSINDEX
                                Case AIRDAYSINDEX
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

Private Function mMinHeaderFieldsDefined() As Integer
    Dim llRow As Long
    
    llRow = grdLib.FixedRows
    If Trim$(grdLib.TextMatrix(llRow, DAYSINDEX)) = "" Then
        mMinHeaderFieldsDefined = False
        Exit Function
    End If
    If Trim$(grdLib.TextMatrix(llRow, HOURSINDEX)) = "" Then
        mMinHeaderFieldsDefined = False
        Exit Function
    End If
    If Trim$(grdLib.TextMatrix(llRow, BUSESINDEX)) = "" Then
        mMinHeaderFieldsDefined = False
        Exit Function
    End If
    mMinHeaderFieldsDefined = True
End Function

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
    'Check for conflicts
    llRow = grdLib.FixedRows
    If grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Both" Then
        Exit Function
    End If
    If Trim$(grdLib.TextMatrix(llRow, CODEINDEX)) = "" Then
        llDheCode = 0
    Else
        llDheCode = Val(grdLib.TextMatrix(llRow, CODEINDEX))
    End If
    slStartDate = Trim$(grdLib.TextMatrix(llRow, STARTDATEINDEX))
    If Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX)) <> "" Then
        slEndDate = Trim$(grdLib.TextMatrix(llRow, ENDDATEINDEX))
    Else
        slEndDate = "12/31/2069"
    End If
    If Not gIsDate(slStartDate) Then
        mCheckLibConflicts = True
        Exit Function
    End If
    If Not gIsDate(slEndDate) Then
        mCheckLibConflicts = True
        Exit Function
    End If
    ilCols(0) = ERRORFLAGINDEX
    ilCols(1) = EVENTTYPEINDEX
    ilCols(2) = AIRHOURSINDEX
    ilCols(3) = AIRDAYSINDEX
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
    ilCols(15) = PCODEINDEX
    'ilRet = gCheckConflicts("L", llDheCode, lmCurrentDHE, slStartDate, slEndDate, "", grdLibEvents, ilCols(), tmConflictList())
    ilRet = gConflictTableCheck("L", llDheCode, lmCurrentDHE, slStartDate, slEndDate, "", grdLibEvents, ilCols(), tmConflictList())
    If ilRet Then
        mCheckLibConflicts = True
    End If
End Function



Private Function mCheckEventConflicts()
    Dim llRow1 As Long
    Dim llRow2 As Long
    Dim llRowT1 As Long
    Dim llRowT2 As Long
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
    Dim slProtItemID1 As String
    Dim slBkupItemID1 As String
    Dim slPriItemID2 As String
    Dim slProtItemID2 As String
    Dim slBkupItemID2 As String
    Dim ilError As Integer
    Dim ilStartConflictIndex As Integer
    Dim ilConflictIndex As Integer
    Dim llUpper As Long
    Dim ilCheckWrapAround As Integer
    Dim slType As String
    Dim slTestHours1 As String
    Dim llRow1StartTime As Long
    Dim llRow1EndTime As Long
    Dim llOffsetEventStartTime As Long
    Dim llOffsetEventEndTime As Long
    
    mCheckEventConflicts = False
    ilCheckWrapAround = True
    slStr = Trim$(grdLibEvents.TextMatrix(llRow1, AIRDAYSINDEX))
    slDays1 = gCreateDayStr(slStr)
    If InStr(1, slDays1, "YY", vbTextCompare) = 0 Then
        If (Left(slDays1, 1) <> "Y") Or (right(slDays1, 1) <> "Y") Then
            ilCheckWrapAround = False
        End If
    End If
    ReDim tmConflictTest(1 To 1) As CONFLICTTEST
    llUpper = 1
    For llRow1 = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If (Trim$(grdLibEvents.TextMatrix(llRow1, EVENTTYPEINDEX)) <> "") Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow1, AIRHOURSINDEX))
            slHours1 = gCreateHourStr(slStr)
            slStr = Trim$(grdLibEvents.TextMatrix(llRow1, AIRDAYSINDEX))
            slDays1 = gCreateDayStr(slStr)
            slTestHours1 = slHours1
            slStr = grdLibEvents.TextMatrix(llRow1, TIMEINDEX)
            llOffsetEventStartTime = gStrLengthInTenthToLong(slStr)
            slStr = grdLibEvents.TextMatrix(llRow1, DURATIONINDEX)
            llOffsetEventEndTime = llOffsetEventStartTime + gStrLengthInTenthToLong(slStr)  ' - 1
            If llOffsetEventEndTime < llOffsetEventStartTime Then
                llOffsetEventEndTime = llOffsetEventStartTime
            End If
            slPriAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, AUDIONAMEINDEX))
            slProtAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, PROTNAMEINDEX))
            slBkupAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, BACKUPNAMEINDEX))
            For ilHour1 = 1 To 24 Step 1
                If (Mid$(slTestHours1, ilHour1, 1) = "Y") Then
                    llRow1StartTime = 36000 * (ilHour1 - 1) + llOffsetEventStartTime
                    llRow1EndTime = 36000 * (ilHour1 - 1) + llOffsetEventEndTime
                    mCreateBusRecs ilCheckWrapAround, llRow1, "B", llRow1StartTime, llRow1EndTime, slDays1, tmConflictTest()
                    mCreateAudioRecs ilCheckWrapAround, llRow1, "1", slPriAudio1, llRow1StartTime, llRow1EndTime, slDays1, tmConflictTest()
                    mCreateAudioRecs ilCheckWrapAround, llRow1, "2", slProtAudio1, llRow1StartTime, llRow1EndTime, slDays1, tmConflictTest()
                    mCreateAudioRecs ilCheckWrapAround, llRow1, "3", slBkupAudio1, llRow1StartTime, llRow1EndTime, slDays1, tmConflictTest()
                End If
            Next ilHour1
        End If
    Next llRow1
    If UBound(tmConflictTest) > LBound(tmConflictTest) Then
        For llRowT1 = 1 To UBound(tmConflictTest) - 1 Step 1
            llRow1 = tmConflictTest(llRowT1).lRow
            slStr = Trim$(grdLibEvents.TextMatrix(llRow1, EVENTTYPEINDEX))
            If slStr <> "" Then
                slEvtType1 = slStr
                llStartTime1 = tmConflictTest(llRowT1).lEventStartTime
                llEndTime1 = tmConflictTest(llRowT1).lEventEndTime
                slDays1 = tmConflictTest(llRowT1).sDays
                slPriAudio1 = grdLibEvents.TextMatrix(llRow1, AUDIONAMEINDEX)
                slProtAudio1 = grdLibEvents.TextMatrix(llRow1, PROTNAMEINDEX)
                slBkupAudio1 = grdLibEvents.TextMatrix(llRow1, BACKUPNAMEINDEX)
                slPriItemID1 = grdLibEvents.TextMatrix(llRow1, AUDIOITEMIDINDEX)
                slProtItemID1 = grdLibEvents.TextMatrix(llRow1, PROTITEMIDINDEX)
                slBkupItemID1 = grdLibEvents.TextMatrix(llRow1, AUDIOITEMIDINDEX)
                slBuses1 = Trim$(grdLibEvents.TextMatrix(llRow1, BUSNAMEINDEX))
                gParseCDFields slBuses1, False, slBus1()
                For ilBus1 = LBound(slBus1) To UBound(slBus1) Step 1
                    For ilDay1 = 1 To 7 Step 1
                        If Mid$(slDays1, ilDay1, 1) = "Y" Then
                            'For llRow2 = llRow1 + 1 To grdLibEvents.Rows - 1 Step 1
                            For llRowT2 = 1 To UBound(tmConflictTest) - 1 Step 1
                                llRow2 = tmConflictTest(llRowT2).lRow
                                slStr = Trim$(grdLibEvents.TextMatrix(llRow2, EVENTTYPEINDEX))
                                If (slStr <> "") And (llRow1 <> llRow2) Then
                                    slEvtType2 = slStr
                                    llStartTime2 = tmConflictTest(llRowT2).lEventStartTime
                                    llEndTime2 = tmConflictTest(llRowT2).lEventEndTime
                                    slDays2 = tmConflictTest(llRowT2).sDays
                                    slPriAudio2 = grdLibEvents.TextMatrix(llRow2, AUDIONAMEINDEX)
                                    slProtAudio2 = grdLibEvents.TextMatrix(llRow2, PROTNAMEINDEX)
                                    slBkupAudio2 = grdLibEvents.TextMatrix(llRow2, BACKUPNAMEINDEX)
                                    slPriItemID2 = grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX)
                                    slProtItemID2 = grdLibEvents.TextMatrix(llRow2, PROTITEMIDINDEX)
                                    slBkupItemID2 = grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX)
                                    ilError = False
                                    ilConflictIndex = UBound(tmConflictList)
                                    tmConflictList(ilConflictIndex).sType = "E"
                                    tmConflictList(ilConflictIndex).sStartDate = ""
                                    tmConflictList(ilConflictIndex).sEndDate = ""
                                    tmConflictList(ilConflictIndex).lIndex = llRow2
                                    tmConflictList(ilConflictIndex).iNextIndex = -1
                                    If Mid$(slDays2, ilDay1, 1) = "Y" Then
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX))
                                        slBuses2 = Trim$(grdLibEvents.TextMatrix(llRow2, BUSNAMEINDEX))
                                        gParseCDFields slBuses2, False, slBus2()
                                        If (tmConflictTest(llRowT1).sType = "B") And (tmConflictTest(llRowT2).sType = "B") Then
                                            For ilBus2 = LBound(slBus2) To UBound(slBus2) Step 1
                                                If StrComp(slBus1(ilBus1), slBus2(ilBus2), vbTextCompare) = 0 Then
                                                    'If (llEndTime2 >= llStartTime1) And (llStartTime2 <= llEndTime1) Then
                                                    If (llEndTime2 > llStartTime1) And (llStartTime2 < llEndTime1) Or (llStartTime1 = llStartTime2) Then
                                                        grdLibEvents.Row = llRow2
                                                        grdLibEvents.Col = TIMEINDEX
                                                        grdLibEvents.CellForeColor = vbRed
                                                        grdLibEvents.Col = DURATIONINDEX
                                                        grdLibEvents.CellForeColor = vbRed
                                                        If Not ilError Then
                                                            grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                        End If
                                                        ilError = True
                                                        mCheckEventConflicts = True
                                                    End If
                                                End If
                                            Next ilBus2
                                        End If
                                        For ilBus2 = LBound(slBus2) To UBound(slBus2) Step 1
                                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "1") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slPriAudio2, slPriItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "2") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slProtAudio2, slPriItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "3") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slBkupAudio2, slPriItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "1") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slPriAudio2, slProtItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "2") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slProtAudio2, slProtItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "3") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slBkupAudio2, slProtItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "1") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slPriAudio2, slBkupItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = AUDIONAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "2") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slProtAudio2, slBkupItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = PROTNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "3") Then
                                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slBkupAudio2, slBkupItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBus1(ilBus1), slBus2(ilBus2)) Then
                                                    grdLibEvents.Row = llRow1
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    grdLibEvents.Row = llRow2
                                                    grdLibEvents.Col = BACKUPNAMEINDEX
                                                    grdLibEvents.CellForeColor = vbRed
                                                    If Not ilError Then
                                                        grdLibEvents.TextMatrix(llRow1, ERRORFLAGINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                                    End If
                                                    ilError = True
                                                    mCheckEventConflicts = True
                                                End If
                                            End If
                                        Next ilBus2
                                    End If
                                End If
                            Next llRowT2

                        End If
                    Next ilDay1
                Next ilBus1
            End If
        Next llRowT1
    End If
    Erase tmConflictTest
End Function

Private Sub mLoadCTE_1()
    Dim llRow As Long
    Dim slStr As String
    
    lbcCTE_1.Clear
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
            If slStr <> "" Then
                If gListBoxFind(lbcCTE_1, slStr, True) < 0 Then
                    lbcCTE_1.AddItem slStr
                End If
            End If
        End If
    Next llRow
End Sub

Private Sub mLoadCTE_2()
    Dim llRow As Long
    Dim slStr As String
    
    lbcCTE_2.Clear
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
            If slStr <> "" Then
                If gListBoxFind(lbcCTE_2, slStr, True) < 0 Then
                    lbcCTE_2.AddItem slStr
                End If
            End If
        End If
    Next llRow
End Sub


Private Sub mSetUsedFlags(tlDEE As DEE)
    Dim ilRet As Integer
    
    ilRet = gPutUpdate_ANE_UsedFlag(tlDEE.iBkupAneCode, tgCurrANE())
    ilRet = gPutUpdate_ANE_UsedFlag(tlDEE.iProtAneCode, tgCurrANE())
    ilRet = gPutUpdate_ASE_UsedFlag(tlDEE.iAudioAseCode, tmCurrASE())
    ilRet = gPutUpdate_CCE_UsedFlag(tlDEE.iAudioCceCode, tgCurrAudioCCE())
    ilRet = gPutUpdate_CCE_UsedFlag(tlDEE.iBkupCceCode, tgCurrAudioCCE())
    ilRet = gPutUpdate_CCE_UsedFlag(tlDEE.iProtCceCode, tgCurrAudioCCE())
    ilRet = gPutUpdate_CCE_UsedFlag(tlDEE.iCceCode, tgCurrBusCCE())
    '7/8/11: Make T2 work like T1
    'ilRet = gPutUpdate_CTE_UsedFlag(tlDEE.l2CteCode, tgCurrCTE(), hmCTE)
    ilRet = gPutUpdate_ETE_UsedFlag(tlDEE.iEteCode, tgCurrETE())
    ilRet = gPutUpdate_FNE_UsedFlag(tlDEE.iFneCode, tgCurrFNE())
    ilRet = gPutUpdate_MTE_UsedFlag(tlDEE.iMteCode, tgCurrMTE())
    ilRet = gPutUpdate_NNE_UsedFlag(tlDEE.iEndNneCode, tgCurrNNE())
    ilRet = gPutUpdate_NNE_UsedFlag(tlDEE.iStartNneCode, tgCurrNNE())
    ilRet = gPutUpdate_RNE_UsedFlag(tlDEE.i1RneCode, tgCurrRNE())
    ilRet = gPutUpdate_RNE_UsedFlag(tlDEE.i2RneCode, tgCurrRNE())
    ilRet = gPutUpdate_SCE_UsedFlag(tlDEE.i1SceCode, tgCurrSCE())
    ilRet = gPutUpdate_SCE_UsedFlag(tlDEE.i2SceCode, tgCurrSCE())
    ilRet = gPutUpdate_SCE_UsedFlag(tlDEE.i3SceCode, tgCurrSCE())
    ilRet = gPutUpdate_SCE_UsedFlag(tlDEE.i4SceCode, tgCurrSCE())
    ilRet = gPutUpdate_TTE_UsedFlag(tlDEE.iEndTteCode, tgCurrEndTTE())
    ilRet = gPutUpdate_TTE_UsedFlag(tlDEE.iStartTteCode, tgCurrStartTTE())
End Sub

Private Sub mInitReplaceInfo()
    Dim ilUpper As Integer
    ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If ((tgUsedSumEPE.sAudioName <> "N") Or (tgUsedSumEPE.sProtAudioName <> "N") Or (tgUsedSumEPE.sBkupAudioName <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Name"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgReplaceFields(ilUpper).sListFile = "ANE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sAudioName
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sBus <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgReplaceFields(ilUpper).sListFile = "BDE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sBus
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sFollow <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Follow"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgReplaceFields(ilUpper).sListFile = "FNE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sFollow
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sMaterialType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Material"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgReplaceFields(ilUpper).sListFile = "MTE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sMaterialType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sStartNetcue <> "N") Or (tgUsedSumEPE.sStopNetcue <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Netcue"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgReplaceFields(ilUpper).sListFile = "NNE"
        If (tgManSumEPE.sStartNetcue = "Y") Or (tgManSumEPE.sStopNetcue = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sRelay1 <> "N") Or (tgUsedSumEPE.sRelay2 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Relay"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgReplaceFields(ilUpper).sListFile = "RNE"
        If (tgManSumEPE.sRelay1 = "Y") Or (tgManSumEPE.sRelay2 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sStartType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Start Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTES"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sStartType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sEndType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "End Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTEE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sEndType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioControl <> "N") Or (tgUsedSumEPE.sProtAudioControl <> "N") Or (tgUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEA"
        If (tgManSumEPE.sAudioControl = "Y") Or (tgManSumEPE.sProtAudioControl = "Y") Or (tgManSumEPE.sBkupAudioControl = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sBusControl <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEB"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sBusControl
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sTitle2 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 2"
        '7/8/11: Make T2 work like T1
        'tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE2"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sTitle2
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sTitle1 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 1"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title1", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE1"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sTitle1
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sSilence1 <> "N") Or (tgUsedSumEPE.sSilence2 <> "N") Or (tgUsedSumEPE.sSilence3 <> "N") Or (tgUsedSumEPE.sSilence4 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Silence Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgReplaceFields(ilUpper).sListFile = "SCE"
        If (tgManSumEPE.sSilence1 = "Y") Or (tgManSumEPE.sSilence2 = "Y") Or (tgManSumEPE.sSilence3 = "Y") Or (tgManSumEPE.sSilence4 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgUsedSumEPE.sFixedTime <> "N" Then
        tgReplaceFields(ilUpper).sFieldName = "Fixed Time"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = 1
        tgReplaceFields(ilUpper).sListFile = "FTYN"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sFixedTime
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioItemID <> "N") Or (tgUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Item ID"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("AudioItemID")
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgManSumEPE.sAudioItemID = "Y") Or (tgManSumEPE.sProtAudioItemID = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioISCI <> "N") Or (tgUsedSumEPE.sProtAudioISCI <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "ISCI"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("AudioISCI")
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgManSumEPE.sAudioISCI = "Y") Or (tgManSumEPE.sProtAudioISCI = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (sgClientFields = "A") Then
        If (tgUsedSumEPE.sABCFormat <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Format"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCFormat")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgManSumEPE.sABCFormat = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgUsedSumEPE.sABCPgmCode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Pgm Code"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCPgmCode")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgManSumEPE.sABCPgmCode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgUsedSumEPE.sABCXDSMode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC XDS Mode"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCXDSMODE")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgManSumEPE.sABCXDSMode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgUsedSumEPE.sABCRecordItem <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Recd Item"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCRecordItem")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgManSumEPE.sABCRecordItem = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
    End If
End Sub

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
    Dim slStr As String
    Dim ilBus As Integer
    
    ReDim tgYNMatchList(0 To 2) As MATCHLIST
    tgYNMatchList(0).sValue = "Y"
    tgYNMatchList(0).lValue = 0
    tgYNMatchList(1).sValue = "N"
    tgYNMatchList(1).lValue = 1
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
    'ReDim tgUsedT2CTE(0 To 0) As CTE
    ReDim tgT1MatchList(0 To 0) As MATCHLIST
    ReDim tgT2MatchList(0 To 0) As MATCHLIST
    For llLoop = LBound(tmCurrDEE) To UBound(tmCurrDEE) - 1 Step 1
        slStr = grdLib.TextMatrix(grdLib.FixedRows, BUSESINDEX)
        gParseCDFields slStr, False, smBuses()
        For ilBus = LBound(smBuses) To UBound(smBuses) Step 1
            slStr = Trim$(smBuses(ilBus))
            For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                If StrComp(slStr, Trim$(tgCurrBDE(ilBDE).sName), vbTextCompare) = 0 Then
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
                    Exit For
                End If
            Next ilBDE
        Next ilBus
        'For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
        '    If tmCurrDEE(llLoop).iAudioAseCode = tmCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tmCurrDEE(llLoop).iAudioAseCode, tmCurrASE())
            If ilASE <> -1 Then
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
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
        '    If tmCurrDEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrDEE(llLoop).iProtAneCode, tgCurrANE())
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
        '    If tmCurrDEE(llLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrDEE(llLoop).iBkupAneCode, tgCurrANE())
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
            If tmCurrDEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
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
            If tmCurrDEE(llLoop).iFneCode = tgCurrFNE(ilFNE).iCode Then
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
            If tmCurrDEE(llLoop).iMteCode = tgCurrMTE(ilMTE).iCode Then
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
        '    If tmCurrDEE(llLoop).iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrDEE(llLoop).iStartNneCode, tgCurrNNE())
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
        '    If tmCurrDEE(llLoop).iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrDEE(llLoop).iEndNneCode, tgCurrNNE())
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
        '    If tmCurrDEE(llLoop).i1RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrDEE(llLoop).i1RneCode, tgCurrRNE())
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
        '    If tmCurrDEE(llLoop).i2RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrDEE(llLoop).i2RneCode, tgCurrRNE())
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
            If tmCurrDEE(llLoop).iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
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
            If tmCurrDEE(llLoop).iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
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
            If tmCurrDEE(llLoop).iCceCode = tgCurrBusCCE(ilCCE).iCode Then
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
            If tmCurrDEE(llLoop).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
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
            If tmCurrDEE(llLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
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
            If tmCurrDEE(llLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
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
            If tmCurrDEE(llLoop).i1SceCode = tgCurrSCE(ilSCE).iCode Then
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
            If tmCurrDEE(llLoop).i2SceCode = tgCurrSCE(ilSCE).iCode Then
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
            If tmCurrDEE(llLoop).i3SceCode = tgCurrSCE(ilSCE).iCode Then
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
            If tmCurrDEE(llLoop).i4SceCode = tgCurrSCE(ilSCE).iCode Then
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
        '7/8/11:  Make T2 work like T1
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    If tmCurrDEE(llLoop).l2CteCode = tgCurrCTE(ilCTE).lCode Then
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
End Sub

Private Sub mReplaceValues()
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilReplace As Integer
    Dim ilField As Integer
    Dim ilFieldType As Integer
    Dim slGridBuses As String
    Dim slGridHours As String
    Dim ilGLoop As Integer
    Dim ilRLoop As Integer
    Dim ilBusMatch As Integer
    Dim ilHourMatch As Integer
    Dim slReplaceBuses As String
    Dim slReplaceHours As String
    Dim slFileName As String
    Dim ilColumn As Integer
    Dim ilSet As Integer
    Dim slNewValue As String
    Dim slOldValue As String
    Dim ilAllBusesMatch As Integer
    Dim ilAllHoursMatch As Integer
    Dim slFromHours As String
    Dim slToHours As String
    Dim llFromRow As Long
    Dim llToRow As Long
    Dim ilFieldChanged As Integer
    Dim ilPass As Integer
    Dim ilSplit As Integer
    Dim ilETE As Integer
    Dim ilCol(0 To 3) As Integer
    
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            grdLibEvents.Row = llRow
            For ilColumn = EVENTTYPEINDEX To imMaxCols Step 1
                grdLibEvents.Col = ilColumn
                If grdLibEvents.CellForeColor <> vbRed Then
                    If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                        grdLibEvents.CellForeColor = vbBlue
                    Else
                        grdLibEvents.CellForeColor = vbBlack
                    End If
                End If
            Next ilColumn
        End If
    Next llRow
    For ilPass = 0 To 1 Step 1
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
                smGridRow(0) = ""
                For ilColumn = EVENTTYPEINDEX To imMaxCols Step 1
                     smGridRow(ilColumn) = Trim$(grdLibEvents.TextMatrix(llRow, ilColumn))
                Next ilColumn
                slGridBuses = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
                slGridHours = gCreateHourStr(Trim$(grdLibEvents.TextMatrix(llRow, AIRHOURSINDEX)))
                ilSplit = False
                For ilReplace = LBound(tgLibReplaceValues) To UBound(tgLibReplaceValues) - 1 Step 1
                    For ilField = LBound(tgReplaceFields) To UBound(tgReplaceFields) - 1 Step 1
                        If tgReplaceFields(ilField).sFieldName = tgLibReplaceValues(ilReplace).sFieldName Then
                            ilFieldType = tgReplaceFields(ilField).iFieldType
                            slFileName = tgReplaceFields(ilField).sListFile
                            'Check if Bus and Hour filter matched
                            slReplaceBuses = tgLibReplaceValues(ilReplace).sBuses
                            
                            ilBusMatch = 0
                            gParseCDFields slGridBuses, False, smGridValues()
                            gParseCDFields slReplaceBuses, False, smReplaceValues()
                            For ilGLoop = LBound(smGridValues) To UBound(smGridValues) Step 1
                                For ilRLoop = LBound(smReplaceValues) To UBound(smReplaceValues) Step 1
                                    If StrComp(Trim$(smGridValues(ilGLoop)), Trim$(smReplaceValues(ilRLoop)), vbTextCompare) = 0 Then
                                        ilBusMatch = ilBusMatch + 1
                                        Exit For
                                    End If
                                Next ilRLoop
                            Next ilGLoop
                            If ilBusMatch = (UBound(smGridValues) - LBound(smGridValues) + 1) Then
                                ilAllBusesMatch = True
                            Else
                                ilAllBusesMatch = False
                            End If
                            
                            ilHourMatch = False
                            slReplaceHours = gCreateHourStr(Trim$(tgLibReplaceValues(ilReplace).sHours))
                            If StrComp(slGridHours, slReplaceHours, vbTextCompare) = 0 Then
                                ilHourMatch = True
                                ilAllHoursMatch = True
                            Else
                                ilAllHoursMatch = True
                                For ilGLoop = 1 To 24 Step 1
                                    If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "Y") Then
                                        ilHourMatch = True
                                    End If
                                    If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "N") Then
                                        ilAllHoursMatch = False
                                    End If
                                Next ilGLoop
                            End If
                            
                            If (ilBusMatch <> 0) And ilHourMatch Then
                                ilFieldChanged = False
                            
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
                                        slStr = Trim$(smGridRow(ilCol(ilSet)))    'Trim$(grdLibEvents.TextMatrix(llRow, ilCol(ilSet)))
                                        slOldValue = Trim$(tgLibReplaceValues(ilReplace).sOldValue)
                                        slNewValue = Trim$(tgLibReplaceValues(ilReplace).sNewValue)
                                        If (StrComp(slOldValue, slStr, vbTextCompare) = 0) Or ((slStr = "") And (StrComp(slOldValue, "[None]", vbTextCompare) = 0)) Then
                                            ilFieldChanged = True
                                            Exit For
                                        End If
                                    End If
                                Next ilSet
                                If ilFieldChanged Then
                                    If (Not ilAllBusesMatch) Or (Not ilAllHoursMatch) Then
                                        'Remove Buses and Hours from Current record and make new row with buses and hours
                                        If ilPass = 0 Then
                                            ilSplit = True
                                            llFromRow = llRow
                                            llToRow = llFromRow + 1
                                            grdLibEvents.AddItem "", llRow + 1
                                            For ilColumn = EVENTTYPEINDEX To imMaxCols Step 1
                                                grdLibEvents.TextMatrix(llToRow, ilColumn) = grdLibEvents.TextMatrix(llFromRow, ilColumn)
                                            Next ilColumn
                                            grdLibEvents.TextMatrix(llToRow, PCODEINDEX) = 0
                                            If (Not ilAllBusesMatch) Then
                                                grdLibEvents.TextMatrix(llFromRow, BUSNAMEINDEX) = ""
                                                grdLibEvents.TextMatrix(llToRow, BUSNAMEINDEX) = ""
                                                For ilGLoop = LBound(smGridValues) To UBound(smGridValues) Step 1
                                                    ilBusMatch = False
                                                    For ilRLoop = LBound(smReplaceValues) To UBound(smReplaceValues) Step 1
                                                        If StrComp(Trim$(smGridValues(ilGLoop)), Trim$(smReplaceValues(ilRLoop)), vbTextCompare) = 0 Then
                                                            ilBusMatch = True
                                                            Exit For
                                                        End If
                                                    Next ilRLoop
                                                    If ilBusMatch Then
                                                        If grdLibEvents.TextMatrix(llToRow, BUSNAMEINDEX) = "" Then
                                                            grdLibEvents.TextMatrix(llToRow, BUSNAMEINDEX) = smGridValues(ilGLoop)
                                                        Else
                                                            grdLibEvents.TextMatrix(llToRow, BUSNAMEINDEX) = grdLibEvents.TextMatrix(llToRow, BUSNAMEINDEX) & "," & smGridValues(ilGLoop)
                                                        End If
                                                    Else
                                                        If grdLibEvents.TextMatrix(llFromRow, BUSNAMEINDEX) = "" Then
                                                            grdLibEvents.TextMatrix(llFromRow, BUSNAMEINDEX) = smGridValues(ilGLoop)
                                                        Else
                                                            grdLibEvents.TextMatrix(llFromRow, BUSNAMEINDEX) = grdLibEvents.TextMatrix(llFromRow, BUSNAMEINDEX) & "," & smGridValues(ilGLoop)
                                                        End If
                                                    End If
                                                Next ilGLoop
                                            End If
                                            If (Not ilAllHoursMatch) Then
                                                slFromHours = String(24, "N")
                                                slToHours = String(24, "N")
                                                For ilGLoop = 1 To 24 Step 1
                                                    If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "Y") Then
                                                        Mid$(slToHours, ilGLoop, 1) = "Y"
                                                    End If
                                                    If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "N") Then
                                                        Mid$(slFromHours, ilGLoop, 1) = "Y"
                                                    End If
                                                Next ilGLoop
                                                grdLibEvents.TextMatrix(llFromRow, AIRHOURSINDEX) = gHourMap(slFromHours)
                                                grdLibEvents.TextMatrix(llToRow, AIRHOURSINDEX) = gHourMap(slToHours)
                                            End If
                                            Exit For
                                        End If
                                    Else
                                        If ilPass = 1 Then
                                            For ilSet = 0 To 3 Step 1
                                                If ilCol(ilSet) >= 0 Then
                                                    slStr = Trim$(smGridRow(ilCol(ilSet)))    'Trim$(grdLibEvents.TextMatrix(llRow, ilCol(ilSet)))
                                                    slOldValue = Trim$(tgLibReplaceValues(ilReplace).sOldValue)
                                                    slNewValue = Trim$(tgLibReplaceValues(ilReplace).sNewValue)
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
                                End If
                            End If
                        End If
                    Next ilField
                    If ilSplit Then
                        Exit For
                    End If
                Next ilReplace
            End If
        Next llRow
    Next ilPass
    mSetCommands
    
End Sub

Private Sub mPopATE()
    Dim ilRet As Integer
    
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
End Sub

Private Sub mESetFocus()
    Dim slStr As String
    Dim llColPos As Long
    Dim ilCol As Integer
    
    llColPos = 0
    For ilCol = 0 To grdLibEvents.Col - 1 Step 1
        If grdLibEvents.ColIsVisible(ilCol) And (grdLibEvents.ColWidth(ilCol) > 0) Then
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
            pbcEDefine.Move grdLibEvents.Left + grdLibEvents.ColPos(grdLibEvents.Col) + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            pbcEDefine.Width = gSetCtrlWidth("BusName", lmCharacterWidth, pbcEDefine.Width, 0)
            lbcBuses.Move pbcEDefine.Left, pbcEDefine.Top + pbcEDefine.Height, pbcEDefine.Width
            gSetListBoxHeight lbcBuses, CLng(grdLibEvents.Height / 2)
            If lbcBuses.Top + lbcBuses.Height > cmcCancel.Top Then
                lbcBuses.Top = pbcEDefine.Top - lbcBuses.Height
            End If
            pbcEDefine.Visible = True
            lbcBuses.Visible = True
            lbcBuses.SetFocus
        Case BUSCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("BusCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("BusCtrl", 6)
            imMaxColChars = gGetMaxChars("BusCtrl")
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
            edcEDropdown.MaxLength = Len(tgETE.sName)
            imMaxColChars = edcEDropdown.MaxLength
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
            If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
                lbcETE.Top = edcEDropdown.Top - lbcETE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            'lbcETE.Visible = True
            edcEDropdown.SetFocus
        Case TIMEINDEX
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, edcEvent.Width, 0)
'                edcEvent.MaxLength = gSetMaxChars("Time", 0)
'                imMaxColChars = gGetMaxChars("Time")
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            ltcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, ltcEvent.Width, 0)
            ltcEvent.Visible = True
            ltcEvent.SetFocus
        Case STARTTYPEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("StartType", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("StartType", 6)
            imMaxColChars = gGetMaxChars("StartType")
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
            smYN = grdLibEvents.text
            If (Trim$(smYN) = "") Or (smYN = "Missing") Then
                smYN = "N"
            End If
            lacHelp.Caption = "Indicate if this is a fixed time event. Enter Y or N or Mouse click to cycle value"
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case ENDTYPEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("EndType", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("EndType", 6)
            imMaxColChars = gGetMaxChars("EndType")
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
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, edcEvent.Width, 0)
'                edcEvent.MaxLength = gSetMaxChars("Duration", 0)
'                imMaxColChars = gGetMaxChars("Duration")
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            ltcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, ltcEvent.Width, 0)
            ltcEvent.Visible = True
            ltcEvent.SetFocus
        Case AIRHOURSINDEX
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.MaxLength = 0
'                imMaxColChars = 0
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            hpcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            'hpcEvent.MaxLength = 0
            hpcEvent.Visible = True
            hpcEvent.SetFocus
        Case AIRDAYSINDEX
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.MaxLength = 0
'                imMaxColChars = 0
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            dpcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            dpcEvent.MaxLength = 0
            dpcEvent.Visible = True
            dpcEvent.SetFocus
        Case MATERIALINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("Material", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("Material", 6)
            imMaxColChars = gGetMaxChars("Material")
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
            edcEDropdown.MaxLength = gSetMaxChars("AudioName", 0)
            imMaxColChars = gGetMaxChars("AudioName")
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
            edcEvent.MaxLength = gSetMaxChars("AudioItemID", 0)
            imMaxColChars = gGetMaxChars("AudioItemID")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case AUDIOISCIINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("AudioISCI", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.MaxLength = gSetMaxChars("AudioISCI", 0)
            imMaxColChars = gGetMaxChars("AudioISCI")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case AUDIOCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("AudioCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("AudioCtrl", 6)
            imMaxColChars = gGetMaxChars("AudioCtrl")
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
            edcEDropdown.MaxLength = gSetMaxChars("BkupName", 6)
            imMaxColChars = gGetMaxChars("BkupName")
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
            edcEDropdown.MaxLength = gSetMaxChars("BkupCtrl", 6)
            imMaxColChars = gGetMaxChars("BkupCtrl")
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
            edcEDropdown.MaxLength = gSetMaxChars("ProtName", 6)
            imMaxColChars = gGetMaxChars("ProtName")
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
            edcEvent.MaxLength = gSetMaxChars("ProtItemID", 0)
            imMaxColChars = gGetMaxChars("ProtItemID")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case PROTISCIINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ProtISCI", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.MaxLength = gSetMaxChars("ProtISCI", 0)
            imMaxColChars = gGetMaxChars("ProtISCI")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case PROTCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("ProtCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.MaxLength = gSetMaxChars("ProtCtrl", 6)
            imMaxColChars = gGetMaxChars("ProtCtrl")
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
            edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
            imMaxColChars = gGetMaxChars(slStr)
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
            edcEDropdown.MaxLength = gSetMaxChars("Follow", 6)
            imMaxColChars = gGetMaxChars("Follow")
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
'            edcEvent.MaxLength = gSetMaxChars("SilenceTime", 0)
'            imMaxColChars = gGetMaxChars("SilenceTime")
'            edcEvent.Text = grdLibEvents.Text
'            lacHelp.Caption = "Enter the allowed silence time of this event. Format is mm:ss"
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
            edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
            imMaxColChars = gGetMaxChars(slStr)
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
            edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
            imMaxColChars = gGetMaxChars(slStr)
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
            edcEDropdown.MaxLength = gSetMaxChars("Title1", 6)
            imMaxColChars = gGetMaxChars("Title1")
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
            edcEDropdown.MaxLength = gSetMaxChars("Title2", 6)
            imMaxColChars = gGetMaxChars("Title2")
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




Private Sub mShowConflictGrid()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slHours As String
    Dim tlDEE As DEE
    Dim tlSEE As SEE
    Dim tlDHE As DHE
    Dim tlDNE As DNE
    Dim tlDSE As DSE
    Dim slCurrEBEStamp As String
    Dim tlCurrEBE() As EBE
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilColumn As Integer
    
    If lmEEnableRow < grdLibEvents.FixedRows Then
        grdConflicts.Visible = False
        Exit Sub
    End If
    If lmConflictRow = lmEEnableRow Then
        If lmEEnableRow <> -1 Then
            grdConflicts.Visible = True
        End If
        Exit Sub
    End If
    gGrid_Clear grdConflicts, True
    slStr = grdLibEvents.TextMatrix(lmEEnableRow, ERRORFLAGINDEX)
    If slStr = "" Then
        lmConflictRow = -1
        grdConflicts.Visible = False
        Exit Sub
    End If
    If Val(slStr) <= 0 Then
        lmConflictRow = -1
        grdConflicts.Visible = False
        Exit Sub
    End If
    llRow = grdConflicts.FixedRows
    ilLoop = Val(slStr)
    Do
        If llRow + 1 > grdConflicts.Rows Then
            grdConflicts.AddItem ""
        End If
        grdConflicts.Row = llRow
        If tmConflictList(ilLoop).sType = "S" Then
            ilRet = gGetRec_SEE_ScheduleEvent(tmConflictList(ilLoop).lSeeCode, "EngrLibDef-gGetRec_SEE_ScheduleEvent", tlSEE)
            ilRet = gGetRec_DEE_DayEvent(tlSEE.lDeeCode, "EngrLibDef-gGetRec_DEE_DayEvent", tlDEE)
            ilRet = gGetRec_DHE_DayHeaderInfo(tlDEE.lDheCode, "EngrLibDef-gGetRec_DHE_DayHeaderInfo", tlDHE)
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
            Select Case Weekday(tmConflictList(ilLoop).sStartDate)
                Case vbMonday
                    slStr = "Mo"
                Case vbTuesday
                    slStr = "Tu"
                Case vbWednesday
                    slStr = "We"
                Case vbThursday
                    slStr = "Th"
                Case vbFriday
                    slStr = "Fr"
                Case vbSaturday
                    slStr = "Sa"
                Case vbSunday
                    slStr = "Su"
            End Select
            grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = Trim$(slStr)
            grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = ""
            grdConflicts.TextMatrix(llRow, CONFLICTHOURSINDEX) = gLongToStrTimeInTenth(tlSEE.lTime)
            If (tlSEE.lDuration > 0) Then
                grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = gLongToStrLengthInTenth(tlSEE.lDuration, True)
            Else
                grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = gLongToStrLengthInTenth(tlSEE.lDuration, True)    '""
            End If
            slStr = ""
            'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '    If tlSEE.iBdeCode = tgCurrBDE(ilBDE).iCode Then
                ilBDE = gBinarySearchBDE(tlSEE.iBdeCode, tgCurrBDE())
                If ilBDE <> -1 Then
                    slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName)
            '        Exit For
                End If
            'Next ilBDE
            grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = slStr
            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = ""
            'For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
            '    If tlSEE.iAudioAseCode = tmCurrASE(ilASE).iCode Then
                ilASE = gBinarySearchASE(tlSEE.iAudioAseCode, tmCurrASE())
                If ilASE <> -1 Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                        End If
                    'Next ilANE
            '        Exit For
                End If
            'Next ilASE
            grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tlSEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tlSEE.iBkupAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = Trim$(tgCurrANE(ilANE).sName)
            '        Exit For
                End If
            'Next ilANE
            grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tlSEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tlSEE.iProtAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = Trim$(tgCurrANE(ilANE).sName)
            '        Exit For
                End If
            'Next ilANE
            llRow = llRow + 1
        ElseIf (tmConflictList(ilLoop).sType = "L") Or (tmConflictList(ilLoop).sType = "T") Then
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
            'For ilASE = 0 To UBound(tmCurrASE) - 1 Step 1
            '    If tlDEE.iAudioAseCode = tmCurrASE(ilASE).iCode Then
                ilASE = gBinarySearchASE(tlDEE.iAudioAseCode, tmCurrASE())
                If ilASE <> -1 Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tmCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tmCurrASE(ilASE).iPriAneCode, tgCurrANE())
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
            grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = grdLib.TextMatrix(grdLib.FixedRows, NAMEINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = grdLib.TextMatrix(grdLib.FixedRows, SUBLIBNAMEINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTSTARTDATEINDEX) = ""
            grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = ""
            grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = ""
            grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = ""
            grdConflicts.TextMatrix(llRow, CONFLICTHOURSINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, Time)
            grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, DURATIONINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, BUSNAMEINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, AUDIONAMEINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, BACKUPNAMEINDEX)
            grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = grdLibEvents.TextMatrix(tmConflictList(ilLoop).lIndex, PROTNAMEINDEX)
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
    grdConflicts.Visible = True
    grdConflicts.Redraw = True
End Sub

Private Sub mHideConflictGrid()
    grdConflicts.Visible = False
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
            If Not mExportCol(llRow, llCol) Then
                grdLibEvents.CellForeColor = vbBlue
            Else
                grdLibEvents.CellForeColor = vbBlack
            End If
        Next llCol
    End If
    grdLibEvents.Col = llSvCol
    grdLibEvents.Row = llSvRow
End Sub

Private Sub mCreateAudioRecs(ilCheckWrapAround, llRow As Long, slType As String, slAudio As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim slEventDays As String
    Dim ilDay As Integer
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
                If (ilCheckWrapAround) Then
                    tlConflict(llUpper).lRow = llRow
                    tlConflict(llUpper).sType = slType
                    tlConflict(llUpper).lEventStartTime = 0
                    tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000 + llPostTime
                    slEventDays = String(7, "N")
                    For ilDay = 1 To 7 Step 1
                        If ilDay = 7 Then
                            If Mid$(slDays, ilDay, 1) = "Y" Then
                                Mid(slEventDays, 1, 1) = "Y"
                            End If
                        Else
                            If Mid$(slDays, ilDay, 1) = "Y" Then
                                Mid(slEventDays, ilDay + 1, 1) = "Y"
                            End If
                        End If
                    Next ilDay
                    tlConflict(llUpper).sDays = slEventDays
                    llUpper = llUpper + 1
                    ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
                End If
            End If
        Else
            If (ilCheckWrapAround) Then
                tlConflict(llUpper).lEventStartTime = 864000 + (llEventStartTime - llPreTime)
                tlConflict(llUpper).lEventEndTime = 864000
                slEventDays = String(7, "N")
                For ilDay = 7 To 1 Step -1
                    If ilDay = 1 Then
                        If Mid$(slDays, ilDay, 1) = "Y" Then
                            Mid(slEventDays, 7, 1) = "Y"
                        End If
                    Else
                        If Mid$(slDays, ilDay, 1) = "Y" Then
                            Mid(slEventDays, ilDay - 1, 1) = "Y"
                        End If
                    End If
                Next ilDay
                tlConflict(llUpper).sDays = slEventDays
                llUpper = llUpper + 1
                ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
            End If
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
        If (ilCheckWrapAround) Then
            tlConflict(llUpper).lRow = llRow
            tlConflict(llUpper).sType = slType
            tlConflict(llUpper).sDays = slDays
            tlConflict(llUpper).lEventStartTime = 0
            tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000 + llPostTime
            slEventDays = String(7, "N")
            For ilDay = 1 To 7 Step 1
                If ilDay = 7 Then
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, 1, 1) = "Y"
                    End If
                Else
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, ilDay + 1, 1) = "Y"
                    End If
                End If
            Next ilDay
            tlConflict(llUpper).sDays = slEventDays
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    End If
End Sub

Private Sub mCreateBusRecs(ilCheckWrapAround As Integer, llRow As Long, slType As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    Dim ilDay As Integer
    Dim slEventDays As String
    
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
        If (ilCheckWrapAround) Then
            tlConflict(llUpper).lRow = llRow
            tlConflict(llUpper).sType = slType
            tlConflict(llUpper).lEventStartTime = 0
            tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000
            slEventDays = String(7, "N")
            For ilDay = 1 To 7 Step 1
                If ilDay = 7 Then
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, 1, 1) = "Y"
                    End If
                Else
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, ilDay + 1, 1) = "Y"
                    End If
                End If
            Next ilDay
            tlConflict(llUpper).sDays = slEventDays
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    End If
End Sub


Private Sub mInitConflictTest()
    Dim llRow As Long
    Dim slIgnoreConflicts As String
    
    lmConflictRow = -1
    llRow = grdLib.FixedRows
    If grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Both" Then
        slIgnoreConflicts = "I"
    ElseIf grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Bus" Then
        slIgnoreConflicts = "B"
    ElseIf grdLib.TextMatrix(llRow, CONFLICTINDEX) = "Audio" Then
        slIgnoreConflicts = "A"
    Else
        slIgnoreConflicts = "N"
    End If
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            If Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) = "" Then
                grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
            End If
            If grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0" Then
                grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
            End If
            grdLibEvents.TextMatrix(llRow, ERRORFLAGINDEX) = "0"
            grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX) = slIgnoreConflicts
        End If
    Next llRow
End Sub

Private Function mCompareExtract(llRow As Long, tlExtract As SCHDEXTRACT) As Integer
    
    mCompareExtract = False
    If Left$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX), 1) = "P" Then
        If tlExtract.sEventType <> "P" Then
            Exit Function
        End If
    Else
        If tlExtract.sEventType <> "A" Then
            Exit Function
        End If
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX)), Trim$(tlExtract.sBusCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If (gStrLengthInTenthToLong(grdLibEvents.TextMatrix(llRow, TIMEINDEX)) <> tlExtract.lOffset) Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX)), Trim$(tlExtract.sStartType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)), Trim$(tlExtract.sEndType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If gStrLengthInTenthToLong(grdLibEvents.TextMatrix(llRow, DURATIONINDEX)) <> gStrLengthInTenthToLong(tlExtract.sDuration) Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, MATERIALINDEX)), Trim$(tlExtract.sMaterialType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX)), Trim$(tlExtract.sAudioName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)), Trim$(tlExtract.sAudioID), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX)), Trim$(tlExtract.sAudioISCI), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX)), Trim$(tlExtract.sAudioCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX)), Trim$(tlExtract.sBackupName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX)), Trim$(tlExtract.sBackupCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX)), Trim$(tlExtract.sProtName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)), Trim$(tlExtract.sProtItemID), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, PROTISCIINDEX)), Trim$(tlExtract.sProtISCI), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX)), Trim$(tlExtract.sProtCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX)), Trim$(tlExtract.sRelay1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX)), Trim$(tlExtract.sRelay2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, FOLLOWINDEX)), Trim$(tlExtract.sFollow), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If gLengthToLong(grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX)) <> gLengthToLong(tlExtract.sSilenceTime) Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), Trim$(tlExtract.sSilence1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX)), Trim$(tlExtract.sSilence2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX)), Trim$(tlExtract.sSilence3), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX)), Trim$(tlExtract.sSilence4), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX)), Trim$(tlExtract.sNetcue1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX)), Trim$(tlExtract.sNetcue2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If Left$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX), 1) = "P" Then
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX)), Trim$(tlExtract.sTitle1), vbTextCompare) <> 0 Then
            Exit Function
        End If
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX)), Trim$(tlExtract.sTitle2), vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, FIXEDINDEX)), Trim$(tlExtract.sFixedTime), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If sgClientFields = "A" Then
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX)), Trim$(tlExtract.sABCFormat), vbTextCompare) <> 0 Then
            Exit Function
        End If
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX)), Trim$(tlExtract.sABCPgmCode), vbTextCompare) <> 0 Then
            Exit Function
        End If
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX)), Trim$(tlExtract.sABCXDSMode), vbTextCompare) <> 0 Then
            Exit Function
        End If
        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX)), Trim$(tlExtract.sABCRecordItem), vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    mCompareExtract = True
End Function

Private Sub mSetBackgroundColor()
    Dim llRow As Long
    Dim llCol As Long
    Dim slDate As String
    Dim ilRet As Integer
        
    If (igLibCallType = 1) Then
        slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
        smNowDate = Format(gNow(), sgShowDateForm)
        If gDateValue(slDate) < gDateValue(smNowDate) Then
            grdLib.Col = STARTDATEINDEX
            grdLib.Row = grdLib.FixedRows
            grdLib.CellBackColor = LIGHTYELLOW
            imStartDateChangable = False
        End If
        If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
            imLimboAllowed = True
        End If
    ElseIf igLibCallType = 2 Then
        grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX) = ""
        grdLib.TextMatrix(grdLib.FixedRows, ENDDATEINDEX) = ""
        grdLib.TextMatrix(grdLib.FixedRows, CODEINDEX) = "0"
        igLibCallType = 0
        lgLibCallCode = 0
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
                grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
            End If
        Next llRow
        tmDHE.lCode = 0
        imLimboAllowed = True
    ElseIf (igLibCallType = 4) Then
        slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
        smNowDate = Format(gNow(), sgShowDateForm)
        If gDateValue(slDate) < gDateValue(smNowDate) Then
            grdLib.Col = STARTDATEINDEX
            grdLib.Row = grdLib.FixedRows
            grdLib.CellBackColor = LIGHTYELLOW
            imStartDateChangable = False
        End If
        grdLib.Row = grdLib.FixedRows
        For llCol = NAMEINDEX To STATEINDEX Step 1
            If (llCol <> ENDDATEINDEX) Then
                If ((llCol <> NAMEINDEX) And (llCol <> SUBLIBNAMEINDEX)) Or (Trim$(grdLib.TextMatrix(grdLib.FixedRows, llCol)) <> "") Then
                    If (llCol <> DESCRIPTIONINDEX) Or ((llCol = DESCRIPTIONINDEX) And (Trim$(grdLib.TextMatrix(grdLib.FixedRows, NAMEINDEX)) <> "")) Then
                        grdLib.Col = llCol
                        grdLib.CellBackColor = LIGHTYELLOW
                    End If
                End If
            End If
        Next llCol
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                For llCol = EVENTTYPEINDEX To imMaxCols Step 1
                    grdLibEvents.Col = llCol
                    grdLibEvents.Row = llRow
                    grdLibEvents.CellBackColor = LIGHTYELLOW
                Next llCol
            End If
        Next llRow
        If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
            imLimboAllowed = True
        End If
    End If

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

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If lgLibCallCode > 0 Then
        grdLibEvents.Redraw = False
        mMoveRecToCtrls
        grdLibEvents.Redraw = False
        grdLibEvents.Visible = False
        mMoveDEERecToCtrls
        mSortCol TIMEINDEX
        'If (igLibCallType = 1) Then
        '    slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
        '    smNowDate = Format(gNow(), sgShowDateForm)
        '    If gDateValue(slDate) < gDateValue(smNowDate) Then
        '        grdLib.Col = STARTDATEINDEX
        '        grdLib.Row = grdLib.FixedRows
        '        grdLib.CellBackColor = LIGHTYELLOW
        '        imStartDateChangable = False
        '    End If
        '    If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
        '        imLimboAllowed = True
        '    End If
        'ElseIf igLibCallType = 2 Then
        '    grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX) = ""
        '    grdLib.TextMatrix(grdLib.FixedRows, ENDDATEINDEX) = ""
        '    grdLib.TextMatrix(grdLib.FixedRows, CODEINDEX) = "0"
        '    igLibCallType = 0
        '    lgLibCallCode = 0
        '    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        '        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        '            grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
        '            grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
        '        End If
        '    Next llRow
        '    tmDHE.lCode = 0
        '    imLimboAllowed = True
        'ElseIf (igLibCallType = 4) Then
        '    slDate = grdLib.TextMatrix(grdLib.FixedRows, STARTDATEINDEX)
        '    smNowDate = Format(gNow(), sgShowDateForm)
        '    If gDateValue(slDate) < gDateValue(smNowDate) Then
        '        grdLib.Col = STARTDATEINDEX
        '        grdLib.Row = grdLib.FixedRows
        '        grdLib.CellBackColor = LIGHTYELLOW
        '        imStartDateChangable = False
        '    End If
        '    grdLib.Row = grdLib.FixedRows
        '    For llCol = NAMEINDEX To STATEINDEX Step 1
        '        If (llCol <> ENDDATEINDEX) Then
        '            If (llCol <> SUBLIBNAMEINDEX) Or (Trim$(grdLib.TextMatrix(grdLib.FixedRows, llCol)) <> "") Then
        '                grdLib.Col = llCol
        '                grdLib.CellBackColor = LIGHTYELLOW
        '            End If
        '        End If
        '    Next llCol
        '    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        '        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        '            For llCol = EVENTTYPEINDEX To imMaxCols Step 1
        '                grdLibEvents.Col = llCol
        '                grdLibEvents.Row = llRow
        '                grdLibEvents.CellBackColor = LIGHTYELLOW
        '            Next llCol
        '        End If
        '    Next llRow
        '    If grdLib.TextMatrix(grdLib.FixedRows, STATEINDEX) = "Limbo" Then
        '        imLimboAllowed = True
        '    End If
        'End If
        mSetBackgroundColor
        grdLibEvents.Redraw = True
        grdLibEvents.Visible = True
    Else
        bmIntegralSet = True
        gGrid_IntegralHeight grdLibEvents
        gGrid_FillWithRows grdLibEvents
        '8/26/11: Remove one row so row is not behind scroll bar
        grdLibEvents.Height = grdLibEvents.Height - grdLibEvents.RowHeight(0) '+ 30
        imLimboAllowed = True
    End If

End Sub
