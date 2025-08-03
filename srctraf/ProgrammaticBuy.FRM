VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ProgrammaticBuy 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7080
   ClientLeft      =   540
   ClientTop       =   2370
   ClientWidth     =   9765
   ClipControls    =   0   'False
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
   Icon            =   "ProgrammaticBuy.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9765
   WindowState     =   1  'Minimized
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   6120
   End
   Begin VB.CommandButton btn_Stop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6120
      TabIndex        =   41
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btn_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   40
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Start"
      Height          =   375
      Left            =   4560
      TabIndex        =   39
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lb_Messages 
      Height          =   645
      Left            =   480
      TabIndex        =   38
      Top             =   6720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox lacButtonHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4500
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3585
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Frame frcTab 
      Caption         =   "Research"
      Height          =   4245
      Left            =   375
      TabIndex        =   13
      Top             =   1290
      Width           =   9240
      Begin VB.ListBox lbcVefGp5 
         Height          =   840
         Left            =   7320
         TabIndex        =   43
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcPer 
         Height          =   360
         Left            =   5790
         TabIndex        =   18
         Top             =   2850
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin VB.TextBox edcData 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2820
         TabIndex        =   22
         Top             =   1950
         Visible         =   0   'False
         Width           =   1665
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcType 
         Height          =   360
         Left            =   3885
         TabIndex        =   15
         Top             =   2565
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcProduct 
         Height          =   345
         Left            =   2010
         TabIndex        =   24
         Top             =   1545
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   609
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcNetwork 
         Height          =   360
         Left            =   1980
         TabIndex        =   21
         Top             =   2925
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcLength 
         Height          =   360
         Left            =   4725
         TabIndex        =   17
         Top             =   2475
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcDemo 
         Height          =   360
         Left            =   3480
         TabIndex        =   16
         Top             =   2370
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcDaypart 
         Height          =   360
         Left            =   3570
         TabIndex        =   20
         Top             =   3210
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin V81ProgrammaticBuy.CSI_ComboBoxList cbcAdvt 
         Height          =   360
         Left            =   1290
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   0
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   9765
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   14
         Top             =   3210
         Width           =   105
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   9750
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   26
         Top             =   3180
         Width           =   75
      End
      Begin VB.TextBox edcComment 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3765
         Width           =   5295
      End
      Begin V81ProgrammaticBuy.CSI_Calendar dtcEndDate 
         Height          =   345
         Left            =   3375
         TabIndex        =   19
         Top             =   2055
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         Text            =   "4/14/2020"
         BackColor       =   -2147483643
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CSI_CurDayBackColor=   16777215
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81ProgrammaticBuy.CSI_Calendar dtcStartDate 
         Height          =   420
         Left            =   2505
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   741
         Text            =   "4/14/2020"
         BackColor       =   -2147483643
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
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CSI_CurDayBackColor=   16777215
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   -1  'True
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHeader 
         Height          =   675
         Index           =   0
         Left            =   15
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   1191
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLines 
         Height          =   3795
         Index           =   0
         Left            =   30
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1185
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   6694
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
      Begin MSWinsockLib.Winsock sckClient 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lacSpacer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         TabIndex        =   36
         Top             =   690
         Width           =   9060
      End
      Begin VB.Label lacComment 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   35
         Top             =   3780
         Width           =   1215
      End
      Begin VB.Label lacTotalSpots 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   6945
         TabIndex        =   34
         Top             =   3750
         Width           =   930
      End
      Begin VB.Label lacTotalPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   8475
         TabIndex        =   33
         Top             =   3765
         Width           =   825
      End
      Begin VB.Label lacTotalCPP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   6990
         TabIndex        =   32
         Top             =   4230
         Width           =   930
      End
      Begin VB.Label lacTotalGrpPct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   8040
         TabIndex        =   31
         Top             =   4185
         Width           =   930
      End
      Begin VB.Label lacTotalGRP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   6105
         TabIndex        =   30
         Top             =   4350
         Width           =   930
      End
   End
   Begin VB.CommandButton cmcHelp 
      Appearance      =   0  'Flat
      Caption         =   "&Help (Off)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   300
      Width           =   1230
   End
   Begin VB.CommandButton cmcSubmit 
      Appearance      =   0  'Flat
      Caption         =   "&Submit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   7
      Top             =   6165
      Width           =   1560
   End
   Begin VB.CommandButton cmcSignOff 
      Appearance      =   0  'Flat
      Caption         =   "Sign-&Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   6
      Top             =   6165
      Width           =   1560
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      TabIndex        =   5
      Top             =   6165
      Width           =   1560
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   255
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6030
      Width           =   75
   End
   Begin VB.PictureBox pbcFocusStop 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   -30
      ScaleHeight     =   210
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   6015
      Width           =   135
   End
   Begin VB.PictureBox pbcBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9765
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9765
      Begin VB.PictureBox pbcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   60
         ScaleHeight     =   675
         ScaleWidth      =   2985
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   2985
      End
      Begin VB.Label lacBanner 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   375
         TabIndex        =   9
         Top             =   780
         Visible         =   0   'False
         Width           =   5415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacBanner 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   2
         Left            =   3750
         TabIndex        =   12
         Top             =   60
         Width           =   1710
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2003-2011"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   7290
         TabIndex        =   2
         Top             =   405
         Width           =   630
      End
      Begin VB.Image imcLogo 
         Height          =   615
         Left            =   4950
         Picture         =   "ProgrammaticBuy.frx":1542
         Top             =   0
         Width           =   3180
      End
      Begin VB.Label lacBanner 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10020
      Top             =   6105
   End
   Begin ComctlLib.TabStrip tscProgrammaticBuy 
      Height          =   4080
      Left            =   195
      TabIndex        =   37
      Top             =   1575
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   7197
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CPP"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CPM"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Price"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Proof of Performance"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   0
      TabIndex        =   44
      Tag             =   "The number and extension of the buyer."
      Top             =   5760
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA Ext(AAAA)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Not connected."
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   6360
      Width           =   5415
   End
End
Attribute VB_Name = "ProgrammaticBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 2018 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ProgrammaticBuy.Frm

Option Explicit
Option Compare Text
Dim imSlfCode As Integer 'Salesperson defined as programmatic

Dim imMnfComp(0 To 1) As Integer
Dim imMnfExcl(0 To 1) As Integer
Dim smSalesBrochure As String
Dim bmProgrammaticVehExist As Boolean
Dim bmProgrammaticRifExist As Boolean
Dim smTabType As String
Dim imFixing As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imLbcArrowSetting As Integer
Dim lmNowDate As Long
Dim lmTimeWidth As Long
Dim lmLengthWidth As Long
Dim imChgMode As Integer
Dim imAllowUndo As Integer  'Undo only allowed if not complete
Dim bmHelpOff As Boolean
Dim bmIgnoreScroll As Boolean
Dim dmCurrentTime As Date

Dim imSignOn As Integer     '0=Have not sign on; 1=Sign-on

Dim lmHdEnableRow As Long
Dim lmHdEnableCol As Long
Dim lmLnEnableRow As Long
Dim lmLnEnableCol As Long
Dim imCtrlVisible As Integer
Dim imWhichGrid As Integer '0=Header; 1=Line
Dim smZone As String

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Dim fmFontSize As Single
Dim bmFontBold As Boolean
Dim fmTotalFontSize As Single
Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

'Required by gMakeSsf
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim imUrfCode As Integer
Dim smUrfPrgmmaticAlert As String

Dim adf_rst As ADODB.Recordset

Dim prf_rst As ADODB.Recordset

Dim rcf_rst As ADODB.Recordset
Dim tmRcf As RCF
Dim imRcfCode As Integer
Dim smRcfName As String

Dim rif_rst As ADODB.Recordset
Dim tmRif() As RIF

Dim rdf_rst As ADODB.Recordset
Dim tmRdf() As RDF
Dim vef_rst As ADODB.Recordset
Dim pvf_rst As ADODB.Recordset
Dim imVefCode() As Integer
Dim att_rst As ADODB.Recordset

Dim hmDrf As Integer
Dim hmMnf As Integer
Dim hmDpf As Integer
Dim hmDef As Integer
Dim hmRaf As Integer
Dim hmVLF As Integer
Dim hmAUF As Integer

Dim imTabIndex As Integer
Dim tmVehGp5Code() As SORTCODE
Dim smVehGp5CodeTag As String
Dim smFileName As String
Dim imTokenIdx As Integer

'Comm Server Vars
Private ilFirstResponse As Integer
Private smResponse As String
Private smCurrentToken As String
Private CommServerIP As String
Private CommServerPort As String
Private bmShutdownNeeded As Boolean
Private smNetworkName As String
Private smArgs() As String

'Header Request Arrived
Private smHeaderArray() As String
Private smLineArray() As String
Private smHeaderAdvt As String
Private imHeaderAdvtCode As Integer
Private smHeaderProd  As String
Private smHeaderDemo  As String
Private imHeaderDemo  As Integer
Private smHeaderStartDate  As String
Private smHeaderWeeks  As String
Private smHeaderEndDate  As String
Private smHeaderEstimate  As String
Private smHeaderTargetGRP  As String
Private smHeaderSpotLenghts  As String
Private smPriceRowPrice  As String

'Totals at bottom
Private smTtlCPP As String
Private smTtlPercentGRP As String
Private smTtlSpots As String
Private smTtlPrice As String
Private smTtlGRP As String

'Site Options
Private smPrgmmaticAllow As String
Private smShowAvailCount As String
Private smShowCPPTab As String
Private smShowCPMTab As String
Private smShowPriceTab As String

'Proposals
Dim hmPvf As Integer
Dim tmPvf As PVF
Dim tmPvfSrchKey As LONGKEY0
Dim imPvfRecLen As Integer

Dim hmCHF As Integer
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1
Dim imCHFRecLen As Integer        'RVF record length

Dim hmClf As Integer
Dim tmClf() As CLF            'CLF record image
Dim imGrdClfRow() As Integer
Dim imClfRecLen As Integer        'RVF record length

Dim hmTmf As Integer
Dim tmTmf() As CLF            'TMF record image
Dim imTmfRecLen As Integer    'TMF record length

Dim hmCff As Integer
Dim tmCff() As CFF            'CFF record image
Dim imCffRecLen As Integer        'RVF record length
Dim bmAllowDays(0 To 6) As Boolean

Dim hmCxf As Integer
Dim tmCxf As CXF            'CXF record image
Dim imCxfRecLen As Integer        'RVF record length

'MultiName
Dim tmMnf As MNF            'MNF record image
Dim tmMnfSrchKey As INTKEY0 'MNF key record image
Dim imMnfRecLen As Integer     'MNF record length

Private Type RATEDGP
    iMnfVehGp5Rsch As Integer
    sRated As String * 1
End Type
Dim tmRatedGp() As RATEDGP

Const ADVERTISERINDEX = 0
Const PRODUCTINDEX = 1
Const STARTDATEINDEX = 2
Const NUMBERWEEKSINDEX = 3
Const ENDDATEINDEX = 4
Const ESTIMATENUMBERINDEX = 5
Const DEMOINDEX = 6
Const TARGETINDEX = 7

Const LENGTHINDEX = 0
Const NETWORKINDEX = 1
Const DAYPARTINDEX = 2
Const STATIONINDEX = 3
Const SPOTPRICEINDEX = 4
Const AVAILINDEX = 5
Const CPPCPMINDEX = 6
Const GRPGRIMPPCTINDEX = 7
Const NUMBERSPOTSINDEX = 8
Const SPOTSPERINDEX = 9
Const TOTALPRICEINDEX = 10
Const LNGRPINDEX = 11
Const AUDIENCEINDEX = 12
Const POPULATIONINDEX = 13
Const RATINGINDEX = 14
Const PROGRAMMATICALLOWED = &H1  'Programmatic Allowed
Const SHOWAVAILCOUNT = &H2       'Programmatic Show Avail Count
Const SHOWCPPTAB = &H4           'Programmatic Show CPP Tab
Const SHOWCPMTAB = &H8           'Programmatic Show CPM Tab
Const SHOWPRICETAB = &H10        'Programmatic Show Price Tab

'*******************************************************************************
' This code is called everytime we receive data from the server.
'*******************************************************************************
    Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strResponse As String
    Dim slScreenRequested As String
    Dim slArray() As String
    Dim slTempArray() As String
    Dim slTemp As String
    Dim blSend As Boolean
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilLineCount As Integer
    Dim llCurrRowNum As Long
    Dim ilIdx As Integer
    Dim ilVefCode As Integer
    Dim ilVff As Integer
    
        
    strData = ""
    sckClient.GetData strData, vbString, 2048 ' Receive the data
    If strData = "NOTHING" Then
        If bmShutdownNeeded Then
            cmcSignOff_Click
        Else
            'Idle time timer in seconds; value set in Traffic.ini file
            If DateDiff("s", dmCurrentTime, Now) > lgIdleTime Then
                gLogMsg "<ProgMsg> Idle time exceeded. Shutting down Programmatic Buy", "ProgrammaticBuy.Txt", False
                bmShutdownNeeded = True
                cmcSignOff_Click
            End If
            Exit Sub
        End If
    Else
        If strData <> "Welcome to the CSI Communications Center." Then
            dmCurrentTime = Now
        End If
    End If
        
    If ilFirstResponse = 0 Then ' If this is the first response, validate it.
        lb_Messages.Clear
        ilFirstResponse = ilFirstResponse + 1
        If strData <> "Welcome to the CSI Communications Center." Then
            sckClient.Close
            gLogMsg "<Comm>    " & strData, "ProgrammaticBuy.Txt", False
        End If
        Exit Sub
    End If
    
    ' Requests are formated like this.  REQUEST|1|sTokenID
    If InStr(1, strData, "REQUEST|") > 0 Then
        gLogMsg "<Comm>    " & strData, "ProgrammaticBuy.Txt", False
        lb_Messages.AddItem (strData)
        DoEvents
        slArray = Split(strData, "|")
        slScreenRequested = slArray(1)
        strResponse = "RESPONSE|"
        
        'set the correct index for array of grids
        If UBound(tgTokenInfo) > 0 And slScreenRequested <> "LOGIN" Then
            For ilIdx = LBound(tgTokenInfo) To UBound(tgTokenInfo) Step 1
                If Trim(tgTokenInfo(ilIdx).sCommToken) = Trim(slArray(2)) Then
                    imTokenIdx = ilIdx
                    Exit For
                End If
            Next ilIdx
            mSetCurrGridVisable
        End If
        
        Select Case slScreenRequested
        '******************************************************** Misc. Support Functions ***************************************************************
            Case "LOGIN"
                imTokenIdx = -1
                If UBound(tgTokenInfo) >= 0 Then
                    For ilIdx = LBound(tgTokenInfo) To UBound(tgTokenInfo) Step 1
                        If Trim(tgTokenInfo(ilIdx).sCommToken) = Trim(slArray(2)) Then
                            imTokenIdx = ilIdx
                            Exit For
                        End If
                    Next ilIdx
                End If
                
                If imTokenIdx = -1 Then
                    'We didn't find the token in use so look for an existing element that is not in use
                    For ilIdx = 0 To UBound(tgTokenInfo)
                        If UCase(tgTokenInfo(ilIdx).sInUseStatus) = "N" Then
                            imTokenIdx = ilIdx
                            tgTokenInfo(ilIdx).sCommToken = Trim(slArray(2))
                            tgTokenInfo(ilIdx).sInUseStatus = "Y"
                            ReDim Preserve tgTokenInfo(0 To ilIdx + 1) As TOKENINFO
                            tgTokenInfo(ilIdx + 1).sInUseStatus = "N"
                            Exit For
                        End If
                    Next ilIdx
                End If
                
                If imTokenIdx >= 0 Then
                    'We found an existing element not in use
                    If imTokenIdx > 0 Then
                        Load grdHeader(imTokenIdx)
                        Load grdLines(imTokenIdx)
                    End If
                    slTempArray = Split(slArray(3), "~")
                    tgTokenInfo(imTokenIdx).sCommToken = Trim(slArray(2))
                    smCurrentToken = Trim(tgTokenInfo(imTokenIdx).sCommToken)
                    tgTokenInfo(imTokenIdx).iAgfCode = Trim(slTempArray(0))
                    tgTokenInfo(imTokenIdx).sAgency = Trim(slTempArray(1))
                    tgTokenInfo(imTokenIdx).sEMail = Trim(slTempArray(2))
                    tgTokenInfo(imTokenIdx).sBuyer = Trim(slTempArray(3)) & " " & Trim(slTempArray(4))
                    tgTokenInfo(imTokenIdx).sPhone = Trim(slTempArray(5))
                    tgTokenInfo(imTokenIdx).sInUseStatus = "Y"
                    tgTokenInfo(imTokenIdx).sDate = Format(gNow, "mm/dd/yyyy")
                    tgTokenInfo(imTokenIdx).sTime = Format(gNow, "hh:mm:ssam/pm")
                    tgTokenInfo(imTokenIdx).iRcfCode = mDetermineRateCard
                Else
                    'We did not find any existing element that was not in use
                    ReDim tgTokenInfo(0 To UBound(tgTokenInfo) + 1) As TOKENINFO
                    slTempArray = Split(slArray(3), "~")
                    tgTokenInfo(UBound(tgTokenInfo)).sCommToken = Trim(slArray(2))
                    smCurrentToken = Trim(tgTokenInfo(UBound(tgTokenInfo)).sCommToken)
                    tgTokenInfo(UBound(tgTokenInfo)).iAgfCode = Trim(slTempArray(0))
                    tgTokenInfo(UBound(tgTokenInfo)).sAgency = Trim(slTempArray(1))
                    tgTokenInfo(UBound(tgTokenInfo)).sEMail = Trim(slTempArray(2))
                    tgTokenInfo(UBound(tgTokenInfo)).sBuyer = Trim(slTempArray(3)) & " " & Trim(slTempArray(4))
                    tgTokenInfo(UBound(tgTokenInfo)).sPhone = Trim(slTempArray(5))
                    tgTokenInfo(UBound(tgTokenInfo)).sInUseStatus = "Y"
                    tgTokenInfo(UBound(tgTokenInfo)).sDate = Format(gNow, "mm/dd/yyyy")
                    tgTokenInfo(UBound(tgTokenInfo)).sTime = Format(gNow, "hh:mm:ssam/pm")
                    tgTokenInfo(UBound(tgTokenInfo)).iRcfCode = mDetermineRateCard
                End If
                mSetCurrGridVisable
                strResponse = strResponse & "READY"
                blSend = True
            Case "SEND_SITE_OPTIONS"
                mGetSiteOptions
                strResponse = strResponse & "[""safShowAvailCount"", ""safShowCPPTab"", ""safShowCPMTab"", ""safShowPriceTab"", ""spfsAllowDailyBuys""],"
                strResponse = strResponse & "["
                strResponse = strResponse & """" & Trim$(smShowAvailCount) & """"
                strResponse = strResponse & "," & """" & Trim$(smShowCPPTab) & """"
                strResponse = strResponse & "," & """" & Trim$(smShowCPMTab) & """"
                strResponse = strResponse & "," & """" & Trim$(smShowPriceTab) & """"
                strResponse = strResponse & "," & """" & Trim$(tgSpf.sAllowDailyBuys) & """"
                strResponse = strResponse & "]"
                blSend = True
                If smShowCPPTab = "Y" Then
                    Set tscProgrammaticBuy.SelectedItem = tscProgrammaticBuy.Tabs.Item(1)
                ElseIf smShowCPMTab = "Y" Then
                    Set tscProgrammaticBuy.SelectedItem = tscProgrammaticBuy.Tabs.Item(2)
                Else
                    Set tscProgrammaticBuy.SelectedItem = tscProgrammaticBuy.Tabs.Item(3)
                End If
            Case "SEND_ADVERTISERS"
                mAdvtPop
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_PRODUCTS"
                mGetAdfCode (slArray(3))
                mProdPop CInt(smResponse)
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_AGENCY_LIST"
                mGatherAgencyInfo
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_LENGTHS"
                cmcUndo_Click
                mLengthPop
                strResponse = strResponse & smResponse
                blSend = True
            Case "SIGNOFF_PAGE"
                mSignOff
                strResponse = strResponse & "SUCCESS"
                blSend = True
                bmShutdownNeeded = True
            Case "ADMIM_PAGE"
                strResponse = strResponse & "SUCCESS"
                blSend = True
            '**************************************** Submit, Reset and Tab/Page Change Options ****************************************
            Case "SUBMIT_PAGE"
                edcComment.Text = Trim(slArray(3))
                mSubmit
                strResponse = strResponse & "SUCCESS: " & smResponse
                'MsgBox "Submitting Proposal/Order", vbOKCancel
                blSend = True
            Case "SEND_BROCHURE"
                strResponse = strResponse & SendBrochure(Trim(slArray(3)))
                sckClient.SendData strResponse
                blSend = True
            Case "RESET_PAGE"
                cmcUndo_Click
                strResponse = strResponse & "SUCCESS"
                blSend = True
            Case "PAGE_CHANGE_CPP"
                cmcUndo_Click
                strResponse = strResponse & "SUCCESS"
                tscProgrammaticBuy.Tabs(1).Selected = True
                tscProgrammaticBuy_Click
                strResponse = strResponse & "SUCCESS"
                blSend = True
            Case "PAGE_CHANGE_CPM"
                cmcUndo_Click
                strResponse = strResponse & "SUCCESS"
                tscProgrammaticBuy.Tabs(2).Selected = True
                tscProgrammaticBuy_Click
                strResponse = strResponse & "SUCCESS"
                blSend = True
            Case "PAGE_CHANGE_Price"
                cmcUndo_Click
                strResponse = strResponse & "SUCCESS"
                tscProgrammaticBuy.Tabs(3).Selected = True
                tscProgrammaticBuy_Click
                strResponse = strResponse & "SUCCESS"
                blSend = True
            '*************************************************** CPP Tab ***************************************************************
            Case "SEND_TAB_CPP_BUYLIST"
                smTabType = "CPP"
                smHeaderArray = Split(slArray(3), "~")
                imWhichGrid = 0
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ADVERTISERINDEX
                cbcAdvt.Text = smHeaderArray(0)
                smHeaderAdvt = smHeaderArray(0)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = PRODUCTINDEX
                mGetAdfCode smHeaderArray(0)
                mProdPop CInt(smResponse)
                cbcProduct.Text = smHeaderArray(1)
                smHeaderProd = smHeaderArray(1)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = STARTDATEINDEX
                dtcStartDate.Text = smHeaderArray(2)
                smHeaderStartDate = dtcStartDate.Text
                mSetShow
                
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = NUMBERWEEKSINDEX
                edcData.Text = smHeaderArray(3)
                smHeaderWeeks = edcData.Text
                mSetShow
                smHeaderEndDate = smHeaderArray(4)
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ESTIMATENUMBERINDEX
                edcData.Text = smHeaderArray(5)
                smHeaderEstimate = edcData.Text
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = DEMOINDEX
                cbcDemo.Text = smHeaderArray(6)
                smHeaderDemo = cbcDemo.Text
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = TARGETINDEX
                edcData.Text = smHeaderArray(7)
                smHeaderTargetGRP = edcData.Text
                mSetShow
                imWhichGrid = 1
                lmLnEnableRow = grdLines(imTokenIdx).FixedRows
                For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        lmLnEnableRow = lmLnEnableRow + 1
                    Else
                       ' Exit For
                    End If
                Next ilRow
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = LENGTHINDEX
                cbcLength.Text = smHeaderArray(8)
                smHeaderSpotLenghts = cbcLength.Text
                mSetShow
                smResponse = ""
                smResponse = smResponse & "[""Length"", ""Network"", ""B"", ""Daypart"", ""Stations"", ""Avails"", ""CPP"", ""%GRP"", ""# Spots"", ""Total Price"", ""GRP""]"
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                If ilLineCount - llCurrRowNum >= 0 Then
                    smResponse = smResponse & ","
                End If
                For ilRow = llCurrRowNum To ilLineCount Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        smResponse = smResponse & "[" & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)) & """"
                        ilVefCode = gGetVehCodeByVefName(Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)), False)
                        ilVff = gBinarySearchVff(ilVefCode)
                        smSalesBrochure = ""
                        If ilVff <> -1 Then
                            smSalesBrochure = Trim(tgVff(ilVff).sSalesBrochure)
                        End If
                        smResponse = smResponse & "," & """" & smSalesBrochure & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, DAYPARTINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, STATIONINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, AVAILINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, CPPCPMINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, GRPGRIMPPCTINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, SPOTSPERINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, GRPGRIMPPCTINDEX)) & """"
                        
                        If ilRow < ilLineCount Then
                            smResponse = smResponse & "],"
                        Else
                            smResponse = smResponse & "]"
                        End If
                    End If
                Next ilRow
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_CPP_ROWTOTAL"
                smLineArray = Split(slArray(3), "~")
                imWhichGrid = 1
                lmLnEnableRow = smLineArray(0) + 1
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = GRPGRIMPPCTINDEX
                edcData.Text = smLineArray(1)
                mSetShow
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""RowNum"", ""NumSpots"", ""TTLPrice"", ""GRP""],"
                smResponse = smResponse & "[" & """" & smLineArray(0) & """"
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, NUMBERSPOTSINDEX)) & """"
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, TOTALPRICEINDEX)) & """"
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, LNGRPINDEX)) & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_CPP_TOTALS"
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""Total CPM"", ""Total %GRP"", ""Total Spots"", ""Total Price"", ""Total GRP""],"
                smResponse = smResponse & "[" & """" & smTtlCPP & """"
                smResponse = smResponse & "," & """" & smTtlPercentGRP & """"
                smResponse = smResponse & "," & """" & smTtlSpots & """"
                smResponse = smResponse & "," & """" & smTtlPrice & """"
                smResponse = smResponse & "," & """" & smTtlGRP & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            '**************************************** CPM Tab ****************************************
            Case "SEND_TAB_CPM_BUYLIST"
                smTabType = "CPM"
                smHeaderArray = Split(slArray(3), "~")
                imWhichGrid = 0
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ADVERTISERINDEX
                cbcAdvt.Text = smHeaderArray(0)
                smHeaderAdvt = smHeaderArray(0)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = PRODUCTINDEX
                mGetAdfCode smHeaderArray(0)
                mProdPop CInt(smResponse)
                cbcProduct.Text = smHeaderArray(1)
                smHeaderProd = smHeaderArray(1)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = STARTDATEINDEX
                dtcStartDate.Text = smHeaderArray(2)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = NUMBERWEEKSINDEX
                edcData.Text = smHeaderArray(3)
                smHeaderWeeks = edcData.Text
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ESTIMATENUMBERINDEX
                edcData.Text = smHeaderArray(5)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = DEMOINDEX
                cbcDemo.Text = smHeaderArray(6)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = TARGETINDEX
                edcData.Text = smHeaderArray(7)
                '   smHeaderSpotLenghts = cbcLength.Text
                mSetShow
                imWhichGrid = 1
                lmLnEnableRow = grdLines(imTokenIdx).FixedRows
                For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        lmLnEnableRow = lmLnEnableRow + 1
                    Else
                        Exit For
                    End If
                Next ilRow
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = LENGTHINDEX
                cbcLength.Text = smHeaderArray(8)
                mSetShow
                smResponse = ""
                smResponse = smResponse & "[""Length"", ""Network"", ""B"", ""Daypart"", ""Stations"", ""Avails"", ""CPM"", ""%Grimp"", ""# Spots"", ""Total Price"", ""Grimp""],"
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                
                For ilRow = llCurrRowNum To ilLineCount Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        smResponse = smResponse & "[" & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)) & """"
                        'smResponse = smResponse & "," & """" & "" & """"
                        
                        
                        ilVefCode = gGetVehCodeByVefName(Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)), False)
                        ilVff = gBinarySearchVff(ilVefCode)
                        smSalesBrochure = ""
                        If ilVff <> -1 Then
                            smSalesBrochure = Trim(tgVff(ilVff).sSalesBrochure)
                            smResponse = smResponse & "," & """" & smSalesBrochure & """"
                        Else
                            smResponse = smResponse & "," & """" & "" & """"
                        End If
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, DAYPARTINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, STATIONINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, AVAILINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, CPPCPMINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, GRPGRIMPPCTINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, SPOTSPERINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, GRPGRIMPPCTINDEX)) & """"
                        If ilRow < ilLineCount Then
                            smResponse = smResponse & "],"
                        Else
                            smResponse = smResponse & "]"
                        End If
                    End If
                Next ilRow
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_CPM_ROWTOTAL"
                smLineArray = Split(slArray(3), "~")
                imWhichGrid = 1
                lmLnEnableRow = smLineArray(0) + 1
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = GRPGRIMPPCTINDEX
                edcData.Text = smLineArray(1)
                ilIdx = llCurrRowNum - 1
                mSetShow
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""RowNum"", ""NumSpots"", ""TTLPrice"", ""Grimp""],"
                smResponse = smResponse & "[" & """" & ilIdx & """"
                'llCurrRowNum = llCurrRowNum + 1
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, NUMBERSPOTSINDEX)) & """"
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, TOTALPRICEINDEX)) & """"
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, LNGRPINDEX)) & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_CPM_TOTALS"
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""Total CPM"", ""Total %GRIMP"", ""Total Spots"", ""Total Price"", ""Total Grimp""],"
                smResponse = smResponse & "[" & """" & smTtlCPP & """"
                smResponse = smResponse & "," & """" & smTtlPercentGRP & """"
                smResponse = smResponse & "," & """" & smTtlSpots & """"
                smResponse = smResponse & "," & """" & smTtlPrice & """"
                smResponse = smResponse & "," & """" & smTtlGRP & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            '**************************************************** Price Tab *****************************************************
            Case "SEND_TAB_PRICE_BUYLIST"
                smTabType = "Price"
                smHeaderArray = Split(slArray(3), "~")
                imWhichGrid = 0
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ADVERTISERINDEX
                cbcAdvt.Text = smHeaderArray(0)
                smHeaderAdvt = smHeaderArray(0)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = PRODUCTINDEX
                mGetAdfCode smHeaderArray(0)
                mProdPop CInt(smResponse)
                cbcProduct.Text = smHeaderArray(1)
                smHeaderProd = smHeaderArray(1)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = STARTDATEINDEX
                dtcStartDate.Text = smHeaderArray(2)
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = NUMBERWEEKSINDEX
                edcData.Text = smHeaderArray(3)
                smHeaderWeeks = edcData.Text
                mSetShow
                lmHdEnableRow = grdHeader(imTokenIdx).FixedRows
                lmHdEnableCol = ESTIMATENUMBERINDEX
                edcData.Text = smHeaderArray(5)
                smHeaderSpotLenghts = cbcLength.Text
                mSetShow
                imWhichGrid = 1
                lmLnEnableRow = grdLines(imTokenIdx).FixedRows
                For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        lmLnEnableRow = lmLnEnableRow + 1
                    Else
                        Exit For
                    End If
                Next ilRow
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = LENGTHINDEX
                cbcLength.Text = smHeaderArray(8)
                mSetShow
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""Length"", ""Network"", ""B"", ""Daypart"", ""Stations"", ""Spot Price"", ""Avails"", ""# Spots"", ""Per"", ""Total Price""],"
                For ilRow = llCurrRowNum To ilLineCount Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        smResponse = smResponse & "[" & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)) & """"
                        ilVefCode = gGetVehCodeByVefName(Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)), False)
                        ilVff = gBinarySearchVff(ilVefCode)
                        smSalesBrochure = ""
                        If ilVff <> -1 Then
                            smSalesBrochure = Trim(tgVff(ilVff).sSalesBrochure)
                            smResponse = smResponse & "," & """" & smSalesBrochure & """"
                        Else
                            smResponse = smResponse & "," & """" & "" & """"
                        End If
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, DAYPARTINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, STATIONINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, SPOTPRICEINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, AVAILINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, SPOTSPERINDEX)) & """"
                        smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX)) & """"
                        If ilRow < ilLineCount Then
                            smResponse = smResponse & "],"
                        Else
                            smResponse = smResponse & "]"
                        End If
                    End If
                Next ilRow
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_PRICE_ROWTOTAL"
                smLineArray = Split(slArray(3), "~")
                imWhichGrid = 2
                lmLnEnableRow = smLineArray(0)
                llCurrRowNum = lmLnEnableRow
                lmLnEnableCol = NUMBERSPOTSINDEX
                grdLines(imTokenIdx).TextMatrix(llCurrRowNum + 1, NUMBERSPOTSINDEX) = smLineArray(1)
                mSetShow
                lmLnEnableCol = SPOTSPERINDEX
                grdLines(imTokenIdx).TextMatrix(llCurrRowNum + 1, SPOTSPERINDEX) = smLineArray(2)
                mSetShow
                lmLnEnableCol = TOTALPRICEINDEX
                mComputeTotals
                'D.S. 12-12-18  Commented out line below.  It screws up the row totals and grand totals
                'grdLines(imTokenIdx).TextMatrix(llCurrRowNum + 1, TOTALPRICEINDEX) = smPriceRowPrice
                mSetShow
                ilLineCount = 0
                For ilRow = llCurrRowNum To grdLines(imTokenIdx).Rows - 1 Step 1
                    If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)) <> "" Then
                        ilLineCount = ilLineCount + 1
                    End If
                Next ilRow
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""RowNum"", ""Total Price""],"
                smResponse = smResponse & "[" & """" & llCurrRowNum & """"
                llCurrRowNum = llCurrRowNum + 1
                smResponse = smResponse & "," & """" & Trim$(grdLines(imTokenIdx).TextMatrix(llCurrRowNum, TOTALPRICEINDEX)) & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            Case "SEND_TAB_PRICE_TOTALS"
                strResponse = "RESPONSE|"
                smResponse = ""
                smResponse = smResponse & "[""Spot Total"", ""Total Price""],"
                smResponse = smResponse & "[" & """" & smTtlSpots & """"
                smResponse = smResponse & "," & """" & Trim$(smTtlPrice) & """"
                smResponse = smResponse & "]"
                strResponse = strResponse & smResponse
                blSend = True
            End Select
        
        If blSend Then
            strResponse = strResponse & "<EOF>"
            sckClient.SendData strResponse
            DoEvents
            gLogMsg "<To: Comm> " & Trim(tgTokenInfo(imTokenIdx).sEMail) & " " & strResponse, "ProgrammaticBuy.Txt", False
            lb_Messages.AddItem strResponse
            blSend = False
        End If
    End If
End Sub

Private Sub cmcUndo_Click()
    Dim slCopy As String
    Dim ilStep As Integer
    
    'Exit Sub
    'If grdLines(imTokenIdx).TextMatrix(grdLines(imTokenIdx).FixedRows, NETWORKINDEX) <> "" Then
    
    ilStep = 0
        grdLines(imTokenIdx).Clear
        lacTotalSpots.Caption = ""
        lacTotalPrice.Caption = ""
        lacTotalGrpPct.Caption = ""
        lacTotalGRP.Caption = ""
        lacTotalCPP.Caption = ""
        smTtlCPP = ""
        ilStep = ilStep + 1
        mLineColumnNames
        mSetTitles
        imWhichGrid = 1
        grdLines(imTokenIdx).Row = grdLines(imTokenIdx).FixedRows
        grdLines(imTokenIdx).Col = LENGTHINDEX
        'mEnableBox
    'ElseIf grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, ADVERTISERINDEX) <> "" Then
        grdHeader(imTokenIdx).Clear
        ilStep = ilStep + 1
        mHeaderColumnNames
        grdLines(imTokenIdx).Clear
        lacTotalSpots.Caption = ""
        lacTotalPrice.Caption = ""
        lacTotalGrpPct.Caption = ""
        lacTotalGRP.Caption = ""
        lacTotalCPP.Caption = ""
        smTtlCPP = ""
        mLineColumnNames
        mSetTitles
        imWhichGrid = 0
        grdHeader(imTokenIdx).Row = grdHeader(imTokenIdx).FixedRows
        grdHeader(imTokenIdx).Col = ADVERTISERINDEX
        ''''imTerminate = True
        mEnableBox
    'End If
    mSetCommands
End Sub


Private Sub cmcUndo_GotFocus()
    mSetShow
End Sub


Private Sub edcComment_GotFocus()
    mSetShow
End Sub

Private Sub edcData_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcData_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcData_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcData.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If imWhichGrid = 0 Then
        Select Case lmHdEnableCol
            Case ADVERTISERINDEX
            Case PRODUCTINDEX
            Case STARTDATEINDEX
            Case NUMBERWEEKSINDEX
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
                slStr = edcData.Text
                slStr = Left$(slStr, edcData.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcData.SelStart - edcData.SelLength)
                If gCompNumberStr(slStr, "100") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case ENDDATEINDEX
            Case ESTIMATENUMBERINDEX
            Case DEMOINDEX
            Case TARGETINDEX
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    Else
        Select Case lmLnEnableCol
            Case LENGTHINDEX
            Case NUMBERSPOTSINDEX
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
                slStr = edcData.Text
                slStr = Left$(slStr, edcData.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcData.SelStart - edcData.SelLength)
                If gCompNumberStr(slStr, "9999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case GRPGRIMPPCTINDEX
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
                slStr = edcData.Text
                slStr = Left$(slStr, edcData.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcData.SelStart - edcData.SelLength)
                If gCompNumberStr(slStr, "100") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If

End Sub

Private Sub edcData_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If imWhichGrid = 0 Then
            Select Case lmHdEnableCol
                Case ADVERTISERINDEX
                Case PRODUCTINDEX
                Case STARTDATEINDEX
                Case NUMBERWEEKSINDEX
                Case ENDDATEINDEX
                Case ESTIMATENUMBERINDEX
                Case DEMOINDEX
                Case TARGETINDEX
            End Select
        Else
            Select Case lmLnEnableCol
                Case LENGTHINDEX
                Case NUMBERSPOTSINDEX
                Case GRPGRIMPPCTINDEX
            End Select
        End If
        'edcDropdown.SelStart = 0
        'edcDropdown.SelLength = Len(edcDropdown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If imWhichGrid = 0 Then
            Select Case lmHdEnableCol
                Case ADVERTISERINDEX
                Case PRODUCTINDEX
                Case STARTDATEINDEX
                Case NUMBERWEEKSINDEX
                Case ENDDATEINDEX
                Case ESTIMATENUMBERINDEX
                Case DEMOINDEX
                Case TARGETINDEX
            End Select
        Else
            Select Case lmLnEnableCol
                Case LENGTHINDEX
                Case NUMBERSPOTSINDEX
                Case GRPGRIMPPCTINDEX
            End Select
        End If
    End If

End Sub

Private Sub Form_Activate()
    sgDoneMsg = ""
End Sub

Private Sub Form_Load()

'    If App.PrevInstance Then
'        MsgBox "Only one copy of Programmatic Buy can run at a time, sorry", vbOKOnly + vbInformation, "Counterpoint"
'        End
'    End If
    bmShutdownNeeded = False
    mInit
    If Not imTerminate Then
        mCheckGG
        If igGGFlag = 0 Then
            tmcTerminate.Enabled = True
        End If
        gOpenTmf
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
        ReDim tgTokenInfo(0 To 0) As TOKENINFO
        'Initialize the first element to not in use
        tgTokenInfo(0).sInUseStatus = "N"
        dmCurrentTime = Now
        cmdGo_Click
    End If
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSetShow
End Sub


Private Sub Form_Resize()
    Dim llWidth As Long
    Dim ilCol As Integer
    Dim llColPos As Long
    
    If ProgrammaticBuy.Height - 2 * pbcBanner.Height - 120 < 0 Then
        Exit Sub
    End If
    If imTerminate Then
        Exit Sub
    End If
    grdHeader(imTokenIdx).Font.Bold = bmFontBold
    grdHeader(imTokenIdx).Font.Size = fmFontSize
    grdHeader(imTokenIdx).Font.Name = "Arial"
    grdHeader(imTokenIdx).FontFixed.Bold = bmFontBold
    grdHeader(imTokenIdx).FontFixed.Size = fmFontSize
    grdHeader(imTokenIdx).FontFixed.Name = "Arial"
    grdHeader(imTokenIdx).Font.Bold = bmFontBold
    grdHeader(imTokenIdx).FontFixed.Bold = bmFontBold
    grdHeader(imTokenIdx).RowHeight(0) = grdHeader(imTokenIdx).RowHeight(1)
    
    grdLines(imTokenIdx).Font.Bold = bmFontBold
    grdLines(imTokenIdx).Font.Size = fmFontSize
    grdLines(imTokenIdx).Font.Name = "Arial"
    grdLines(imTokenIdx).Font.Bold = bmFontBold
    grdLines(imTokenIdx).FontFixed.Bold = bmFontBold
    grdLines(imTokenIdx).FontFixed.Size = fmFontSize
    grdLines(imTokenIdx).FontFixed.Name = "Arial"
    grdLines(imTokenIdx).FontFixed.Bold = bmFontBold
    grdLines(imTokenIdx).RowHeight(0) = grdLines(imTokenIdx).RowHeight(1)
    
    tscProgrammaticBuy.Width = 0.95 * ProgrammaticBuy.Width
    tscProgrammaticBuy.Top = pbcBanner.Top + pbcBanner.Height
    tscProgrammaticBuy.Left = (ProgrammaticBuy.Width - tscProgrammaticBuy.Width) / 2 - 60
    tscProgrammaticBuy.Height = ProgrammaticBuy.Height - pbcBanner.Height - 2 * cmcSubmit.Height - lacTotalSpots.Height - 300
    frcTab.Top = tscProgrammaticBuy.Top + 355
    frcTab.Left = tscProgrammaticBuy.Left + 60
    frcTab.Width = tscProgrammaticBuy.Width - 120
    frcTab.Height = tscProgrammaticBuy.Height - 415
    frcTab.BorderStyle = 0
    frcTab.Caption = ""
    
    grdHeader(imTokenIdx).Height = 2 * grdHeader(imTokenIdx).RowHeight(1) + 60
    grdHeader(imTokenIdx).Width = frcTab.Width  '0.95 * ProgrammaticBuy.Width
    grdLines(imTokenIdx).Height = frcTab.Height - grdHeader(imTokenIdx).Height - 4 * lacSpacer.Height - 5 * grdLines(imTokenIdx).RowHeight(0)
    grdLines(imTokenIdx).Width = frcTab.Width   'grdHeader(imTokenIdx).Width     '0.9 * pbcProposal.Width
    grdHeader(imTokenIdx).Top = 120
    grdHeader(imTokenIdx).Left = 0
    grdLines(imTokenIdx).Top = grdHeader(imTokenIdx).Top + grdHeader(imTokenIdx).Height + 4 * lacSpacer.Height '60
    grdLines(imTokenIdx).Left = grdHeader(imTokenIdx).Left
    
    llWidth = cmcSubmit.Width
    cmcSubmit.Move ProgrammaticBuy.Width / 2 - (3 * cmcSubmit.Width) / 2 - llWidth, ProgrammaticBuy.Height - (6 * cmcSubmit.Height) / 2 'grdHeader(imTokenIdx).Top + grdHeader(imTokenIdx).Height + cmcSignOn.Height / 2
    cmcSignOff.Move cmcSubmit.Left + cmcSubmit.Width + llWidth, cmcSubmit.Top
    cmcUndo.Move cmcSignOff.Left + cmcSignOff.Width + llWidth, cmcSubmit.Top
    
    'lblStatus.Top = cmcSignOff.Top - 360
    lblStatus.Top = cmcSignOff.Top - 160
    lblStatus.Left = tscProgrammaticBuy.Left
    lblStatus.Width = lb_Messages.Width
    
    lb_Messages.Top = lblStatus.Top + 180
    lb_Messages.Left = lblStatus.Left
    
    cmdGo.Top = lblStatus.Top + 60
    cmdGo.Left = lb_Messages.Left + lb_Messages.Width + 60
    
    btn_Stop.Top = cmdGo.Top + cmdGo.Height + 30
    btn_Stop.Left = cmdGo.Left
    
    btn_Close.Top = cmdGo.Top + cmdGo.Height / 2
    btn_Close.Left = btn_Stop.Left + btn_Stop.Width + 60
    btn_Close.Visible = False
    
    mHeaderColumnWidths
    mLineColumnWidths
    lacSpacer.Move grdLines(imTokenIdx).Left, grdLines(imTokenIdx).Top - lacSpacer.Height, grdLines(imTokenIdx).Width
    cmcHelp.Left = grdHeader(imTokenIdx).Left + grdHeader(imTokenIdx).Width - cmcHelp.Width - 60
    lacBanner(0).Left = grdHeader(imTokenIdx).Left
    lacBanner(0).Height = 2 * grdLines(imTokenIdx).RowHeight(0)
    lacBanner(2).Top = lacBanner(0).Top
    lacBanner(2).Left = Me.Width \ 2 - lacBanner(2).Width / 2
    lacBanner(1).Width = cmcHelp.Left - lacBanner(2).Left
    lacBanner(1).Left = lacBanner(2).Left - 120
    lacBanner(1).Top = lacBanner(0).Top
    lacBanner(1).Height = pbcBanner.Height - 60
    pbcBanner.Visible = False


    imcLogo.Width = 2460
    imcLogo.Left = cmcHelp.Left - imcLogo.Width - lacCopyright.Width - cmcHelp.Width / 2
    lacCopyright.Left = imcLogo.Left + imcLogo.Width + 15
    lacCopyright.Caption = "2003-" & Year(Now)
    If imcLogo.Width + cmcHelp.Width + lacCopyright.Width + cmcHelp.Width / 2 < pbcBanner.Width Then
        lacBanner(0).Width = pbcBanner.Width - imcLogo.Width - cmcHelp.Width - lacCopyright.Width - cmcHelp.Width / 2
    Else
        lacBanner(0).Width = 0
    End If
    pbcBanner.Visible = True
    gGrid_IntegralHeight grdLines(imTokenIdx), grdLines(imTokenIdx).RowHeight(0)
    gGrid_FillWithRows grdLines(imTokenIdx), grdLines(imTokenIdx).RowHeight(0)
    mSetColumnColor -1
    
    grdLines(imTokenIdx).Rows = grdLines(imTokenIdx).Rows + 1
    grdLines(imTokenIdx).Height = grdLines(imTokenIdx).Height + 60
    
    mSetTotals

    lacTotalSpots.Top = grdLines(imTokenIdx).Top + grdLines(imTokenIdx).Height + lacTotalSpots.Height / 2
    lacTotalPrice.Top = lacTotalSpots.Top
    lacTotalCPP.Top = lacTotalSpots.Top
    lacTotalGrpPct.Top = lacTotalSpots.Top
    lacTotalGRP.Top = lacTotalSpots.Top
    
    lacComment.Top = lacTotalSpots.Top + lacTotalSpots.Height + lacTotalSpots.Height / 2
    lacComment.Left = grdLines(imTokenIdx).Left
    edcComment.Top = lacComment.Top
    edcComment.Left = lacComment.Left + lacComment.Width + 90
    ''edcComment.Width = grdLines(imTokenIdx).Width - lacComment.Left - lacComment.Width - lacTotalSpots.Width - lacTotalPrice.Width - 120
    'edcComment.Width = lacTotalSpots.Left - 2 * grdLines(imTokenIdx).Left - lacComment.Left - lacComment.Width - 120
    edcComment.Width = grdLines(imTokenIdx).Width - edcComment.Left
    edcComment.Height = 3 * grdLines(imTokenIdx).RowHeight(0) - 60
    
    
    lacButtonHelp.Width = cmcUndo.Left + cmcUndo.Width - cmcSubmit.Left + 3 * cmcSubmit.Width
    lacButtonHelp.Left = Me.Width / 2 - lacButtonHelp.Width / 2
    lacButtonHelp.Height = 3 * grdLines(imTokenIdx).RowHeight(0)  'cmcSubmit.Top - edcComment.Top - cmcSubmit.Height - 120
    lacButtonHelp.Top = cmcSubmit.Top - lacButtonHelp.Height - 60
    lacButtonHelp.Text = "SUBMIT button - Send this request to the Network." & sgCR & sgLF & "SIGNOFF button - End this program. " & sgCR & sgLF & "UNDO button - Press once to clear detail, press twice to clear header and detail."

    lacBanner(2).Caption = "Ted Bates" & sgCR & sgLF & "Request"
    lacBanner(0).Caption = sgClientName & sgCR & sgLF & "Network Programmatic Buy"
    
    mLoadLogo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tmCff
    Erase tmClf
    
    Erase tgAdfInfo
    Erase tmRif
    Erase tmRdf
    Erase imVefCode
    
    adf_rst.Close
    prf_rst.Close
    rcf_rst.Close
    rif_rst.Close
    rdf_rst.Close
    vef_rst.Close
    pvf_rst.Close
    att_rst.Close
    
    btrDestroy hmSaf
    btrDestroy hmSsf
    btrDestroy hmDrf
    btrDestroy hmMnf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmRaf
    btrDestroy hmVLF
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hgAuf
    
    tmcSetTime.Enabled = False
    gCloseTmf
    
    gEraseGlobalVar True
    btn_Stop_Click
    Set ProgrammaticBuy = Nothing   'Remove data segment
    gLogMsg "<ProgMsg> Unloading Programmatic Buy", "ProgrammaticBuy.Txt", False
    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slDate As String
    imTerminate = False
    Dim temp_rst As ADODB.Recordset
    Dim slTemp As String
    Dim slSQLQuery As String
    
    grdLines(0).Enabled = False
    grdHeader(0).Enabled = False
    
    igShowMsgBox = 0
    slTemp = ""
    Screen.MousePointer = vbHourglass
    igBkgdProg = 20
    If igExitTraffic Then
        End
    End If
    mParseCmmdLine
    If imTerminate Then
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    bmHelpOff = True
    If Screen.Width / 15 = 800 Then
        fmFontSize = 8
        bmFontBold = False
        fmTotalFontSize = 10
    ElseIf Screen.Width / 15 = 1024 Then
        fmFontSize = 10
        bmFontBold = False
        fmTotalFontSize = 10
    Else
        fmFontSize = 14
        bmFontBold = True
        fmTotalFontSize = 14
    End If

    pbcClickFocus.Left = pbcClickFocus.Width - 30
    pbcFocusStop.Left = pbcFocusStop.Width - 30
    pbcSTab.Left = pbcSTab.Width - 30
    pbcTab.Left = pbcTab.Width - 30
    imFixing = False
    imFirstFocus = True
    imBypassFocus = False
    lmHdEnableRow = -1
    lmHdEnableCol = -1
    lmLnEnableRow = -1
    lmLnEnableCol = -1
    imWhichGrid = -1
    imChgMode = False
    imAllowUndo = True
    bmIgnoreScroll = False
    
    slDate = Format$(gNow(), "ddddd")   'Get year
    lmNowDate = gDateValue(slDate)
    
    lacBanner(0) = Trim$(tgSpf.sGClient)

    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    

    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmDpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'D.S. 5/1/19 - Added below for Dick's email on 4/20/19 subject: Programmatic Buy change, 4/20/19
    'setup global variable for Demo Plus file (to see if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    hmDef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmRaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmVLF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hgAuf, "", sgDBPath & "Auf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    mGetSiteOptions
    
    'make sure all of the qualifications to run programmatic buy are met
    If smPrgmmaticAllow <> "Y" Then
        MsgBox "Please Contact Counterpoint about Programmatic Buy"
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    'mAdvtPop
    mDetermineRateCard
    mNetworkPop
    mSpotsPer
    mLengthPop
    mDemoPop
    mTypePop
    gObtainUrf
    
    'Verify that the Proposal System is turned on
    slSQLQuery = "select spfGUsePropSys, spfgClient from SPF_Site_Options"
    Set temp_rst = cnn.Execute(slSQLQuery)
    If Not temp_rst.EOF Then
        sgClientName = Trim(temp_rst!spfgClient)
        If temp_rst!spfGUsePropSys <> "Y" Then
            imTerminate = True
            slTemp = "Error: The Traffic Proposal System must be in use in order to use the Programmatic Buy System"
        End If
    End If
    
    mProgrammaticUser
    mProgrammaticSalesPerson
    slTemp = ""
    'Verify that the user if set to use programmatic buy and that they are a salesperson

    If imUrfCode = 0 Then
        'D.S. Are ther any users defined as 'allow to select programmatic buys'
        'D.S. 03/14/18 Error must have one and only one user defined to 'allow to select programmatic buys' needs to be green
        imTerminate = True
        slTemp = "Error: Currently, there's no user assigned to Programmatic Buys.  There must be one and only one user defined to 'allow to select programmatic buys' needs to be green"
    End If
    
    If imSlfCode = 0 Then
        'D.S. 03/14/18 Error must have one and only one user defined to 'allow to select programmatic buys' needs to be green
        If slTemp <> "" Then
            slTemp = slTemp & vbCrLf & "Error: Currently no salesperson is defined with the name 'Programmatic Buy'."
        Else
            slTemp = "Error: Currently no salesperson is defined with the name 'Programmatic Buy'."
        End If
        imTerminate = True
    End If
    
    If sgCommServerIP = "" Then
        If slTemp <> "" Then
            slTemp = slTemp & vbCrLf & "Error: Currently there is no definition for CommServerIP in Traffic.ini"
        Else
            slTemp = "Error: Currently there is no definition for CommServerIP in Traffic.ini"
        
        End If
        imTerminate = True
    End If
    
    If sgCommServerPort = "" Then
        If slTemp <> "" Then
            slTemp = slTemp & vbCrLf & "Error: Currently there is no definition for CommServerPort in Traffic.ini"
        Else
            slTemp = "Error: Currently there is no definition for CommServerPort in Traffic.ini"
        
        End If
        imTerminate = True
    End If

    'Verify that at least one vehicle is setup as Programmatic Buys
    If Not bmProgrammaticVehExist Then
        If slTemp <> "" Then
            slTemp = slTemp & vbCrLf & "Error: Currently there is no Package defined as Programmatic Buys."
        Else
            slTemp = "Error: Currently there is no Package defined as Programmatic Buys."
        End If
        imTerminate = True
    End If
    
    If Not bmProgrammaticRifExist Then
        If slTemp <> "" Then
            slTemp = slTemp & vbCrLf & "Error: Currently there are no Rate Card Items defined for the Programmatic Buy Package."
        Else
            slTemp = "Error: Currently there are no Ratecard Items defined for the Programmatic Buy Package."
        End If
'D.S. 4/11/20        imTerminate = True
    End If

    If imTerminate Then
        gLogMsg slTemp, "ProgrammaticBuy.Txt", False
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    imTabIndex = 0
    mHeaderColumnNames
    mLineColumnNames
    mSetTitles
    
    'Me.Width = (Screen.Width - 360) / (640 * 15 / Me.Width)
    'Me.Height = (Screen.Height - 500) / (480 * 15 / Me.Height)
    If Screen.Width / 15 = 800 Then
        lgPercentAdjH = 100
        lgPercentAdjW = 100
        fmFontSize = 8
        bmFontBold = False
        fmTotalFontSize = 10
    ElseIf Screen.Width / 15 = 1024 Then
        lgPercentAdjH = 95
        lgPercentAdjW = 100
        fmFontSize = 10
        bmFontBold = False
        fmTotalFontSize = 10
    Else
        lgPercentAdjH = 80
        lgPercentAdjW = 95
        fmFontSize = 14
        bmFontBold = True
        fmTotalFontSize = 14
    End If
    If Screen.Width / 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = (((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = ((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    
    mSetCtrlProperities
    
    gCenterStdAlone Me

    ilRet = mOpenFiles
    If Not ilRet Then
        imTerminate = True
    End If
    
    mSetCommands
    mPopRateGp
    mRemDormantVeh
    
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


Private Sub mNetworkPop()

    'Populate Vehicle selection box
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilIdx As Integer
    Dim slSQLQuery As String
    
    bmProgrammaticRifExist = False
    bmProgrammaticVehExist = False
    cbcNetwork.Clear
    ilRet = gVffRead()
    slSQLQuery = "select distinct vefName, vefCode from rif_Rate_Card_Items Left Outer Join vef_Vehicles On rifVefCode = vefCode Where rifRcfCode = " & imRcfCode & " And vefType = 'P' And vefPvfCode > 0 and vefState <> 'D' Order by vefName"
    Set vef_rst = cnn.Execute(slSQLQuery)
    Do While Not vef_rst.EOF
        bmProgrammaticVehExist = True
        ilVff = gBinarySearchVff(vef_rst!vefCode)
        If ilVff <> -1 Then
            If tgVff(ilVff).sPrgmmaticAllow = "Y" Then
                For ilIdx = 0 To UBound(tmRif) - 1 Step 1
                    If tmRif(ilIdx).iVefCode = vef_rst!vefCode Then
                        cbcNetwork.AddItem Trim$(vef_rst!VEFNAME)
                        cbcNetwork.SetItemData = vef_rst!vefCode
                        bmProgrammaticRifExist = True
                        Exit For
                    End If
                Next ilIdx
            End If
        End If
        vef_rst.MoveNext
    Loop
End Sub


Private Sub grdHeader_GotFocus(Index As Integer)
    If imWhichGrid = 1 Then
        mSetShow
    End If
End Sub

Private Sub grdHeader_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer
    Dim llSdfCode As Long
    Dim llTstSdfCode As Long

    ilCol = grdHeader(imTokenIdx).MouseCol
    ilRow = grdHeader(imTokenIdx).MouseRow
    If ilCol < grdHeader(imTokenIdx).FixedCols Then
        Exit Sub
    End If
    If ilRow < grdHeader(imTokenIdx).FixedRows Then
        Exit Sub
    End If
    'If ilRow = grdHeader(imTokenIdx).FixedRows Then
    '    If grdHeader(imTokenIdx).TextMatrix(ilRow, ADVERTISERINDEX) = "" Then
    '        Exit Sub
    '    End If
    'End If
    If ilRow <> lmHdEnableRow Then
        'mSetShow True
        'If ilRow = grdHeader(imTokenIdx).FixedRows Then
        '    If grdHeader(imTokenIdx).TextMatrix(ilRow, ADVERTISERINDEX) = "" Then
        '        Exit Sub
        '    End If
        'End If

        If ilRow >= grdHeader(imTokenIdx).Rows Then
            ilRow = grdHeader(imTokenIdx).Rows - 1
        End If
    End If
    If ilRow >= grdHeader(imTokenIdx).Rows Then
        Exit Sub
    End If
    mSetShow
    grdHeader(imTokenIdx).Col = ilCol
    grdHeader(imTokenIdx).Row = ilRow
    lmHdEnableRow = grdHeader(imTokenIdx).Row
    lmHdEnableCol = grdHeader(imTokenIdx).Col
    imWhichGrid = 0
    mEnableBox
End Sub

Private Sub grdHeader_Scroll(Index As Integer)
    If (imCtrlVisible) And (grdHeader(imTokenIdx).Row >= grdHeader(imTokenIdx).FixedRows) And (grdHeader(imTokenIdx).Col >= grdHeader(imTokenIdx).FixedCols) And (imWhichGrid = 0) Then
        If grdHeader(imTokenIdx).RowIsVisible(grdHeader(imTokenIdx).Row) Then
            mSetFocus
        Else
            cbcAdvt.Visible = False
            cbcProduct.Visible = False
        End If
    End If

End Sub

Private Sub grdLines_GotFocus(Index As Integer)
    If imWhichGrid = 0 Then
        mSetShow
    End If
End Sub

Private Sub grdLines_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer
    Dim llSdfCode As Long
    Dim llTstSdfCode As Long
    Dim slType As String
    Dim slTstType As String
    Dim ilRet As Integer

    ilCol = grdLines(imTokenIdx).MouseCol
    ilRow = grdLines(imTokenIdx).MouseRow
    If ilCol < grdLines(imTokenIdx).FixedCols Then
        Exit Sub
    End If
    If ilRow < grdLines(imTokenIdx).FixedRows Then
        Exit Sub
    End If
'    If (ilRow <> lmLnEnableRow) Then
'        mSetShow
'        If ilRow >= grdLines(imTokenIdx).Rows Then
'            ilRow = grdLines(imTokenIdx).Rows - 1
'        End If
'    Else
'        mSetShow
'    End If
    If ilRow >= grdLines(imTokenIdx).Rows Then
        Exit Sub
    End If
    mSetShow
    If (ilCol = SPOTPRICEINDEX) Or (ilCol = AVAILINDEX) Or (ilCol = TOTALPRICEINDEX) Then
        Exit Sub
    End If
    If (ilCol = NETWORKINDEX) Or (ilCol = DAYPARTINDEX) Or (ilCol = STATIONINDEX) Then
        Exit Sub
    End If
    If (ilCol = CPPCPMINDEX) Or (ilCol = LNGRPINDEX) Then
        Exit Sub
    End If
    If (ilCol = LENGTHINDEX) And (grdLines(imTokenIdx).TextMatrix(ilRow, ilCol) <> "") And (grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX) <> "") Then
        Exit Sub
    End If
    If (ilCol <> LENGTHINDEX) And (grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX) = "") Then
        Exit Sub
    End If
    If (ilCol = LENGTHINDEX) And (ilRow > grdLines(imTokenIdx).FixedRows) And (grdLines(imTokenIdx).TextMatrix(ilRow - 1, LENGTHINDEX) = "") Then
        Exit Sub
    End If
    If (ilCol = LENGTHINDEX) And (ilRow > grdLines(imTokenIdx).FixedRows) And (grdLines(imTokenIdx).TextMatrix(ilRow - 1, LENGTHINDEX) <> "") And (grdLines(imTokenIdx).TextMatrix(ilRow - 1, NETWORKINDEX) = "") Then
        Exit Sub
    End If
    grdLines(imTokenIdx).Col = ilCol
    grdLines(imTokenIdx).Row = ilRow
    imWhichGrid = 1
    lmLnEnableRow = grdLines(imTokenIdx).Row
    lmLnEnableCol = grdLines(imTokenIdx).Col
    mEnableBox
End Sub

Private Sub grdLines_Scroll(Index As Integer)
    If bmIgnoreScroll Then
        Exit Sub
    End If
    If (imCtrlVisible) And (grdLines(imTokenIdx).Row >= grdLines(imTokenIdx).FixedRows) And (grdLines(imTokenIdx).Col >= grdLines(imTokenIdx).FixedCols) And (imWhichGrid = 1) Then
        mSetFocus
        pbcClickFocus.SetFocus
    Else
        pbcClickFocus.SetFocus
    End If
End Sub


Private Sub mHeaderColumnNames()
    grdHeader(imTokenIdx).Row = 0
    grdHeader(imTokenIdx).RowHeight(0) = grdHeader(imTokenIdx).RowHeight(1)
    
    grdHeader(imTokenIdx).Col = ADVERTISERINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, ADVERTISERINDEX) = "Advertiser"
    grdHeader(imTokenIdx).ColAlignment(ADVERTISERINDEX) = flexAlignLeftCenter
    
    grdHeader(imTokenIdx).Col = PRODUCTINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, PRODUCTINDEX) = "Product"
    grdHeader(imTokenIdx).ColAlignment(PRODUCTINDEX) = flexAlignLeftCenter
    
    grdHeader(imTokenIdx).Col = STARTDATEINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, STARTDATEINDEX) = "Start Date"
    grdHeader(imTokenIdx).ColAlignment(STARTDATEINDEX) = flexAlignRightCenter
    
    grdHeader(imTokenIdx).Col = NUMBERWEEKSINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, NUMBERWEEKSINDEX) = "# Weeks"
    grdHeader(imTokenIdx).ColAlignment(NUMBERWEEKSINDEX) = flexAlignRightCenter
    
    grdHeader(imTokenIdx).Col = ENDDATEINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, ENDDATEINDEX) = "End Date"
    grdHeader(imTokenIdx).ColAlignment(ENDDATEINDEX) = flexAlignRightCenter
    
    grdHeader(imTokenIdx).Col = ESTIMATENUMBERINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, ESTIMATENUMBERINDEX) = "Estimate #"
    grdHeader(imTokenIdx).ColAlignment(ESTIMATENUMBERINDEX) = flexAlignRightCenter

    grdHeader(imTokenIdx).Col = DEMOINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    grdHeader(imTokenIdx).TextMatrix(0, DEMOINDEX) = "Demo"
    grdHeader(imTokenIdx).ColAlignment(DEMOINDEX) = flexAlignLeftCenter

    grdHeader(imTokenIdx).Col = TARGETINDEX
    grdHeader(imTokenIdx).CellFontBold = bmFontBold
    grdHeader(imTokenIdx).CellFontName = "Arial"
    grdHeader(imTokenIdx).CellFontSize = fmFontSize
    'grdHeader(imTokenIdx).CellForeColor = vbBlue
    If imTabIndex = 0 Then
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = "Target GRP's"
    ElseIf imTabIndex = 1 Then
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = "Target GrImp's"
    Else
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = ""
    End If
    grdHeader(imTokenIdx).ColAlignment(TARGETINDEX) = flexAlignRightCenter

End Sub

Private Sub mHeaderColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long
    
    If imTabIndex = 0 Or imTabIndex = 1 Then
        grdHeader(imTokenIdx).ColWidth(PRODUCTINDEX) = 0.11 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(STARTDATEINDEX) = 0.08 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(NUMBERWEEKSINDEX) = 0.07 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(ENDDATEINDEX) = 0.08 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(ESTIMATENUMBERINDEX) = 0.09 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(DEMOINDEX) = 0.08 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(TARGETINDEX) = 0.12 * grdHeader(imTokenIdx).Width
    Else
        grdHeader(imTokenIdx).ColWidth(PRODUCTINDEX) = 0.25 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(STARTDATEINDEX) = 0.11 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(NUMBERWEEKSINDEX) = 0.1 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(ENDDATEINDEX) = 0.1 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(ESTIMATENUMBERINDEX) = 0.11 * grdHeader(imTokenIdx).Width
        grdHeader(imTokenIdx).ColWidth(DEMOINDEX) = 0
        grdHeader(imTokenIdx).ColWidth(TARGETINDEX) = 0
    End If

    llWidth = 45    'GRIDSCROLLWIDTH + 45
    For ilCol = 0 To grdHeader(imTokenIdx).Cols - 1 Step 1
        If ilCol <> ADVERTISERINDEX Then
            llWidth = llWidth + grdHeader(imTokenIdx).ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdHeader(imTokenIdx).Width - llWidth
    If llWidth > 0 Then
        grdHeader(imTokenIdx).ColWidth(ADVERTISERINDEX) = llWidth - 30
    End If

End Sub

Private Sub mLineColumnNames()
    grdLines(imTokenIdx).Row = 0
    
    grdLines(imTokenIdx).Col = LENGTHINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, LENGTHINDEX) = "Length"
    grdLines(imTokenIdx).ColAlignment(LENGTHINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).RowHeight(0) = grdLines(imTokenIdx).RowHeight(1)
    grdLines(imTokenIdx).Col = NETWORKINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, NETWORKINDEX) = "Network"
    grdLines(imTokenIdx).ColAlignment(NETWORKINDEX) = flexAlignLeftCenter
    
    grdLines(imTokenIdx).Col = DAYPARTINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, DAYPARTINDEX) = "Daypart"
    grdLines(imTokenIdx).ColAlignment(DAYPARTINDEX) = flexAlignLeftCenter
    
    grdLines(imTokenIdx).Col = STATIONINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, STATIONINDEX) = "Stations"
    grdLines(imTokenIdx).ColAlignment(STATIONINDEX) = flexAlignRightCenter
        
    grdLines(imTokenIdx).Col = SPOTPRICEINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, SPOTPRICEINDEX) = "Spot Price"
    grdLines(imTokenIdx).ColAlignment(SPOTPRICEINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = AVAILINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, AVAILINDEX) = "Avails"
    grdLines(imTokenIdx).ColAlignment(AVAILINDEX) = flexAlignLeftCenter
    
    grdLines(imTokenIdx).Col = CPPCPMINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, CPPCPMINDEX) = ""
    grdLines(imTokenIdx).ColAlignment(CPPCPMINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = GRPGRIMPPCTINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, GRPGRIMPPCTINDEX) = ""
    grdLines(imTokenIdx).ColAlignment(GRPGRIMPPCTINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = NUMBERSPOTSINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, NUMBERSPOTSINDEX) = "# Spots"
    grdLines(imTokenIdx).ColAlignment(NUMBERSPOTSINDEX) = flexAlignRightCenter
        
    grdLines(imTokenIdx).Col = SPOTSPERINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, SPOTSPERINDEX) = "Per"
    grdLines(imTokenIdx).ColAlignment(SPOTSPERINDEX) = flexAlignLeftCenter
                
    grdLines(imTokenIdx).Col = TOTALPRICEINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, TOTALPRICEINDEX) = "Total Price"
    grdLines(imTokenIdx).ColAlignment(TOTALPRICEINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = LNGRPINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, LNGRPINDEX) = ""
    grdLines(imTokenIdx).ColAlignment(LNGRPINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = AUDIENCEINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, AUDIENCEINDEX) = "Audience"
    grdLines(imTokenIdx).ColAlignment(AUDIENCEINDEX) = flexAlignRightCenter

    grdLines(imTokenIdx).Col = POPULATIONINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, POPULATIONINDEX) = "Population"
    grdLines(imTokenIdx).ColAlignment(POPULATIONINDEX) = flexAlignRightCenter
    
    grdLines(imTokenIdx).Col = RATINGINDEX
    grdLines(imTokenIdx).CellFontBold = bmFontBold
    grdLines(imTokenIdx).CellFontName = "Arial"
    grdLines(imTokenIdx).CellFontSize = fmFontSize
    'grdLines(imTokenIdx).CellForeColor = vbBlue
    grdLines(imTokenIdx).TextMatrix(0, RATINGINDEX) = "Rating"
    grdLines(imTokenIdx).ColAlignment(RATINGINDEX) = flexAlignRightCenter
    
End Sub

Private Sub mLineColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long

    If (imTabIndex = 0) Or (imTabIndex = 1) Then
        grdLines(imTokenIdx).ColWidth(AUDIENCEINDEX) = 0
        grdLines(imTokenIdx).ColWidth(RATINGINDEX) = 0
        grdLines(imTokenIdx).ColWidth(POPULATIONINDEX) = 0
        grdLines(imTokenIdx).ColWidth(SPOTPRICEINDEX) = 0
        grdLines(imTokenIdx).ColWidth(SPOTSPERINDEX) = 0
        grdLines(imTokenIdx).ColWidth(LENGTHINDEX) = 0.06 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(DAYPARTINDEX) = 0.15 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(STATIONINDEX) = 0.07 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(AVAILINDEX) = 0.07 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(CPPCPMINDEX) = 0.1 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(GRPGRIMPPCTINDEX) = 0.07 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(LNGRPINDEX) = 0.06 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(NUMBERSPOTSINDEX) = 0.065 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(TOTALPRICEINDEX) = 0.1 * grdLines(imTokenIdx).Width
    Else
        grdLines(imTokenIdx).ColWidth(AUDIENCEINDEX) = 0
        grdLines(imTokenIdx).ColWidth(RATINGINDEX) = 0
        grdLines(imTokenIdx).ColWidth(POPULATIONINDEX) = 0
        grdLines(imTokenIdx).ColWidth(CPPCPMINDEX) = 0
        grdLines(imTokenIdx).ColWidth(GRPGRIMPPCTINDEX) = 0
        grdLines(imTokenIdx).ColWidth(LNGRPINDEX) = 0
        grdLines(imTokenIdx).ColWidth(LENGTHINDEX) = 0.08 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(DAYPARTINDEX) = 0.2 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(SPOTPRICEINDEX) = 0.11 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(STATIONINDEX) = 0.07 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(AVAILINDEX) = 0.07 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(NUMBERSPOTSINDEX) = 0.1 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(SPOTSPERINDEX) = 0.06 * grdLines(imTokenIdx).Width
        grdLines(imTokenIdx).ColWidth(TOTALPRICEINDEX) = 0.12 * grdLines(imTokenIdx).Width
    End If

    llWidth = GRIDSCROLLWIDTH + 45
    For ilCol = 0 To grdLines(imTokenIdx).Cols - 1 Step 1
        If ilCol <> NETWORKINDEX Then
            llWidth = llWidth + grdLines(imTokenIdx).ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdLines(imTokenIdx).Width - llWidth
    If llWidth > 0 Then
        grdLines(imTokenIdx).ColWidth(NETWORKINDEX) = llWidth
    End If

End Sub






'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If imWhichGrid = 0 Then
        If mGridEnable(grdHeader(imTokenIdx)) Then
            'D.S. 10/30/18 commented out set focus below
            'mGridSetFocus grdHeader(imTokenIdx)
        End If
    ElseIf imWhichGrid = 1 Then
        If mGridEnable(grdHeader(imTokenIdx)) Then
            'D.S. 10/30/18 commented out set focus below
            'mGridSetFocus grdLines(imTokenIdx)
        End If
    End If
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
    Dim ilRet As Integer

    If imWhichGrid = 0 Then
        ilRet = mGridSetFocus(grdHeader(imTokenIdx))
    ElseIf imWhichGrid = 1 Then
        ilRet = mGridSetFocus(grdHeader(imTokenIdx))
    End If
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
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilRet As Integer

    If imWhichGrid = 0 Then
        ilRet = mGridSetShow(grdHeader(imTokenIdx))
    ElseIf imWhichGrid = 1 Then
        ilRet = mGridSetShow(grdLines(imTokenIdx))
    End If
End Sub


Private Sub imcLogo_Click()
    mSetShow
End Sub

Private Sub lacBanner_Click(Index As Integer)
    mSetShow
End Sub

Private Sub cmcHelp_Click()
'    mSetShow

'    igHelpScreen = 1
'    'If pbcProposal.Visible Then
'        igHelpScreen = 2
'    'End If
'    ProgrammaticBuyHelp.Show vbModal
    If cmcHelp.Caption = "&Help (Off)" Then
        cmcHelp.Caption = "&Help (On)"
        bmHelpOff = False
        mShowHelp
    Else
        bmHelpOff = True
        lacButtonHelp.Visible = False
        lacBanner(1).Visible = False
        cmcHelp.Caption = "&Help (Off)"
    End If
    If imCtrlVisible Then
        mSetFocus
    End If
End Sub

Private Sub lacComment_Click()
    mSetShow
End Sub

Private Sub lacCopyright_Click()
    mSetShow
End Sub

Private Sub lacSpacer_Click()
    mSetShow
End Sub

Private Sub lacTotalCPP_Click()
    mSetShow
End Sub

Private Sub lacTotalGRP_Click()
    mSetShow
End Sub

Private Sub lacTotalGrpPct_Click()
    mSetShow
End Sub

Private Sub lacTotalPrice_Click()
    mSetShow
End Sub

Private Sub lacTotalSpots_Click()
    mSetShow
End Sub

Private Sub pbcBanner_Click()
    mSetShow
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
'    grdHeader(imTokenIdx).Col = 0
'    grdHeader(imTokenIdx).Row = 0
'    grdLines(imTokenIdx).Col = 0
'    grdLines(imTokenIdx).Row = 0
    cbcAdvt.Visible = False
    cbcProduct.Visible = False
    lmHdEnableRow = -1
    lmHdEnableCol = -1
    
    cbcNetwork.Visible = False
    cbcDaypart.Visible = False
    cbcLength.Visible = False
    lmLnEnableRow = -1
    lmLnEnableCol = -1
    
    edcData.Visible = False
End Sub

Private Sub pbcFocusStop_GotFocus()
    mSetFocus
End Sub


Private Sub pbcSTab_GotFocus()
    Dim slStr As String

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imWhichGrid = -9 Then
        imWhichGrid = 1
        grdLines(imTokenIdx).Row = grdLines(imTokenIdx).FixedRows
        grdLines(imTokenIdx).Col = LENGTHINDEX
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        mSetShow
        If imWhichGrid = 0 Then
            If grdHeader(imTokenIdx).Col = ADVERTISERINDEX Then
                Exit Sub
            Else
                grdHeader(imTokenIdx).Col = grdHeader(imTokenIdx).Col - 1
                mEnableBox
            End If
        Else
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilEnableRow As Integer
    Dim ilEnableCol As Integer
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        If imWhichGrid = 0 Then
            ilEnableRow = lmHdEnableRow
            ilEnableCol = lmHdEnableCol
        Else
            ilEnableRow = lmLnEnableRow
            ilEnableCol = lmLnEnableCol
        End If
        mSetShow
        If imWhichGrid = 0 Then
            If (ilEnableCol = TARGETINDEX) Or ((imTabIndex = 2) And (ilEnableCol = ESTIMATENUMBERINDEX)) Then
                imWhichGrid = -9
                pbcSTab.SetFocus
                Exit Sub
            Else
                grdHeader(imTokenIdx).Col = grdHeader(imTokenIdx).Col + 1
                mEnableBox
            End If
        Else
            If ((imTabIndex = 0) Or (imTabIndex = 1)) And ((ilEnableCol = GRPGRIMPPCTINDEX) Or (grdLines(imTokenIdx).TextMatrix(ilEnableRow, ilEnableCol) = "")) Then
                'Advance to next row
                If grdLines(imTokenIdx).Row + 1 >= grdLines(imTokenIdx).Rows Then
                    grdLines(imTokenIdx).Rows = grdLines(imTokenIdx).Rows + 1
                End If
                grdLines(imTokenIdx).Row = grdLines(imTokenIdx).Row + 1
                grdLines(imTokenIdx).Col = GRPGRIMPPCTINDEX
                'Test if row visible
                bmIgnoreScroll = True
                If Not grdLines(imTokenIdx).RowIsVisible(grdLines(imTokenIdx).Row) Then
                    grdLines(imTokenIdx).TopRow = grdLines(imTokenIdx).TopRow + 1
                End If
                If grdLines(imTokenIdx).RowPos(grdLines(imTokenIdx).Row) + grdLines(imTokenIdx).RowHeight(grdLines(imTokenIdx).Row) > grdLines(imTokenIdx).Height Then
                    grdLines(imTokenIdx).TopRow = grdLines(imTokenIdx).TopRow + 1
                End If
                bmIgnoreScroll = False
                If grdLines(imTokenIdx).TextMatrix(grdLines(imTokenIdx).Row, LENGTHINDEX) <> "" Then
                    mEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            ElseIf (imTabIndex = 2) And ((ilEnableCol = SPOTSPERINDEX) Or (grdLines(imTokenIdx).TextMatrix(ilEnableRow, ilEnableCol) = "")) Then
                'Advance to next row
                If grdLines(imTokenIdx).Row + 1 >= grdLines(imTokenIdx).Rows Then
                    grdLines(imTokenIdx).Rows = grdLines(imTokenIdx).Rows + 1
                End If
                grdLines(imTokenIdx).Row = grdLines(imTokenIdx).Row + 1
                grdLines(imTokenIdx).Col = NUMBERSPOTSINDEX
                'Test if row visible
                bmIgnoreScroll = True
                If Not grdLines(imTokenIdx).RowIsVisible(grdLines(imTokenIdx).Row) Then
                    grdLines(imTokenIdx).TopRow = grdLines(imTokenIdx).TopRow + 1
                End If
                If grdLines(imTokenIdx).RowPos(grdLines(imTokenIdx).Row) + grdLines(imTokenIdx).RowHeight(grdLines(imTokenIdx).Row) > grdLines(imTokenIdx).Height Then
                    grdLines(imTokenIdx).TopRow = grdLines(imTokenIdx).TopRow + 1
                End If
                bmIgnoreScroll = False
                If grdLines(imTokenIdx).TextMatrix(grdLines(imTokenIdx).Row, LENGTHINDEX) <> "" Then
                    mEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            ElseIf ilEnableCol = LENGTHINDEX Then
                If (imTabIndex = 0) Or (imTabIndex = 1) Then
                    grdLines(imTokenIdx).Col = GRPGRIMPPCTINDEX
                Else
                    grdLines(imTokenIdx).Col = NUMBERSPOTSINDEX
                End If
                mEnableBox
            Else
                grdLines(imTokenIdx).Col = grdLines(imTokenIdx).Col + 1
                mEnableBox
            End If
        End If
    End If
End Sub



Private Function mGridEnable(grdGrid As MSHFlexGrid) As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAdfCode As Integer
    Dim ilAdf As Integer
    Dim ilVefCode As Integer
    Dim ilRet As Integer
    Dim blFd As Boolean
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim ilRif As Integer
    Dim ilDp As Integer
    Dim slNetwork As String
    Dim slDayPart As String
    
    If (grdGrid.Row < grdGrid.FixedRows) Or (grdGrid.Row >= grdGrid.Rows) Or (grdGrid.Col < grdGrid.FixedCols) Or (grdGrid.Col >= grdGrid.Cols) Then
        mGridEnable = False
        Exit Function
    End If
    If imWhichGrid = 0 Then
        lmHdEnableRow = grdGrid.Row
        lmHdEnableCol = grdGrid.Col
        Select Case lmHdEnableCol
            Case ADVERTISERINDEX
                cbcAdvt.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
                If (cbcAdvt.ListCount > 0) And (Trim$(cbcAdvt.Text) = "") Then
                    cbcAdvt.Text = cbcAdvt.GetName(0)
                End If
            Case PRODUCTINDEX
                ilAdfCode = cbcAdvt.GetItemData(cbcAdvt.ListIndex)
                cbcProduct.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
                mProdPop ilAdfCode
                If (cbcProduct.ListCount > 0) And (Trim$(cbcProduct.Text) = "") Then
                    cbcProduct.Text = cbcProduct.GetName(0)
                End If
            Case STARTDATEINDEX
                If grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol) = "" Then
                    dtcStartDate.Text = gObtainNextMonday(Format(gNow(), "ddddd"))
                Else
                    dtcStartDate.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
                End If
            Case NUMBERWEEKSINDEX
                edcData.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
            Case ENDDATEINDEX
                If grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol) = "" Then
                    dtcEndDate.Text = gObtainNextSunday(grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, STARTDATEINDEX))
                Else
                    dtcEndDate.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
                End If
                
            Case ESTIMATENUMBERINDEX
                edcData.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
            Case DEMOINDEX
                cbcDemo.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
                If (cbcDemo.ListCount > 0) And (Trim$(cbcDemo.Text) = "") Then
                    ilAdfCode = cbcAdvt.GetItemData(cbcAdvt.ListIndex)
                    For ilAdf = 0 To UBound(tgAdfInfo) - 1 Step 1
                        If tgAdfInfo(ilAdf).iCode = ilAdfCode Then
                            cbcDemo.Text = Trim$(tgAdfInfo(ilAdf).sDemo)
                            Exit For
                        End If
                    Next ilAdf
                End If
            Case TARGETINDEX
                edcData.Text = grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, lmHdEnableCol)
        End Select
    Else
        lmLnEnableRow = grdGrid.Row
        lmLnEnableCol = grdGrid.Col
        Select Case lmLnEnableCol
            Case LENGTHINDEX
                cbcLength.Text = grdLines(imTokenIdx).TextMatrix(lmLnEnableRow, lmLnEnableCol)
                If (cbcLength.ListCount > 0) And (Trim$(cbcLength.Text = "")) Then
                    cbcLength.Text = cbcLength.GetName(0)
                End If
            Case NUMBERSPOTSINDEX
                edcData.Text = grdLines(imTokenIdx).TextMatrix(lmLnEnableRow, lmLnEnableCol)
            Case GRPGRIMPPCTINDEX
                    edcData.Text = grdLines(imTokenIdx).TextMatrix(lmLnEnableRow, lmLnEnableCol)
            Case SPOTSPERINDEX
                cbcPer.Text = grdLines(imTokenIdx).TextMatrix(lmLnEnableRow, lmLnEnableCol)
                If (cbcPer.ListCount > 0) And (Trim$(cbcPer.Text = "")) Then
                    cbcPer.Text = cbcPer.GetName(1)
                End If
        End Select
    End If
    mGridEnable = True
End Function

Private Function mGridSetFocus(grdGrid As MSHFlexGrid) As Integer
    Dim llColPos As Long
    Dim ilCol As Integer

    If imTerminate Then
        mGridSetFocus = False
        Exit Function
    End If
    If (grdGrid.Row < grdGrid.FixedRows) Or (grdGrid.Row >= grdGrid.Rows) Or (grdGrid.Col < grdGrid.FixedCols) Or (grdGrid.Col >= grdGrid.Cols) Then
        mGridSetFocus = False
        Exit Function
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdGrid.Col - 1 Step 1
        llColPos = llColPos + grdGrid.ColWidth(ilCol)
    Next ilCol
    If imWhichGrid = 0 Then
        'mSetRowColor lmHdEnableRow, GRAY
        Select Case lmHdEnableCol
            Case ADVERTISERINDEX
                cbcAdvt.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                cbcAdvt.SetDropDownWidth cbcAdvt.Width
                cbcAdvt.Visible = True
                cbcAdvt.SetFocus
            Case PRODUCTINDEX
                cbcProduct.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                cbcProduct.SetDropDownWidth cbcProduct.Width
                cbcProduct.Visible = True
                cbcProduct.SetFocus
            Case STARTDATEINDEX
                dtcStartDate.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                dtcStartDate.Visible = True
                dtcStartDate.SetFocus
            Case NUMBERWEEKSINDEX
                edcData.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                edcData.MaxLength = 3
                edcData.Visible = True
                edcData.SetFocus
            Case ENDDATEINDEX
                dtcEndDate.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                dtcEndDate.Visible = True
                dtcEndDate.SetFocus
            Case ESTIMATENUMBERINDEX
                edcData.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                edcData.MaxLength = 20
                edcData.Visible = True
                edcData.SetFocus
            Case DEMOINDEX
                cbcDemo.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                cbcDemo.SetDropDownWidth cbcDemo.Width
                cbcDemo.Visible = True
                cbcDemo.SetFocus
            Case TARGETINDEX
                edcData.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmHdEnableCol)
                If imTabIndex = 0 Then
                    edcData.MaxLength = 5
                Else
                    edcData.MaxLength = 10
                End If
                edcData.Visible = True
                edcData.SetFocus
        End Select
    Else
        mSetRowColor lmLnEnableRow, BISQUE  'LIGHTBLUE   '&HFFFF00    'LIGHTGREEN  'ORANGE  'LIGHTBLUE   'GRAY
        Select Case lmLnEnableCol
            Case LENGTHINDEX
                cbcLength.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmLnEnableCol)
                cbcLength.SetDropDownWidth cbcLength.Width
                cbcLength.Visible = True
                cbcLength.SetFocus
            Case NUMBERSPOTSINDEX
                edcData.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmLnEnableCol)
                edcData.MaxLength = 4
                edcData.Visible = True
                edcData.SetFocus
            Case GRPGRIMPPCTINDEX
                edcData.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmLnEnableCol)
                edcData.MaxLength = 3
                edcData.Visible = True
                edcData.SetFocus
            Case SPOTSPERINDEX
                cbcPer.Move grdGrid.Left + llColPos + 15, grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 15, grdGrid.ColWidth(lmLnEnableCol)
                cbcPer.SetDropDownWidth cbcPer.Width
                cbcPer.Visible = True
                cbcPer.SetFocus
        End Select
    End If
    mShowHelp
    mGridSetFocus = True
End Function

Private Function mGridSetShow(grdGrid As MSHFlexGrid) As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilUpdateCopy As Integer
    Dim ilMoveRow As Integer
    Dim ilVefCode As Integer
    Dim ilRdfCode As Integer
    Dim slDollars As String
    Dim ilIndex As Integer
    Dim slSpotDiff As String
    Dim blChg As Boolean
    Dim ilLen As Integer

    ilUpdateCopy = False
    If imWhichGrid = 0 Then
        If (lmHdEnableRow >= grdGrid.FixedRows) And (lmHdEnableRow < grdGrid.Rows) Then
            'mSetRowColor lmHdEnableRow, vbWhite
            blChg = False
            Select Case lmHdEnableCol
                Case ADVERTISERINDEX
                    cbcAdvt.Visible = False
                    'If grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) <> Trim$(cbcAdvt.Text) Then
                    If grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) <> Trim$(smHeaderAdvt) Then
                        cbcProduct.Clear
                        grdGrid.TextMatrix(lmHdEnableRow, PRODUCTINDEX) = ""
                        grdGrid.TextMatrix(lmHdEnableRow, DEMOINDEX) = ""
                    End If
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(smHeaderAdvt)
                Case PRODUCTINDEX
                    cbcProduct.Visible = False
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(smHeaderProd)
                Case STARTDATEINDEX
                    dtcStartDate.Visible = False
                    If grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) <> Trim$(dtcStartDate.Text) Then
                        blChg = True
                        'Reset End date
                        If grdGrid.TextMatrix(lmHdEnableRow, NUMBERWEEKSINDEX) <> "" Then
                            grdGrid.TextMatrix(lmHdEnableRow, ENDDATEINDEX) = gObtainNextSunday(Format(gDateValue(dtcStartDate.Text) + 7 * Val(grdGrid.TextMatrix(lmHdEnableRow, NUMBERWEEKSINDEX) - 1), "ddddd"))
                        End If
                    End If
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(dtcStartDate.Text)
                Case NUMBERWEEKSINDEX
                    edcData.Visible = False
                    If grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) <> Trim$(edcData.Text) Then
                        If grdGrid.TextMatrix(lmHdEnableRow, STARTDATEINDEX) <> "" Then
                            blChg = True
                            'Reset End date
                            grdGrid.TextMatrix(lmHdEnableRow, ENDDATEINDEX) = gObtainNextSunday(Format(gDateValue(grdGrid.TextMatrix(lmHdEnableRow, STARTDATEINDEX)) + 7 * Val(edcData.Text) - 1, "ddddd"))
                        End If
                    End If
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(edcData.Text)
                Case ENDDATEINDEX
                    dtcEndDate.Visible = False
                    If grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) <> Trim$(dtcEndDate.Text) Then
                        If grdGrid.TextMatrix(lmHdEnableRow, STARTDATEINDEX) <> "" Then
                            blChg = True
                            'Reset # Weeks date
                            If gDateValue(grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, STARTDATEINDEX)) <= gDateValue(dtcEndDate.Text) Then
                                grdGrid.TextMatrix(lmHdEnableRow, NUMBERWEEKSINDEX) = (gDateValue(dtcEndDate.Text) - gDateValue(grdHeader(imTokenIdx).TextMatrix(lmHdEnableRow, STARTDATEINDEX)) + 1) \ 7
                            End If
                        End If
                    End If
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(dtcEndDate.Text)
                Case ESTIMATENUMBERINDEX
                    edcData.Visible = False
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(edcData.Text)
                Case DEMOINDEX
                    cbcDemo.Visible = False
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(cbcDemo.Text)
                Case TARGETINDEX
                    edcData.Visible = False
                    grdGrid.TextMatrix(lmHdEnableRow, lmHdEnableCol) = Trim$(edcData.Text)
            End Select
            If blChg Then
                If grdLines(imTokenIdx).TextMatrix(grdLines(imTokenIdx).FixedRows, NETWORKINDEX) <> "" Then
                    mComputeTotals
                End If
            End If
        End If
        lmHdEnableRow = -1
        lmHdEnableCol = -1
    Else
        mSetRowColor lmLnEnableRow, vbWhite
        blChg = False
        Select Case lmLnEnableCol
            Case LENGTHINDEX
                cbcLength.Visible = False
                If grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) <> Trim$(cbcLength.Text) Then
                    grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) = Trim$(cbcLength.Text)
                    mPopLineGrid lmLnEnableRow
                End If
            Case NUMBERSPOTSINDEX
                edcData.Visible = False
                If grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) <> Trim$(edcData.Text) Then
                    blChg = True
                End If
                grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) = Trim$(edcData.Text)
            Case SPOTSPERINDEX
                cbcPer.Visible = False
                If grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) <> Trim$(cbcPer.Text) Then
                    blChg = True
                End If
                grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) = Trim$(cbcPer.Text)
            Case GRPGRIMPPCTINDEX
                edcData.Visible = False
                If grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) <> Trim$(edcData.Text) Then
                    grdGrid.TextMatrix(lmLnEnableRow, lmLnEnableCol) = Trim$(edcData.Text)
                    mComputeSpots lmLnEnableRow
                    blChg = True
                End If
        End Select
        If blChg Then
            mComputeTotals
        End If
        lmLnEnableRow = -1
        lmLnEnableCol = -1
    End If
    lacBanner(1).Visible = False
    imCtrlVisible = False
    mGridSetShow = True
    mSetCommands
End Function


Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "PB"
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload ProgrammaticBuy
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim slCSIName As String

    sgCommandStr = Command$
    slStartIn = CurDir$
    'slCSIName = ""
    
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommandStr, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    slCommand = sgCommandStr    'Command$
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    mTestPervasive
    '4/2/11: Add setting of value
    lgUlfCode = 0
    'If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Or (Trim$(sgCommandStr) = "Debug") Then
'    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
'        ProgrammaticBuySignOn.Show vbModal
'        If igExitTraffic Then
'            imTerminate = True
'            ProgrammaticBuy.Hide
'            Exit Sub
'        End If
'        slStr = sgUserName
'        sgCallAppName = "Traffic"
'    Else
'        igSportsSystem = 0
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        ilRet = gParseItem(slCommand, 3, "\", slStr)
'        igRptCallType = Val(slStr)
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'        sgUrfStamp = "~" 'Clear time stamp incase same name
'        sgUserName = Trim$(slStr)
'        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
'        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
'            slCSIName = "CSI"
'            sgUserName = "Guide"
'        End If
'        gUrfRead ProgrammaticBuy, sgUserName, True, tgUrf(), False  'Obtain user records
'        If StrComp(slCSIName, "CSI", vbTextCompare) = 0 Then
'            gExpandGuideAsUser tgUrf(0)
'        End If
'        mGetUlfCode
'    End If
    'End If
    DoEvents
'    gInitStdAlone ReportList, slStr, igTestSystem
    xgInitStdAlone
    mCheckForDate
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ProgrammaticBuy.Frm"
    
'    If tgUrf(0).sLiveLogPostOnly <> "Y" Then
'        If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Or (StrComp(slCSIName, "CSI", vbTextCompare) = 0) Then
'            igSportsSystem = 2
'        Else
'            MsgBox "You are not authorized to use Live Log", vbOKOnly + vbExclamation, "Disallowed"
'            imTerminate = True
'        End If
'    Else
'        igSportsSystem = 2
'    End If
End Sub

Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    xgInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub
Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer
    
    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/1999"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
    
End Sub
'4/2/11: Add routine
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "ddddd"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub
Private Sub mSetColumnColor(llInRow As Long)
    Dim llRow As Integer
    'For llRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
    '    If (llInRow = -1) Or (llInRow = llRow) Then
    '        grdLines(imTokenIdx).Row = llRow
    '        grdLines(imTokenIdx).Col = SPOTPRICEINDEX
    '        grdLines(imTokenIdx).CellBackColor = LIGHTYELLOW
     '       grdLines(imTokenIdx).Col = TOTALPRICEINDEX
     '       grdLines(imTokenIdx).CellBackColor = LIGHTYELLOW
    '    End If
    'Next llRow
End Sub

Private Sub mSetRowColor(llRow As Long, llColor As Long)
    Dim llCol As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    If llRow < 0 Then
        Exit Sub
    End If
    If imWhichGrid = 0 Then
        llSvRow = grdHeader(imTokenIdx).Row
        llSvCol = grdHeader(imTokenIdx).Col
        grdHeader(imTokenIdx).Row = llRow
        For llCol = ADVERTISERINDEX To TARGETINDEX Step 1
            grdHeader(imTokenIdx).Col = llCol
            grdHeader(imTokenIdx).CellBackColor = llColor
        Next llCol
        grdHeader(imTokenIdx).Row = llSvRow
        grdHeader(imTokenIdx).Col = llSvCol
    Else
        llSvRow = grdLines(imTokenIdx).Row
        llSvCol = grdLines(imTokenIdx).Col
        grdLines(imTokenIdx).Row = llRow
        For llCol = LENGTHINDEX To POPULATIONINDEX Step 1
            grdLines(imTokenIdx).Col = llCol
            'If grdLines(imTokenIdx).CellBackColor <> LIGHTYELLOW Then
                grdLines(imTokenIdx).Col = llCol
                grdLines(imTokenIdx).CellBackColor = llColor
            'End If
        Next llCol
        grdLines(imTokenIdx).Row = llSvRow
        grdLines(imTokenIdx).Col = llSvCol
    End If
End Sub

Private Sub mSetCtrlProperities()
    cbcAdvt.FontSize = fmFontSize
    cbcAdvt.BackColor = &HFFFF00
    cbcAdvt.Height = grdHeader(imTokenIdx).RowHeight(1)
    
    cbcProduct.FontSize = fmFontSize
    cbcProduct.BackColor = &HFFFF00
    cbcProduct.Height = grdHeader(imTokenIdx).RowHeight(1)
    
    dtcStartDate.FontSize = fmFontSize
    dtcStartDate.BackColor = &HFFFF00
    dtcStartDate.Height = grdHeader(imTokenIdx).RowHeight(1)
    
    edcData.FontSize = fmFontSize
    edcData.BackColor = &HFFFF00
    edcData.Height = grdHeader(imTokenIdx).RowHeight(1)
    
    dtcEndDate.FontSize = fmFontSize
    dtcEndDate.BackColor = &HFFFF00
    dtcEndDate.Height = grdHeader(imTokenIdx).RowHeight(1)

    cbcNetwork.FontSize = fmFontSize
    cbcNetwork.BackColor = &HFFFF00
    cbcNetwork.Height = grdLines(imTokenIdx).RowHeight(1)

    cbcDaypart.FontSize = fmFontSize
    cbcDaypart.BackColor = &HFFFF00
    cbcDaypart.Height = grdLines(imTokenIdx).RowHeight(1)

    cbcLength.FontSize = fmFontSize
    cbcLength.BackColor = &HFFFF00
    cbcLength.Height = grdLines(imTokenIdx).RowHeight(1)
    
    cbcPer.FontSize = fmFontSize
    cbcPer.BackColor = &HFFFF00
    cbcPer.Height = grdLines(imTokenIdx).RowHeight(1)
    
    cbcDemo.FontSize = fmFontSize
    cbcDemo.BackColor = &HFFFF00
    cbcDemo.Height = grdHeader(imTokenIdx).RowHeight(1)
    
    cbcType.FontSize = fmFontSize
    cbcType.BackColor = &HFFFF00
    cbcType.Height = grdHeader(imTokenIdx).RowHeight(1)
   
    lacTotalSpots.FontSize = fmTotalFontSize
    lacTotalPrice.FontSize = fmTotalFontSize
    lacTotalSpots.Height = grdLines(imTokenIdx).RowHeight(1) + 60   'Add 60 for space beneath value
    lacTotalPrice.Height = grdLines(imTokenIdx).RowHeight(1) + 60
    lacTotalSpots.ForeColor = BLUE
    lacTotalPrice.ForeColor = BLUE
    
    lacTotalCPP.FontSize = fmTotalFontSize
    lacTotalCPP.Height = grdLines(imTokenIdx).RowHeight(1) + 60   'Add 60 for space beneath value
    lacTotalCPP.ForeColor = BLUE
    
    lacTotalGrpPct.FontSize = fmTotalFontSize
    lacTotalGrpPct.Height = grdLines(imTokenIdx).RowHeight(1) + 60   'Add 60 for space beneath value
    lacTotalGrpPct.ForeColor = BLUE
    
    lacTotalGRP.FontSize = fmTotalFontSize
    lacTotalGRP.Height = grdLines(imTokenIdx).RowHeight(1) + 60   'Add 60 for space beneath value
    lacTotalGRP.ForeColor = BLUE
    
    lacBanner(1).FontSize = fmFontSize
    lacButtonHelp.FontSize = fmFontSize
    
    tscProgrammaticBuy.Font.Size = fmFontSize
End Sub

Private Sub mDisplayHelp()

End Sub

Private Sub mShowHelp()
    If bmHelpOff Then
        lacBanner(1).Visible = False
        lacButtonHelp.Visible = False
        Exit Sub
    End If
    lacBanner(1).Caption = ""
    If imWhichGrid = 0 Then
        Select Case lmHdEnableCol
            Case ADVERTISERINDEX
                lacBanner(1).Caption = "Select one of your advertisers, or enter the name of a new one."
            Case PRODUCTINDEX
                lacBanner(1).Caption = "The product name shown is the last one used.  Type in a new one if desired."
            Case STARTDATEINDEX
                lacBanner(1).Caption = "Type in a start date or select one from the Broadcast Calendar."
            Case NUMBERWEEKSINDEX
                lacBanner(1).Caption = "Enter the number of weeks, and the end date will be calculated.  Or, skip the number of weeks if you wish to enter the end date."
            Case ENDDATEINDEX
                lacBanner(1).Caption = "Type in an end date or select one from the Broadcast calendar.  Once you type in an end date the number of weeks will be calculated."
            Case ESTIMATENUMBERINDEX
                lacBanner(1).Caption = "Enter your estimate number or skip."
            Case DEMOINDEX
                lacBanner(1).Caption = "Select demographic."
            Case TARGETINDEX
                If imTabIndex = 0 Then
                    lacBanner(1).Caption = "Enter your Target Gross Rating Points."
                Else
                    lacBanner(1).Caption = "Enter your Target Gross Impressions."
                End If
        End Select
    Else
        Select Case lmLnEnableCol
            Case LENGTHINDEX
                If imTabIndex = 0 Then
                    lacBanner(1).Caption = "Select the desired spot length to see a menu of CPPs of all Networks and all their Dayparts, plus the avails for the next 6 weeks.  Plus sign means more than 9 avails per week."
                ElseIf imTabIndex = 1 Then
                    lacBanner(1).Caption = "Select the desired spot length to see a menu of CPMs of all Networks and all their Dayparts, plus the avails for the next 6 weeks.  Plus sign means more than 9 avails per week."
                Else
                    lacBanner(1).Caption = "Select the desired spot length to see a menu of prices of all Networks and all their Dayparts, followed by the avails for the next 6 weeks.  Plus sign means more than 9 avails per week."
                End If
            Case NUMBERSPOTSINDEX
                lacBanner(1).Caption = "Enter the number of spots desired. Do not exceed the weekly Avails."
            Case SPOTSPERINDEX
                lacBanner(1).Caption = "Specify whether the number of spots you just entered is per day, per week, or total for all weeks."
            Case GRPGRIMPPCTINDEX
                lacBanner(1).Caption = "Enter the percentage of the Target audience that you wish to air on the selected Network/Daypart."
        End Select
    End If
    If lacBanner(1).Caption <> "" Then
        lacBanner(1).Visible = True
    Else
        lacBanner(1).Visible = False
    End If
    lacButtonHelp.Visible = True
End Sub

Private Sub mLoadLogo()
    Dim ilRet As Integer
    On Error GoTo mRetryPicture:
    ilRet = 0
'    pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.jpg")
'    If ilRet <> 0 Then
'        ilRet = 0
'        pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.gif")
'        If ilRet <> 0 Then
'            ilRet = 0
'            pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.Bmp")
'        End If
'    End If
'    If ilRet = 0 Then
'        lacBanner(0).Visible = False
'    Else
        pbcLogo.Visible = False
'    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
End Sub

Private Sub mSetCommands()
    If grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, ADVERTISERINDEX) <> "" Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    If (grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, ADVERTISERINDEX) <> "") And (grdLines(imTokenIdx).TextMatrix(grdLines(imTokenIdx).FixedRows, LENGTHINDEX) <> "") Then
        cmcSubmit.Enabled = True
    Else
        cmcSubmit.Enabled = False
    End If
End Sub


'*** Start UDT ***
Private Sub mCreateUDTForPvf()

    tmPvf.lCode = pvf_rst!pvfCode
    tmPvf.sName = pvf_rst!pvfName
    tmPvf.lLkPvfCode = pvf_rst!pvfLkPvfCode
    tmPvf.iVefCode(0) = pvf_rst!pvfVefCode1
    tmPvf.iVefCode(1) = pvf_rst!pvfVefCode2
    tmPvf.iVefCode(2) = pvf_rst!pvfVefCode3
    tmPvf.iVefCode(3) = pvf_rst!pvfVefCode4
    tmPvf.iVefCode(4) = pvf_rst!pvfVefCode5
    tmPvf.iVefCode(5) = pvf_rst!pvfVefCode6
    tmPvf.iVefCode(6) = pvf_rst!pvfVefCode7
    tmPvf.iVefCode(7) = pvf_rst!pvfVefCode8
    tmPvf.iVefCode(8) = pvf_rst!pvfVefCode9
    tmPvf.iVefCode(9) = pvf_rst!pvfVefCode10
    tmPvf.iVefCode(10) = pvf_rst!pvfVefCode11
    tmPvf.iVefCode(11) = pvf_rst!pvfVefCode12
    tmPvf.iVefCode(12) = pvf_rst!pvfVefCode13
    tmPvf.iVefCode(13) = pvf_rst!pvfVefCode14
    tmPvf.iVefCode(14) = pvf_rst!pvfVefCode15
    tmPvf.iVefCode(15) = pvf_rst!pvfVefCode16
    tmPvf.iVefCode(16) = pvf_rst!pvfVefCode17
    tmPvf.iVefCode(17) = pvf_rst!pvfVefCode18
    tmPvf.iVefCode(18) = pvf_rst!pvfVefCode19
    tmPvf.iVefCode(19) = pvf_rst!pvfVefCode20
    tmPvf.iVefCode(20) = pvf_rst!pvfVefCode21
    tmPvf.iVefCode(21) = pvf_rst!pvfVefCode22
    tmPvf.iVefCode(22) = pvf_rst!pvfVefCode23
    tmPvf.iVefCode(23) = pvf_rst!pvfVefCode24
    tmPvf.iVefCode(24) = pvf_rst!pvfVefCode25
    tmPvf.iRdfCode(0) = pvf_rst!pvfRdfCode1
    tmPvf.iRdfCode(1) = pvf_rst!pvfRdfCode2
    tmPvf.iRdfCode(2) = pvf_rst!pvfRdfCode3
    tmPvf.iRdfCode(3) = pvf_rst!pvfRdfCode4
    tmPvf.iRdfCode(4) = pvf_rst!pvfRdfCode5
    tmPvf.iRdfCode(5) = pvf_rst!pvfRdfCode6
    tmPvf.iRdfCode(6) = pvf_rst!pvfRdfCode7
    tmPvf.iRdfCode(7) = pvf_rst!pvfRdfCode8
    tmPvf.iRdfCode(8) = pvf_rst!pvfRdfCode9
    tmPvf.iRdfCode(9) = pvf_rst!pvfRdfCode10
    tmPvf.iRdfCode(10) = pvf_rst!pvfRdfCode11
    tmPvf.iRdfCode(11) = pvf_rst!pvfRdfCode12
    tmPvf.iRdfCode(12) = pvf_rst!pvfRdfCode13
    tmPvf.iRdfCode(13) = pvf_rst!pvfRdfCode14
    tmPvf.iRdfCode(14) = pvf_rst!pvfRdfCode15
    tmPvf.iRdfCode(15) = pvf_rst!pvfRdfCode16
    tmPvf.iRdfCode(16) = pvf_rst!pvfRdfCode17
    tmPvf.iRdfCode(17) = pvf_rst!pvfRdfCode18
    tmPvf.iRdfCode(18) = pvf_rst!pvfRdfCode19
    tmPvf.iRdfCode(19) = pvf_rst!pvfRdfCode20
    tmPvf.iRdfCode(20) = pvf_rst!pvfRdfCode21
    tmPvf.iRdfCode(21) = pvf_rst!pvfRdfCode22
    tmPvf.iRdfCode(22) = pvf_rst!pvfRdfCode23
    tmPvf.iRdfCode(23) = pvf_rst!pvfRdfCode24
    tmPvf.iRdfCode(24) = pvf_rst!pvfRdfCode25
    tmPvf.iNoSpot(0) = pvf_rst!pvfNoSpot1
    tmPvf.iNoSpot(1) = pvf_rst!pvfNoSpot2
    tmPvf.iNoSpot(2) = pvf_rst!pvfNoSpot3
    tmPvf.iNoSpot(3) = pvf_rst!pvfNoSpot4
    tmPvf.iNoSpot(4) = pvf_rst!pvfNoSpot5
    tmPvf.iNoSpot(5) = pvf_rst!pvfNoSpot6
    tmPvf.iNoSpot(6) = pvf_rst!pvfNoSpot7
    tmPvf.iNoSpot(7) = pvf_rst!pvfNoSpot8
    tmPvf.iNoSpot(8) = pvf_rst!pvfNoSpot9
    tmPvf.iNoSpot(9) = pvf_rst!pvfNoSpot10
    tmPvf.iNoSpot(10) = pvf_rst!pvfNoSpot11
    tmPvf.iNoSpot(11) = pvf_rst!pvfNoSpot12
    tmPvf.iNoSpot(12) = pvf_rst!pvfNoSpot13
    tmPvf.iNoSpot(13) = pvf_rst!pvfNoSpot14
    tmPvf.iNoSpot(14) = pvf_rst!pvfNoSpot15
    tmPvf.iNoSpot(15) = pvf_rst!pvfNoSpot16
    tmPvf.iNoSpot(16) = pvf_rst!pvfNoSpot17
    tmPvf.iNoSpot(17) = pvf_rst!pvfNoSpot18
    tmPvf.iNoSpot(18) = pvf_rst!pvfNoSpot19
    tmPvf.iNoSpot(19) = pvf_rst!pvfNoSpot20
    tmPvf.iNoSpot(20) = pvf_rst!pvfNoSpot21
    tmPvf.iNoSpot(21) = pvf_rst!pvfNoSpot22
    tmPvf.iNoSpot(22) = pvf_rst!pvfNoSpot23
    tmPvf.iNoSpot(23) = pvf_rst!pvfNoSpot24
    tmPvf.iNoSpot(24) = pvf_rst!pvfNoSpot25
    tmPvf.iPctRate(0) = pvf_rst!pvfPctRate1
    tmPvf.iPctRate(1) = pvf_rst!pvfPctRate2
    tmPvf.iPctRate(2) = pvf_rst!pvfPctRate3
    tmPvf.iPctRate(3) = pvf_rst!pvfPctRate4
    tmPvf.iPctRate(4) = pvf_rst!pvfPctRate5
    tmPvf.iPctRate(5) = pvf_rst!pvfPctRate6
    tmPvf.iPctRate(6) = pvf_rst!pvfPctRate7
    tmPvf.iPctRate(7) = pvf_rst!pvfPctRate8
    tmPvf.iPctRate(8) = pvf_rst!pvfPctRate9
    tmPvf.iPctRate(9) = pvf_rst!pvfPctRate10
    tmPvf.iPctRate(10) = pvf_rst!pvfPctRate11
    tmPvf.iPctRate(11) = pvf_rst!pvfPctRate12
    tmPvf.iPctRate(12) = pvf_rst!pvfPctRate13
    tmPvf.iPctRate(13) = pvf_rst!pvfPctRate14
    tmPvf.iPctRate(14) = pvf_rst!pvfPctRate15
    tmPvf.iPctRate(15) = pvf_rst!pvfPctRate16
    tmPvf.iPctRate(16) = pvf_rst!pvfPctRate17
    tmPvf.iPctRate(17) = pvf_rst!pvfPctRate18
    tmPvf.iPctRate(18) = pvf_rst!pvfPctRate19
    tmPvf.iPctRate(19) = pvf_rst!pvfPctRate20
    tmPvf.iPctRate(20) = pvf_rst!pvfPctRate21
    tmPvf.iPctRate(21) = pvf_rst!pvfPctRate22
    tmPvf.iPctRate(22) = pvf_rst!pvfPctRate23
    tmPvf.iPctRate(23) = pvf_rst!pvfPctRate24
    tmPvf.iPctRate(24) = pvf_rst!pvfPctRate25
End Sub

Private Sub mCreateUDTForRcf()

    tmRcf.iCode = rcf_rst!rcfCode
    tmRcf.sName = rcf_rst!rcfName
    tmRcf.iVefCode = rcf_rst!rcfvefCode
    gPackDate Format$(rcf_rst!rcfStartDate, sgShowDateForm), tmRcf.iStartDate(0), tmRcf.iStartDate(1)
    gPackDate Format$(rcf_rst!rcfEndDate, sgShowDateForm), tmRcf.iEndDate(0), tmRcf.iEndDate(1)
    tmRcf.iGridsUsed = rcf_rst!rcfGridsUsed
    tmRcf.iTodayGrid = rcf_rst!rcfTodayGrid
    tmRcf.iBaseLen = rcf_rst!rcfBaseLen
    tmRcf.lRound = rcf_rst!rcfRound
    tmRcf.sUseBB = rcf_rst!rcfUseBB
    tmRcf.lBBValue = rcf_rst!rcfBBValue
    tmRcf.sUseFP = rcf_rst!rcfUseFP
    tmRcf.lFPValue = rcf_rst!rcfFPValue
    tmRcf.sUseNP = rcf_rst!rcfUseNP
    tmRcf.lNPValue = rcf_rst!rcfNPValue
    tmRcf.sUseLen = rcf_rst!rcfUseLen
    tmRcf.iLen(0) = rcf_rst!rcfLen1
    tmRcf.iLen(1) = rcf_rst!rcfLen2
    tmRcf.iLen(2) = rcf_rst!rcfLen3
    tmRcf.iLen(3) = rcf_rst!rcfLen4
    tmRcf.iLen(4) = rcf_rst!rcfLen5
    tmRcf.iLen(5) = rcf_rst!rcfLen6
    tmRcf.iLen(6) = rcf_rst!rcfLen7
    tmRcf.iLen(7) = rcf_rst!rcfLen8
    tmRcf.iLen(8) = rcf_rst!rcfLen9
    tmRcf.iLen(9) = rcf_rst!rcfLen10
    tmRcf.lValue(0) = rcf_rst!rcfValue1
    tmRcf.lValue(1) = rcf_rst!rcfValue2
    tmRcf.lValue(2) = rcf_rst!rcfValue3
    tmRcf.lValue(3) = rcf_rst!rcfValue4
    tmRcf.lValue(4) = rcf_rst!rcfValue5
    tmRcf.lValue(5) = rcf_rst!rcfValue6
    tmRcf.lValue(6) = rcf_rst!rcfValue7
    tmRcf.lValue(7) = rcf_rst!rcfValue8
    tmRcf.lValue(8) = rcf_rst!rcfValue9
    tmRcf.lValue(9) = rcf_rst!rcfValue10
    tmRcf.sUseSpot = rcf_rst!rcfUseSpot
    tmRcf.iSpotMin(0) = rcf_rst!rcfSpotMin1
    tmRcf.iSpotMin(1) = rcf_rst!rcfSpotMin2
    tmRcf.iSpotMin(2) = rcf_rst!rcfSpotMin3
    tmRcf.iSpotMin(3) = rcf_rst!rcfSpotMin4
    tmRcf.iSpotMax(0) = rcf_rst!rcfSpotMax1
    tmRcf.iSpotMax(1) = rcf_rst!rcfSpotMax2
    tmRcf.iSpotMax(2) = rcf_rst!rcfSpotMax3
    tmRcf.iSpotMax(3) = rcf_rst!rcfSpotMax4
    tmRcf.lSpotVal(0) = rcf_rst!rcfSpotVal1
    tmRcf.lSpotVal(1) = rcf_rst!rcfSpotVal2
    tmRcf.lSpotVal(2) = rcf_rst!rcfSpotVal3
    tmRcf.lSpotVal(3) = rcf_rst!rcfSpotVal4
    tmRcf.sUseWeek = rcf_rst!rcfUseWeek
    tmRcf.iWkMin(0) = rcf_rst!rcfWkMin1
    tmRcf.iWkMin(1) = rcf_rst!rcfWkMin2
    tmRcf.iWkMin(2) = rcf_rst!rcfWkMin3
    tmRcf.iWkMin(3) = rcf_rst!rcfWkMin4
    tmRcf.iWkMax(0) = rcf_rst!rcfWkMax1
    tmRcf.iWkMax(1) = rcf_rst!rcfWkMax2
    tmRcf.iWkMax(2) = rcf_rst!rcfWkMax3
    tmRcf.iWkMax(3) = rcf_rst!rcfWkMax4
    tmRcf.lWkVal(0) = rcf_rst!rcfWkVal1
    tmRcf.lWkVal(1) = rcf_rst!rcfWkVal2
    tmRcf.lWkVal(2) = rcf_rst!rcfWkVal3
    tmRcf.lWkVal(3) = rcf_rst!rcfWkVal4
    tmRcf.sUseDay = rcf_rst!rcfUseDay
    tmRcf.sDay(0, 1) = rcf_rst!rcfDyMo1
    tmRcf.sDay(1, 1) = rcf_rst!rcfDyMo2
    tmRcf.sDay(2, 1) = rcf_rst!rcfDyMo3
    tmRcf.sDay(3, 1) = rcf_rst!rcfDyMo4
    tmRcf.sDay(4, 1) = rcf_rst!rcfDyMo5
    tmRcf.sDay(5, 1) = rcf_rst!rcfDyMo6
    tmRcf.sDay(6, 1) = rcf_rst!rcfDyMo7
    tmRcf.sDay(7, 1) = rcf_rst!rcfDyMo8
    tmRcf.sDay(8, 1) = rcf_rst!rcfDyMo9
    tmRcf.sDay(9, 1) = rcf_rst!rcfDyMo10
    tmRcf.sDay(0, 2) = rcf_rst!rcfDyTu1
    tmRcf.sDay(1, 2) = rcf_rst!rcfDyTu2
    tmRcf.sDay(2, 2) = rcf_rst!rcfDyTu3
    tmRcf.sDay(3, 2) = rcf_rst!rcfDyTu4
    tmRcf.sDay(4, 2) = rcf_rst!rcfDyTu5
    tmRcf.sDay(5, 2) = rcf_rst!rcfDyTu6
    tmRcf.sDay(6, 2) = rcf_rst!rcfDyTu7
    tmRcf.sDay(7, 2) = rcf_rst!rcfDyTu8
    tmRcf.sDay(8, 2) = rcf_rst!rcfDyTu9
    tmRcf.sDay(9, 2) = rcf_rst!rcfDyTu10
    tmRcf.sDay(0, 3) = rcf_rst!rcfDyWe1
    tmRcf.sDay(1, 3) = rcf_rst!rcfDyWe2
    tmRcf.sDay(2, 3) = rcf_rst!rcfDyWe3
    tmRcf.sDay(3, 3) = rcf_rst!rcfDyWe4
    tmRcf.sDay(4, 3) = rcf_rst!rcfDyWe5
    tmRcf.sDay(5, 3) = rcf_rst!rcfDyWe6
    tmRcf.sDay(6, 3) = rcf_rst!rcfDyWe7
    tmRcf.sDay(7, 3) = rcf_rst!rcfDyWe8
    tmRcf.sDay(8, 3) = rcf_rst!rcfDyWe9
    tmRcf.sDay(9, 3) = rcf_rst!rcfDyWe10
    tmRcf.sDay(0, 4) = rcf_rst!rcfDyTh1
    tmRcf.sDay(1, 4) = rcf_rst!rcfDyTh2
    tmRcf.sDay(2, 4) = rcf_rst!rcfDyTh3
    tmRcf.sDay(3, 4) = rcf_rst!rcfDyTh4
    tmRcf.sDay(4, 4) = rcf_rst!rcfDyTh5
    tmRcf.sDay(5, 4) = rcf_rst!rcfDyTh6
    tmRcf.sDay(6, 4) = rcf_rst!rcfDyTh7
    tmRcf.sDay(7, 4) = rcf_rst!rcfDyTh8
    tmRcf.sDay(8, 4) = rcf_rst!rcfDyTh9
    tmRcf.sDay(9, 4) = rcf_rst!rcfDyTh10
    tmRcf.sDay(0, 5) = rcf_rst!rcfDyFr1
    tmRcf.sDay(1, 5) = rcf_rst!rcfDyFr2
    tmRcf.sDay(2, 5) = rcf_rst!rcfDyFr3
    tmRcf.sDay(3, 5) = rcf_rst!rcfDyFr4
    tmRcf.sDay(4, 5) = rcf_rst!rcfDyFr5
    tmRcf.sDay(5, 5) = rcf_rst!rcfDyFr6
    tmRcf.sDay(6, 5) = rcf_rst!rcfDyFr7
    tmRcf.sDay(7, 5) = rcf_rst!rcfDyFr8
    tmRcf.sDay(8, 5) = rcf_rst!rcfDyFr9
    tmRcf.sDay(9, 5) = rcf_rst!rcfDyFr10
    tmRcf.sDay(0, 6) = rcf_rst!rcfDySa1
    tmRcf.sDay(1, 6) = rcf_rst!rcfDySa2
    tmRcf.sDay(2, 6) = rcf_rst!rcfDySa3
    tmRcf.sDay(3, 6) = rcf_rst!rcfDySa4
    tmRcf.sDay(4, 6) = rcf_rst!rcfDySa5
    tmRcf.sDay(5, 6) = rcf_rst!rcfDySa6
    tmRcf.sDay(6, 6) = rcf_rst!rcfDySa7
    tmRcf.sDay(7, 6) = rcf_rst!rcfDySa8
    tmRcf.sDay(8, 6) = rcf_rst!rcfDySa9
    tmRcf.sDay(9, 6) = rcf_rst!rcfDySa10
    tmRcf.sDay(0, 6) = rcf_rst!rcfDySu1
    tmRcf.sDay(1, 6) = rcf_rst!rcfDySu2
    tmRcf.sDay(2, 6) = rcf_rst!rcfDySu3
    tmRcf.sDay(3, 6) = rcf_rst!rcfDySu4
    tmRcf.sDay(4, 6) = rcf_rst!rcfDySu5
    tmRcf.sDay(5, 6) = rcf_rst!rcfDySu6
    tmRcf.sDay(6, 6) = rcf_rst!rcfDySu7
    tmRcf.sDay(7, 6) = rcf_rst!rcfDySu8
    tmRcf.sDay(8, 6) = rcf_rst!rcfDySu9
    tmRcf.sDay(9, 6) = rcf_rst!rcfDySu10
    tmRcf.lDyRate(0) = rcf_rst!rcfDayRate1
    tmRcf.lDyRate(1) = rcf_rst!rcfDayRate2
    tmRcf.lDyRate(2) = rcf_rst!rcfDayRate3
    tmRcf.lDyRate(3) = rcf_rst!rcfDayRate4
    tmRcf.lDyRate(4) = rcf_rst!rcfDayRate5
    tmRcf.lDyRate(5) = rcf_rst!rcfDayRate6
    tmRcf.lDyRate(6) = rcf_rst!rcfDayRate7
    tmRcf.lDyRate(7) = rcf_rst!rcfDayRate8
    tmRcf.lDyRate(8) = rcf_rst!rcfDayRate9
    tmRcf.lDyRate(9) = rcf_rst!rcfDayRate10
    tmRcf.sUseHour = rcf_rst!rcfUseHour
    
    tmRcf.sHrDay(0, 0) = rcf_rst!rcfHrMo1
    tmRcf.sHrDay(1, 0) = rcf_rst!rcfHrMo2
    tmRcf.sHrDay(2, 0) = rcf_rst!rcfHrMo3
    tmRcf.sHrDay(3, 0) = rcf_rst!rcfHrMo4
    tmRcf.sHrDay(4, 0) = rcf_rst!rcfHrMo5
    tmRcf.sHrDay(5, 0) = rcf_rst!rcfHrMo6
    tmRcf.sHrDay(6, 0) = rcf_rst!rcfHrMo7
    tmRcf.sHrDay(7, 0) = rcf_rst!rcfHrMo8
    tmRcf.sHrDay(8, 0) = rcf_rst!rcfHrMo9
    tmRcf.sHrDay(9, 0) = rcf_rst!rcfHrMo10
    
    tmRcf.sHrDay(1, 1) = rcf_rst!rcfHrTu1
    tmRcf.sHrDay(2, 1) = rcf_rst!rcfHrTu2
    tmRcf.sHrDay(3, 1) = rcf_rst!rcfHrTu3
    tmRcf.sHrDay(4, 1) = rcf_rst!rcfHrTu4
    tmRcf.sHrDay(5, 1) = rcf_rst!rcfHrTu5
    tmRcf.sHrDay(6, 1) = rcf_rst!rcfHrTu6
    tmRcf.sHrDay(7, 1) = rcf_rst!rcfHrTu7
    tmRcf.sHrDay(8, 1) = rcf_rst!rcfHrTu8
    tmRcf.sHrDay(9, 1) = rcf_rst!rcfHrTu9
    
    tmRcf.sHrDay(0, 2) = rcf_rst!rcfHrWe1
    tmRcf.sHrDay(1, 2) = rcf_rst!rcfHrWe2
    tmRcf.sHrDay(2, 2) = rcf_rst!rcfHrWe3
    tmRcf.sHrDay(3, 2) = rcf_rst!rcfHrWe4
    tmRcf.sHrDay(4, 2) = rcf_rst!rcfHrWe5
    tmRcf.sHrDay(5, 2) = rcf_rst!rcfHrWe6
    tmRcf.sHrDay(6, 2) = rcf_rst!rcfHrWe7
    tmRcf.sHrDay(7, 2) = rcf_rst!rcfHrWe8
    tmRcf.sHrDay(8, 2) = rcf_rst!rcfHrWe9
    tmRcf.sHrDay(9, 2) = rcf_rst!rcfHrWe10
    
    tmRcf.sHrDay(0, 3) = rcf_rst!rcfHrTh1
    tmRcf.sHrDay(1, 3) = rcf_rst!rcfHrTh2
    tmRcf.sHrDay(2, 3) = rcf_rst!rcfHrTh3
    tmRcf.sHrDay(3, 3) = rcf_rst!rcfHrTh4
    tmRcf.sHrDay(4, 3) = rcf_rst!rcfHrTh5
    tmRcf.sHrDay(5, 3) = rcf_rst!rcfHrTh6
    tmRcf.sHrDay(6, 3) = rcf_rst!rcfHrTh7
    tmRcf.sHrDay(7, 3) = rcf_rst!rcfHrTh8
    tmRcf.sHrDay(8, 3) = rcf_rst!rcfHrTh9
    tmRcf.sHrDay(9, 3) = rcf_rst!rcfHrTh10
    
    
    tmRcf.sHrDay(0, 4) = rcf_rst!rcfHrFr1
    tmRcf.sHrDay(1, 4) = rcf_rst!rcfHrFr2
    tmRcf.sHrDay(2, 4) = rcf_rst!rcfHrFr3
    tmRcf.sHrDay(3, 4) = rcf_rst!rcfHrFr4
    tmRcf.sHrDay(4, 4) = rcf_rst!rcfHrFr5
    tmRcf.sHrDay(5, 4) = rcf_rst!rcfHrFr6
    tmRcf.sHrDay(6, 4) = rcf_rst!rcfHrFr7
    tmRcf.sHrDay(7, 4) = rcf_rst!rcfHrFr8
    tmRcf.sHrDay(8, 4) = rcf_rst!rcfHrFr9
    tmRcf.sHrDay(9, 4) = rcf_rst!rcfHrFr10
    
    tmRcf.sHrDay(0, 5) = rcf_rst!rcfHrSa1
    tmRcf.sHrDay(1, 5) = rcf_rst!rcfHrSa2
    tmRcf.sHrDay(2, 5) = rcf_rst!rcfHrSa3
    tmRcf.sHrDay(3, 5) = rcf_rst!rcfHrSa4
    tmRcf.sHrDay(4, 5) = rcf_rst!rcfHrSa5
    tmRcf.sHrDay(5, 5) = rcf_rst!rcfHrSa6
    tmRcf.sHrDay(6, 5) = rcf_rst!rcfHrSa7
    tmRcf.sHrDay(7, 5) = rcf_rst!rcfHrSa8
    tmRcf.sHrDay(8, 5) = rcf_rst!rcfHrSa9
    tmRcf.sHrDay(9, 5) = rcf_rst!rcfHrSa10
    
    
    tmRcf.sHrDay(0, 6) = rcf_rst!rcfHrSu1
    tmRcf.sHrDay(1, 6) = rcf_rst!rcfHrSu2
    tmRcf.sHrDay(2, 6) = rcf_rst!rcfHrSu3
    tmRcf.sHrDay(3, 6) = rcf_rst!rcfHrSu4
    tmRcf.sHrDay(4, 6) = rcf_rst!rcfHrSu5
    tmRcf.sHrDay(5, 6) = rcf_rst!rcfHrSu6
    tmRcf.sHrDay(6, 6) = rcf_rst!rcfHrSu7
    tmRcf.sHrDay(7, 6) = rcf_rst!rcfHrSu8
    tmRcf.sHrDay(8, 6) = rcf_rst!rcfHrSu9
    tmRcf.sHrDay(9, 6) = rcf_rst!rcfHrSu10
    
    
    gPackTime Format$(rcf_rst!rcfHStartTm1, "ttttt"), tmRcf.iHrStartTime(0, 0), tmRcf.iHrStartTime(0, 0)
    gPackTime Format$(rcf_rst!rcfHStartTm2, "ttttt"), tmRcf.iHrStartTime(0, 1), tmRcf.iHrStartTime(0, 1)
    gPackTime Format$(rcf_rst!rcfHStartTm3, "ttttt"), tmRcf.iHrStartTime(0, 2), tmRcf.iHrStartTime(0, 2)
    gPackTime Format$(rcf_rst!rcfHStartTm4, "ttttt"), tmRcf.iHrStartTime(0, 3), tmRcf.iHrStartTime(0, 3)
    gPackTime Format$(rcf_rst!rcfHStartTm5, "ttttt"), tmRcf.iHrStartTime(0, 4), tmRcf.iHrStartTime(0, 4)
    gPackTime Format$(rcf_rst!rcfHStartTm6, "ttttt"), tmRcf.iHrStartTime(0, 5), tmRcf.iHrStartTime(0, 5)
    gPackTime Format$(rcf_rst!rcfHStartTm7, "ttttt"), tmRcf.iHrStartTime(0, 6), tmRcf.iHrStartTime(0, 6)
    gPackTime Format$(rcf_rst!rcfHStartTm8, "ttttt"), tmRcf.iHrStartTime(0, 7), tmRcf.iHrStartTime(0, 7)
    gPackTime Format$(rcf_rst!rcfHStartTm9, "ttttt"), tmRcf.iHrStartTime(0, 8), tmRcf.iHrStartTime(0, 8)
    gPackTime Format$(rcf_rst!rcfHStartTm10, "ttttt"), tmRcf.iHrStartTime(0, 9), tmRcf.iHrStartTime(0, 9)
    
    
    gPackTime Format$(rcf_rst!rcfHEndTm1, "ttttt"), tmRcf.iHrEndTime(0, 0), tmRcf.iHrEndTime(0, 0)
    gPackTime Format$(rcf_rst!rcfHEndTm2, "ttttt"), tmRcf.iHrEndTime(0, 1), tmRcf.iHrEndTime(0, 1)
    gPackTime Format$(rcf_rst!rcfHEndTm3, "ttttt"), tmRcf.iHrEndTime(0, 2), tmRcf.iHrEndTime(0, 2)
    gPackTime Format$(rcf_rst!rcfHEndTm4, "ttttt"), tmRcf.iHrEndTime(0, 3), tmRcf.iHrEndTime(0, 3)
    gPackTime Format$(rcf_rst!rcfHEndTm5, "ttttt"), tmRcf.iHrEndTime(0, 4), tmRcf.iHrEndTime(0, 4)
    gPackTime Format$(rcf_rst!rcfHEndTm6, "ttttt"), tmRcf.iHrEndTime(0, 5), tmRcf.iHrEndTime(0, 5)
    gPackTime Format$(rcf_rst!rcfHEndTm7, "ttttt"), tmRcf.iHrEndTime(0, 6), tmRcf.iHrEndTime(0, 6)
    gPackTime Format$(rcf_rst!rcfHEndTm8, "ttttt"), tmRcf.iHrEndTime(0, 7), tmRcf.iHrEndTime(0, 7)
    gPackTime Format$(rcf_rst!rcfHEndTm9, "ttttt"), tmRcf.iHrEndTime(0, 8), tmRcf.iHrEndTime(0, 8)
    gPackTime Format$(rcf_rst!rcfHEndTm10, "ttttt"), tmRcf.iHrEndTime(0, 9), tmRcf.iHrEndTime(0, 9)
    
    
    tmRcf.lHrRate(0) = rcf_rst!rcfHrRate1
    tmRcf.lHrRate(1) = rcf_rst!rcfHrRate2
    tmRcf.lHrRate(2) = rcf_rst!rcfHrRate3
    tmRcf.lHrRate(3) = rcf_rst!rcfHrRate4
    tmRcf.lHrRate(4) = rcf_rst!rcfHrRate5
    tmRcf.lHrRate(5) = rcf_rst!rcfHrRate6
    tmRcf.lHrRate(6) = rcf_rst!rcfHrRate7
    tmRcf.lHrRate(7) = rcf_rst!rcfHrRate8
    tmRcf.lHrRate(8) = rcf_rst!rcfHrRate9
    tmRcf.lHrRate(9) = rcf_rst!rcfHrRate10
    
    tmRcf.iFltNo(0) = rcf_rst!rcfFltNo1
    tmRcf.iFltNo(1) = rcf_rst!rcfFltNo2
    tmRcf.iFltNo(2) = rcf_rst!rcfFltNo3
    tmRcf.iFltNo(3) = rcf_rst!rcfFltNo4
    tmRcf.iFltNo(4) = rcf_rst!rcfFltNo5
    tmRcf.iFltNo(5) = rcf_rst!rcfFltNo6
    tmRcf.iFltNo(6) = rcf_rst!rcfFltNo7
    tmRcf.iFltNo(7) = rcf_rst!rcfFltNo8
    tmRcf.iFltNo(8) = rcf_rst!rcfFltNo9
    tmRcf.iFltNo(9) = rcf_rst!rcfFltNo10
    tmRcf.iFltNo(10) = rcf_rst!rcfFltNo11
    tmRcf.iFltNo(11) = rcf_rst!rcfFltNo12
    tmRcf.iFltNo(12) = rcf_rst!rcfFltNo13
    tmRcf.iFltNo(13) = rcf_rst!rcfFltNo14
    tmRcf.iFltNo(14) = rcf_rst!rcfFltNo15
    tmRcf.iFltNo(15) = rcf_rst!rcfFltNo16
    tmRcf.iFltNo(16) = rcf_rst!rcfFltNo17
    tmRcf.iFltNo(17) = rcf_rst!rcfFltNo18
    tmRcf.iFltNo(18) = rcf_rst!rcfFltNo19
    tmRcf.iFltNo(19) = rcf_rst!rcfFltNo20
    tmRcf.iFltNo(20) = rcf_rst!rcfFltNo21
    tmRcf.iFltNo(21) = rcf_rst!rcfFltNo22
    tmRcf.iFltNo(23) = rcf_rst!rcfFltNo23
    tmRcf.iFltNo(23) = rcf_rst!rcfFltNo24
    tmRcf.iFltNo(24) = rcf_rst!rcfFltNo25
    tmRcf.iFltNo(25) = rcf_rst!rcfFltNo26
    tmRcf.iFltNo(26) = rcf_rst!rcfFltNo27
    tmRcf.iFltNo(27) = rcf_rst!rcfFltNo28
    tmRcf.iFltNo(28) = rcf_rst!rcfFltNo29
    tmRcf.iFltNo(29) = rcf_rst!rcfFltNo30
    tmRcf.iFltNo(30) = rcf_rst!rcfFltNo31
    tmRcf.iFltNo(31) = rcf_rst!rcfFltNo32
    tmRcf.iFltNo(32) = rcf_rst!rcfFltNo33
    tmRcf.iFltNo(33) = rcf_rst!rcfFltNo34
    tmRcf.iFltNo(34) = rcf_rst!rcfFltNo35
    tmRcf.iFltNo(35) = rcf_rst!rcfFltNo36
    tmRcf.iFltNo(36) = rcf_rst!rcfFltNo37
    tmRcf.iFltNo(37) = rcf_rst!rcfFltNo38
    tmRcf.iFltNo(38) = rcf_rst!rcfFltNo39
    tmRcf.iFltNo(39) = rcf_rst!rcfFltNo40
    tmRcf.iFltNo(40) = rcf_rst!rcfFltNo41
    tmRcf.iFltNo(41) = rcf_rst!rcfFltNo42
    tmRcf.iFltNo(42) = rcf_rst!rcfFltNo43
    tmRcf.iFltNo(43) = rcf_rst!rcfFltNo44
    tmRcf.iFltNo(44) = rcf_rst!rcfFltNo45
    tmRcf.iFltNo(45) = rcf_rst!rcfFltNo46
    tmRcf.iFltNo(46) = rcf_rst!rcfFltNo47
    tmRcf.iFltNo(47) = rcf_rst!rcfFltNo48
    tmRcf.iFltNo(48) = rcf_rst!rcfFltNo49
    tmRcf.iFltNo(49) = rcf_rst!rcfFltNo50
    tmRcf.iFltNo(50) = rcf_rst!rcfFltNo51
    tmRcf.iFltNo(51) = rcf_rst!rcfFltNo52
    tmRcf.iFltNo(52) = rcf_rst!rcfFltNo53
    
    tmRcf.lGridIndex(0) = rcf_rst!rcfGrid1
    tmRcf.lGridIndex(1) = rcf_rst!rcfGrid2
    tmRcf.lGridIndex(2) = rcf_rst!rcfGrid3
    tmRcf.lGridIndex(3) = rcf_rst!rcfGrid4
    tmRcf.lGridIndex(4) = rcf_rst!rcfGrid5
    tmRcf.lGridIndex(5) = rcf_rst!rcfGrid6
    tmRcf.lGridIndex(6) = rcf_rst!rcfGrid7
    tmRcf.lGridIndex(7) = rcf_rst!rcfGrid8
    tmRcf.lGridIndex(8) = rcf_rst!rcfGrid9
    tmRcf.lGridIndex(9) = rcf_rst!rcfGrid10
    tmRcf.lGridIndex(10) = rcf_rst!rcfGrid11
    tmRcf.lGridIndex(11) = rcf_rst!rcfGrid12
    tmRcf.iYear = rcf_rst!rcfYear
    tmRcf.iUrfCode = rcf_rst!rcfurfCode
    tmRcf.iRemoteID = rcf_rst!rcfRemoteID
    tmRcf.iAutoCode = rcf_rst!rcfAutoCode
    gPackDate Format$(rcf_rst!rcfSyncDate, sgShowDateForm), tmRcf.iSyncDate(0), tmRcf.iSyncDate(1)
    gPackTime Format$(rcf_rst!rcfSyncTime, "ttttt"), tmRcf.iSyncTime(0), tmRcf.iSyncTime(1)
    tmRcf.sUse1stPos = rcf_rst!rcfUse1stPos
    tmRcf.l1stPosValue = rcf_rst!rcf1stPosValue
    tmRcf.sUseSoloAvail = rcf_rst!rcfUseSoloAvail
    tmRcf.lSoloAvailValue = rcf_rst!rcfSoloAvailValue
    tmRcf.sUsePrefDT = rcf_rst!rcfUsePrefDT
    tmRcf.lPrefDTValue = rcf_rst!rcfPrefDTValue
    tmRcf.sUnused = rcf_rst!rcfUnused
End Sub

Private Sub mCreateUDTForRif()
    Dim ilUpper As Integer
    
    ilUpper = UBound(tmRif)
    tmRif(ilUpper).lCode = rif_rst!rifCode
    tmRif(ilUpper).iRcfCode = rif_rst!rifrcfCode
    tmRif(ilUpper).iVefCode = rif_rst!rifvefCode
    tmRif(ilUpper).iRdfCode = rif_rst!rifrdfCode
    tmRif(ilUpper).iYear = rif_rst!rifYear
    tmRif(ilUpper).sBase = rif_rst!rifBase
    tmRif(ilUpper).sRpt = rif_rst!rifRpt
    tmRif(ilUpper).iSort = rif_rst!rifSort
    tmRif(ilUpper).iROSPct = rif_rst!rifROSPct
    tmRif(ilUpper).lRate(0) = rif_rst!rifRate0
    tmRif(ilUpper).lRate(1) = rif_rst!rifRate1
    tmRif(ilUpper).lRate(2) = rif_rst!rifRate2
    tmRif(ilUpper).lRate(3) = rif_rst!rifRate3
    tmRif(ilUpper).lRate(4) = rif_rst!rifRate4
    tmRif(ilUpper).lRate(5) = rif_rst!rifRate5
    tmRif(ilUpper).lRate(6) = rif_rst!rifRate6
    tmRif(ilUpper).lRate(7) = rif_rst!rifRate7
    tmRif(ilUpper).lRate(8) = rif_rst!rifRate8
    tmRif(ilUpper).lRate(9) = rif_rst!rifRate9
    tmRif(ilUpper).lRate(10) = rif_rst!rifRate10
    tmRif(ilUpper).lRate(11) = rif_rst!rifRate11
    tmRif(ilUpper).lRate(12) = rif_rst!rifRate12
    tmRif(ilUpper).lRate(13) = rif_rst!rifRate13
    tmRif(ilUpper).lRate(14) = rif_rst!rifRate14
    tmRif(ilUpper).lRate(15) = rif_rst!rifRate15
    tmRif(ilUpper).lRate(16) = rif_rst!rifRate16
    tmRif(ilUpper).lRate(17) = rif_rst!rifRate17
    tmRif(ilUpper).lRate(18) = rif_rst!rifRate18
    tmRif(ilUpper).lRate(19) = rif_rst!rifRate19
    tmRif(ilUpper).lRate(20) = rif_rst!rifRate20
    tmRif(ilUpper).lRate(21) = rif_rst!rifRate21
    tmRif(ilUpper).lRate(22) = rif_rst!rifRate22
    tmRif(ilUpper).lRate(23) = rif_rst!rifRate23
    tmRif(ilUpper).lRate(24) = rif_rst!rifRate24
    tmRif(ilUpper).lRate(25) = rif_rst!rifRate25
    tmRif(ilUpper).lRate(26) = rif_rst!rifRate26
    tmRif(ilUpper).lRate(27) = rif_rst!rifRate27
    tmRif(ilUpper).lRate(28) = rif_rst!rifRate28
    tmRif(ilUpper).lRate(29) = rif_rst!rifRate29
    tmRif(ilUpper).lRate(30) = rif_rst!rifRate30
    tmRif(ilUpper).lRate(31) = rif_rst!rifRate31
    tmRif(ilUpper).lRate(32) = rif_rst!rifRate32
    tmRif(ilUpper).lRate(33) = rif_rst!rifRate33
    tmRif(ilUpper).lRate(34) = rif_rst!rifRate34
    tmRif(ilUpper).lRate(35) = rif_rst!rifRate35
    tmRif(ilUpper).lRate(36) = rif_rst!rifRate36
    tmRif(ilUpper).lRate(37) = rif_rst!rifRate37
    tmRif(ilUpper).lRate(38) = rif_rst!rifRate38
    tmRif(ilUpper).lRate(39) = rif_rst!rifRate39
    tmRif(ilUpper).lRate(40) = rif_rst!rifRate40
    tmRif(ilUpper).lRate(41) = rif_rst!rifRate41
    tmRif(ilUpper).lRate(42) = rif_rst!rifRate42
    tmRif(ilUpper).lRate(43) = rif_rst!rifRate43
    tmRif(ilUpper).lRate(44) = rif_rst!rifRate44
    tmRif(ilUpper).lRate(45) = rif_rst!rifRate45
    tmRif(ilUpper).lRate(46) = rif_rst!rifRate46
    tmRif(ilUpper).lRate(47) = rif_rst!rifRate47
    tmRif(ilUpper).lRate(48) = rif_rst!rifRate48
    tmRif(ilUpper).lRate(49) = rif_rst!rifRate49
    tmRif(ilUpper).lRate(50) = rif_rst!rifRate50
    tmRif(ilUpper).lRate(51) = rif_rst!rifRate51
    tmRif(ilUpper).lRate(52) = rif_rst!rifRate52
    tmRif(ilUpper).lRate(53) = rif_rst!rifRate53
    gPackDate Format$(rif_rst!rifRateChgDate, sgShowDateForm), tmRif(ilUpper).iRateChgDate(0), tmRif(ilUpper).iRateChgDate(1)
    gPackTime Format$(rif_rst!rifRateChgTime, "ttttt"), tmRif(ilUpper).iRateChgTime(0), tmRif(ilUpper).iRateChgTime(1)
    tmRif(ilUpper).iRemoteID = rif_rst!rifRemoteID
    tmRif(ilUpper).lAutoCode = rif_rst!rifAutoCode
    gPackDate Format$(rif_rst!rifSyncDate, sgShowDateForm), tmRif(ilUpper).iSyncDate(0), tmRif(ilUpper).iSyncDate(1)
    gPackTime Format$(rif_rst!rifSyncTime, "ttttt"), tmRif(ilUpper).iSyncTime(0), tmRif(ilUpper).iSyncTime(1)
    tmRif(ilUpper).sUnused = rif_rst!rifUnused
    ilUpper = ilUpper + 1
    ReDim Preserve tmRif(0 To ilUpper) As RIF
End Sub

Private Sub mCreateUDTForRdf()
    Dim ilUpper As Integer
    
    ilUpper = UBound(tmRdf)
    tmRdf(ilUpper).iCode = rdf_rst!rdfCode
    tmRdf(ilUpper).sName = rdf_rst!rdfName
    tmRdf(ilUpper).sBase = rdf_rst!rdfBase
    tmRdf(ilUpper).sReport = rdf_rst!rdfReport
    tmRdf(ilUpper).sInOut = rdf_rst!rdfInOut
    tmRdf(ilUpper).ianfCode = rdf_rst!rdfanfCode
    tmRdf(ilUpper).sTimeOver = rdf_rst!rdfTimeOver
    tmRdf(ilUpper).sState = rdf_rst!rdfState
    
    gPackTime Format$(rdf_rst!rdfStartTime1, "ttttt"), tmRdf(ilUpper).iStartTime(0, 0), tmRdf(ilUpper).iStartTime(0, 0)
    gPackTime Format$(rdf_rst!rdfStartTime2, "ttttt"), tmRdf(ilUpper).iStartTime(0, 1), tmRdf(ilUpper).iStartTime(0, 1)
    gPackTime Format$(rdf_rst!rdfStartTime3, "ttttt"), tmRdf(ilUpper).iStartTime(0, 2), tmRdf(ilUpper).iStartTime(0, 2)
    gPackTime Format$(rdf_rst!rdfStartTime4, "ttttt"), tmRdf(ilUpper).iStartTime(0, 3), tmRdf(ilUpper).iStartTime(0, 3)
    gPackTime Format$(rdf_rst!rdfStartTime5, "ttttt"), tmRdf(ilUpper).iStartTime(0, 4), tmRdf(ilUpper).iStartTime(0, 4)
    gPackTime Format$(rdf_rst!rdfStartTime6, "ttttt"), tmRdf(ilUpper).iStartTime(0, 5), tmRdf(ilUpper).iStartTime(0, 5)
    gPackTime Format$(rdf_rst!rdfStartTime7, "ttttt"), tmRdf(ilUpper).iStartTime(0, 6), tmRdf(ilUpper).iStartTime(0, 6)
    
    gPackTime Format$(rdf_rst!rdfEndTime1, "ttttt"), tmRdf(ilUpper).iEndTime(0, 0), tmRdf(ilUpper).iEndTime(0, 0)
    gPackTime Format$(rdf_rst!rdfEndTime2, "ttttt"), tmRdf(ilUpper).iEndTime(0, 1), tmRdf(ilUpper).iEndTime(0, 1)
    gPackTime Format$(rdf_rst!rdfEndTime3, "ttttt"), tmRdf(ilUpper).iEndTime(0, 2), tmRdf(ilUpper).iEndTime(0, 2)
    gPackTime Format$(rdf_rst!rdfEndTime4, "ttttt"), tmRdf(ilUpper).iEndTime(0, 3), tmRdf(ilUpper).iEndTime(0, 3)
    gPackTime Format$(rdf_rst!rdfEndTime5, "ttttt"), tmRdf(ilUpper).iEndTime(0, 4), tmRdf(ilUpper).iEndTime(0, 4)
    gPackTime Format$(rdf_rst!rdfEndTime6, "ttttt"), tmRdf(ilUpper).iEndTime(0, 5), tmRdf(ilUpper).iEndTime(0, 5)
    gPackTime Format$(rdf_rst!rdfEndTime7, "ttttt"), tmRdf(ilUpper).iEndTime(0, 6), tmRdf(ilUpper).iEndTime(0, 6)
    
    tmRdf(ilUpper).sWkDays(0, 0) = rdf_rst!rdfMo1
    tmRdf(ilUpper).sWkDays(1, 0) = rdf_rst!rdfMo2
    tmRdf(ilUpper).sWkDays(2, 0) = rdf_rst!rdfMo3
    tmRdf(ilUpper).sWkDays(3, 0) = rdf_rst!rdfMo4
    tmRdf(ilUpper).sWkDays(4, 0) = rdf_rst!rdfMo5
    tmRdf(ilUpper).sWkDays(5, 0) = rdf_rst!rdfMo6
    tmRdf(ilUpper).sWkDays(6, 0) = rdf_rst!rdfMo7
    
    tmRdf(ilUpper).sWkDays(0, 1) = rdf_rst!rdfTu1
    tmRdf(ilUpper).sWkDays(1, 1) = rdf_rst!rdfTu2
    tmRdf(ilUpper).sWkDays(2, 1) = rdf_rst!rdfTu3
    tmRdf(ilUpper).sWkDays(3, 1) = rdf_rst!rdfTu4
    tmRdf(ilUpper).sWkDays(4, 1) = rdf_rst!rdfTu5
    tmRdf(ilUpper).sWkDays(5, 1) = rdf_rst!rdfTu6
    tmRdf(ilUpper).sWkDays(6, 1) = rdf_rst!rdfTu7
    
    tmRdf(ilUpper).sWkDays(0, 2) = rdf_rst!rdfWe1
    tmRdf(ilUpper).sWkDays(1, 2) = rdf_rst!rdfWe2
    tmRdf(ilUpper).sWkDays(2, 2) = rdf_rst!rdfWe3
    tmRdf(ilUpper).sWkDays(3, 2) = rdf_rst!rdfWe4
    tmRdf(ilUpper).sWkDays(4, 2) = rdf_rst!rdfWe5
    tmRdf(ilUpper).sWkDays(5, 2) = rdf_rst!rdfWe6
    tmRdf(ilUpper).sWkDays(6, 2) = rdf_rst!rdfWe7
    
    tmRdf(ilUpper).sWkDays(0, 3) = rdf_rst!rdfTh1
    tmRdf(ilUpper).sWkDays(1, 3) = rdf_rst!rdfTh2
    tmRdf(ilUpper).sWkDays(2, 3) = rdf_rst!rdfTh3
    tmRdf(ilUpper).sWkDays(3, 3) = rdf_rst!rdfTh4
    tmRdf(ilUpper).sWkDays(4, 3) = rdf_rst!rdfTh5
    tmRdf(ilUpper).sWkDays(5, 3) = rdf_rst!rdfTh6
    tmRdf(ilUpper).sWkDays(6, 3) = rdf_rst!rdfTh7
    
    tmRdf(ilUpper).sWkDays(0, 4) = rdf_rst!rdfFr1
    tmRdf(ilUpper).sWkDays(1, 4) = rdf_rst!rdfFr2
    tmRdf(ilUpper).sWkDays(2, 4) = rdf_rst!rdfFr3
    tmRdf(ilUpper).sWkDays(3, 4) = rdf_rst!rdfFr4
    tmRdf(ilUpper).sWkDays(4, 4) = rdf_rst!rdfFr5
    tmRdf(ilUpper).sWkDays(5, 4) = rdf_rst!rdfFr6
    tmRdf(ilUpper).sWkDays(6, 4) = rdf_rst!rdfFr7
    
    tmRdf(ilUpper).sWkDays(0, 5) = rdf_rst!rdfSa1
    tmRdf(ilUpper).sWkDays(1, 5) = rdf_rst!rdfSa2
    tmRdf(ilUpper).sWkDays(2, 5) = rdf_rst!rdfSa3
    tmRdf(ilUpper).sWkDays(3, 5) = rdf_rst!rdfSa4
    tmRdf(ilUpper).sWkDays(4, 5) = rdf_rst!rdfSa5
    tmRdf(ilUpper).sWkDays(5, 5) = rdf_rst!rdfSa6
    tmRdf(ilUpper).sWkDays(6, 5) = rdf_rst!rdfSa7
    
    tmRdf(ilUpper).sWkDays(0, 6) = rdf_rst!rdfSu1
    tmRdf(ilUpper).sWkDays(1, 6) = rdf_rst!rdfSu2
    tmRdf(ilUpper).sWkDays(2, 6) = rdf_rst!rdfSu3
    tmRdf(ilUpper).sWkDays(3, 6) = rdf_rst!rdfSu4
    tmRdf(ilUpper).sWkDays(4, 6) = rdf_rst!rdfSu5
    tmRdf(ilUpper).sWkDays(5, 6) = rdf_rst!rdfSu6
    tmRdf(ilUpper).sWkDays(6, 6) = rdf_rst!rdfSu7
    
    tmRdf(ilUpper).iSpotPct(0) = rdf_rst!rdfSpotPct1
    tmRdf(ilUpper).iSpotPct(1) = rdf_rst!rdfSpotPct2
    tmRdf(ilUpper).iSpotPct(2) = rdf_rst!rdfSpotPct3
    tmRdf(ilUpper).iSpotPct(3) = rdf_rst!rdfSpotPct4
    tmRdf(ilUpper).iSpotPct(4) = rdf_rst!rdfSpotPct5
    tmRdf(ilUpper).iSpotPct(5) = rdf_rst!rdfSpotPct6
    tmRdf(ilUpper).iSpotPct(6) = rdf_rst!rdfSpotPct7
    
    tmRdf(ilUpper).sContrRef = rdf_rst!rdfContrRef
    
    tmRdf(ilUpper).iLtfCode(0) = rdf_rst!rdfltfCode1
    tmRdf(ilUpper).iLtfCode(1) = rdf_rst!rdfltfCode2
    tmRdf(ilUpper).iLtfCode(2) = rdf_rst!rdfltfCode3
    
    tmRdf(ilUpper).iSortCode = rdf_rst!rdfSortCode
    tmRdf(ilUpper).iRemoteID = rdf_rst!rdfRemoteID
    tmRdf(ilUpper).iAutoCode = rdf_rst!rdfAutoCode
    gPackDate Format$(rdf_rst!rdfSyncDate, sgShowDateForm), tmRdf(ilUpper).iSyncDate(0), tmRdf(ilUpper).iSyncDate(1)
    gPackTime Format$(rdf_rst!rdfSyncTime, "ttttt"), tmRdf(ilUpper).iSyncTime(0), tmRdf(ilUpper).iSyncTime(1)
    tmRdf(ilUpper).iSourceID = rdf_rst!rdfSourceID
    tmRdf(ilUpper).sOverridePrtRules = rdf_rst!rdfOverridePrtRules
    tmRdf(ilUpper).sUnused = rdf_rst!rdfUnused
    ilUpper = ilUpper + 1
    ReDim Preserve tmRdf(0 To ilUpper) As RDF
End Sub


'*** Start Re-usable Code ***

Public Sub mAdvtPop()
    Dim slSQLQuery As String
    ReDim tgAdfInfo(0 To 0) As ADFINFO
    Dim ilLoop As Integer
    
    cbcAdvt.Clear
    slSQLQuery = "select * from adf_Advertisers where adfAgfCode = " & tgTokenInfo(imTokenIdx).iAgfCode & " Order by adfName"
    Set adf_rst = cnn.Execute(slSQLQuery)
    Do While Not adf_rst.EOF
        If (adf_rst!adfState <> "D") And (Trim$(adf_rst!adfBilAgyDir) <> "D") Then
            'Once Buyer user sign-on is working, only include those advterisers that reference the Agency that the buyer is part of
            cbcAdvt.AddItem Trim$(adf_rst!adfName)
            cbcAdvt.SetItemData = adf_rst!adfCode
            tgAdfInfo(UBound(tgAdfInfo)).sName = Trim$(mGetDataNoQuotes(adf_rst!adfName))
            tgAdfInfo(UBound(tgAdfInfo)).sProd = Trim$(mGetDataNoQuotes(adf_rst!adfProd))
            tgAdfInfo(UBound(tgAdfInfo)).iCode = Trim$(adf_rst!adfCode)
            tgAdfInfo(UBound(tgAdfInfo)).sDemo = Trim$(adf_rst!adfDemo1)
            tgAdfInfo(UBound(tgAdfInfo)).iAdfAgfCode = Trim$(adf_rst!adfagfCode)
            ReDim Preserve tgAdfInfo(0 To UBound(tgAdfInfo) + 1) As ADFINFO
        End If
        adf_rst.MoveNext
    Loop
    
    smResponse = "[""AdfName"", ""AdfProd"", ""AdfAgfCode""]"
    If LBound(tgAdfInfo) <> UBound(tgAdfInfo) Then
        smResponse = smResponse & ","
        For ilLoop = LBound(tgAdfInfo) To UBound(tgAdfInfo) Step 1
            If tgAdfInfo(ilLoop).iCode > 0 Then
                If ilLoop <> UBound(tgAdfInfo) - 1 Then
                    smResponse = smResponse & "[" & """" & Trim$(tgAdfInfo(ilLoop).sName) & """" & "," & """" & Trim$(tgAdfInfo(ilLoop).sProd) & """" & "," & """" & Trim$(tgAdfInfo(ilLoop).iAdfAgfCode) & """" & "],"
                Else
                    smResponse = smResponse & "[" & """" & Trim$(tgAdfInfo(ilLoop).sName) & """" & "," & """" & Trim$(tgAdfInfo(ilLoop).sProd) & """" & "," & """" & Trim$(tgAdfInfo(ilLoop).iAdfAgfCode) & """" & "]"
                End If
            End If
        Next ilLoop
    End If
End Sub
Private Function mGetAdfCode(sAdfName As String) As Integer

    Dim rst_adf As ADODB.Recordset
    Dim slSQLQuery As String
    
    mGetAdfCode = -1
    smResponse = "Unknown"
    slSQLQuery = "Select adfName, adfCode from ADF_Advertisers where adfName = " & "'" & gFixQuote(sAdfName) & "'"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    If Not rst_adf.EOF Then
        smResponse = Trim(rst_adf!adfCode)
        imHeaderAdvtCode = CInt(smResponse)
    Else
        smResponse = "0"
    End If
    mGetAdfCode = imHeaderAdvtCode
    rst_adf.Close
End Function

Public Sub mProdPop(ilAdfCode As Integer)
'
'   mProdPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilAdf As Integer
    Dim blNameFd As Boolean
    Dim slSQLQuery As String
    Dim ilLoop As Integer
    Dim ilFound As Boolean
    Dim slTemp As String
    
'    If cbcProduct.ListCount > 0 Then
'        Exit Sub
'    End If
    If cbcProduct.ListIndex >= 0 Then
        slName = cbcProduct.GetName(cbcProduct.ListIndex)
    Else
        'Use advertiser default if exist
        For ilAdf = 0 To UBound(tgAdfInfo) - 1 Step 1
            If tgAdfInfo(ilAdf).iCode = ilAdfCode Then
                slName = Trim$(tgAdfInfo(ilAdf).sProd)
                Exit For
            End If
        Next ilAdf
    End If
    
    blNameFd = False
    slSQLQuery = "select * from prf_Product_Names Where prfAdfCode = " & ilAdfCode & " Order by prfName"
    Set prf_rst = cnn.Execute(slSQLQuery)
    cbcProduct.Clear
    ilFound = False
    slTemp = ""
    Do While Not prf_rst.EOF
        If (prf_rst!prfState <> "D") Then
            If slTemp <> Trim$(prf_rst!prfName) Then
                cbcProduct.AddItem Trim$(prf_rst!prfName)
                cbcProduct.SetItemData = prf_rst!prfCode
                If slName = Trim$(prf_rst!prfName) Then
                    blNameFd = True
                End If
                ilFound = True
                slTemp = Trim$(prf_rst!prfName)
            End If
        End If
        prf_rst.MoveNext
    Loop
    
    smResponse = "[""ProdName""]"
    If ilFound Then
        smResponse = smResponse & ","
        For ilLoop = 0 To cbcProduct.ListCount Step 1
            If cbcProduct.GetName(ilLoop) <> "" Then
                If ilLoop <> cbcProduct.ListCount - 1 Then
                    smResponse = smResponse & "[" & """" & Trim$(cbcProduct.GetName(ilLoop)) & """" & "],"
                Else
                    smResponse = smResponse & "[" & """" & Trim$(cbcProduct.GetName(ilLoop)) & """" & "]"
                End If
            End If
        Next ilLoop
    End If
    If blNameFd Then
        cbcProduct.Text = slName
    End If
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub mDaypartPop(ilVefCode As Integer)
    Dim slName As String
    Dim ilRif As Integer
    Dim ilRdf As Integer
    ReDim ilRdfCode(0 To 0) As Integer
    Dim blFound As Boolean
    Dim ilLoop As Integer
    
    If cbcDaypart.ListIndex >= 0 Then
        slName = cbcDaypart.GetName(cbcDaypart.ListIndex)
    End If
    If cbcDaypart.ListCount <= 0 Then
        For ilRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
            If tmRif(ilRif).iVefCode = ilVefCode Then
                For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                    If tmRif(ilRif).iRdfCode = tmRdf(ilRdf).iCode Then
                        'Test for duplicate
                        blFound = False
                        For ilLoop = 0 To UBound(ilRdfCode) - 1 Step 1
                            If ilRdfCode(ilLoop) = tmRdf(ilRdf).iCode Then
                                blFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not blFound Then
                            cbcDaypart.AddItem Trim$(tmRdf(ilRdf).sName)
                            cbcDaypart.SetItemData = tmRdf(ilRdf).iCode
                            ilRdfCode(UBound(ilRdfCode)) = tmRdf(ilRdf).iCode
                            ReDim Preserve ilRdfCode(0 To UBound(ilRdfCode) + 1) As Integer
                        End If
                        Exit For
                    End If
                Next ilRdf
            End If
        Next ilRif
    End If
    If slName <> "" Then
        cbcDaypart.Text = slName
'        For ilLoop = 0 To cbcDaypart.ListCount - 1 Step 1
'            If slName = cbcDaypart.GetName(ilLoop) Then
'                cbcDaypart.ListIndex = ilLoop
'                Exit For
'            End If
'        Next ilLoop
    End If
End Sub

Private Sub mLengthPop()
    Dim ilRif As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilLen As Integer
    Dim ilTest As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim blFd As Boolean
    ReDim ilLength(0 To 0) As Integer
    
    ilMin = 9999
    ilMax = -1
    cbcLength.Clear
    For ilRif = 0 To UBound(tmRif) - 1 Step 1
        ilVefCode = tmRif(ilRif).iVefCode
        ilVpfIndex = gBinarySearchVpf(ilVefCode)
        If ilVpfIndex <> -1 Then
            For ilLen = 0 To 9 Step 1
                If tgVpf(ilVpfIndex).iSLen(ilLen) > 0 Then
                    blFd = False
                    For ilTest = 0 To UBound(ilLength) - 1 Step 1
                        If tgVpf(ilVpfIndex).iSLen(ilLen) = ilLength(ilTest) Then
                            blFd = True
                            Exit For
                        End If
                    Next ilTest
                    If Not blFd Then
                        If tgVpf(ilVpfIndex).iSLen(ilLen) < ilMin Then
                            ilMin = tgVpf(ilVpfIndex).iSLen(ilLen)
                        End If
                        If tgVpf(ilVpfIndex).iSLen(ilLen) > ilMax Then
                            ilMax = tgVpf(ilVpfIndex).iSLen(ilLen)
                        End If
                        ilLength(UBound(ilLength)) = tgVpf(ilVpfIndex).iSLen(ilLen)
                        ReDim Preserve ilLength(0 To UBound(ilLength) + 1) As Integer
                    End If
                End If
            Next ilLen
        End If
    Next ilRif
    For ilLen = ilMin To ilMax Step 1
        For ilTest = 0 To UBound(ilLength) - 1 Step 1
            If ilLen = ilLength(ilTest) Then
                cbcLength.AddItem Trim$(Str$(ilLength(ilTest)))
                cbcLength.SetItemData = ilLength(ilTest)
                Exit For
            End If
        Next ilTest
    Next ilLen
    
    smResponse = "[""Length""]"
    If cbcLength.ListCount >= 0 Then
        smResponse = smResponse & ","
    End If
    For ilLen = 0 To cbcLength.ListCount Step 1
        If cbcLength.GetName(ilLen) <> "" Then
            If ilLen <> cbcLength.ListCount - 1 Then
                smResponse = smResponse & "[" & """" & Trim$(cbcLength.GetName(ilLen)) & """" & "],"
            Else
                smResponse = smResponse & "[" & """" & Trim$(cbcLength.GetName(ilLen)) & """" & "]"
            End If
        End If
    Next ilLen
End Sub

Private Function mDetermineRateCard() As Integer
    Dim ilRet As Integer
    Dim tlRcf As RCF
    Dim llDate As Long
    Dim llNewestDate As Long
    Dim slSQLQuery As String
    
    mDetermineRateCard = 0
'    llNewestDate = -1
'    ilRet = btrGetFirst(hmRcf, tlRcf, imRcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Do While ilRet = BTRV_ERR_NONE
'        gUnpackDateLong tlRcf.iStartDate(0), tlRcf.iStartDate(1), llDate
'        If (llNewestDate = -1) Or (llDate > llNewestDate) Then
'            llNewestDate = llDate
'            tmRcf = tlRcf
'        End If
'        ilRet = btrGetNext(hmRcf, tlRcf, imRcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    gRCRead ProgrammaticBuy, tmRcf.iCode, tmRcf, tmRif(), tmRdf()
    
    imRcfCode = 0
    ReDim tmRif(0 To 0) As RIF
    ReDim tmRdf(0 To 0) As RDF
    
    slSQLQuery = "Select * from rcf_Rate_Card Order by rcfStartDate desc"
    Set rcf_rst = cnn.Execute(slSQLQuery)
    If Not rcf_rst.EOF Then
        imRcfCode = rcf_rst!rcfCode
        smRcfName = rcf_rst!rcfName
        mCreateUDTForRcf
    End If
    
    slSQLQuery = "Select * from rif_Rate_Card_Items Where rifRcfCode = " & imRcfCode
    Set rif_rst = cnn.Execute(slSQLQuery)
    Do While Not rif_rst.EOF
        mCreateUDTForRif
        rif_rst.MoveNext
    Loop
   
    slSQLQuery = "Select * from rdf_Standard_Daypart Order by rdfName"
    Set rdf_rst = cnn.Execute(slSQLQuery)
    Do While Not rdf_rst.EOF
        mCreateUDTForRdf
        rdf_rst.MoveNext
    Loop
    mDetermineRateCard = imRcfCode
End Function

Private Function mGetPrice(ilVefCode As Integer, ilRdfCode As Integer, ilLen As Integer) As String
    Dim ilRif As Integer
    Dim llDollars As Long
    Dim slPrice As String
    For ilRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
        If tmRif(ilRif).iVefCode = ilVefCode Then
            If tmRif(ilRif).iRdfCode = ilRdfCode Then
                'Use start date to obtain the Price
                llDollars = gGetWkDollars(0, "1/1/" & tmRcf.iYear, tmRif(ilRif).lRate())
            End If
        End If
    Next ilRif
    slPrice = gLongToStrDec(llDollars, 0)
    mAdjustPrice ilLen, slPrice
    mGetPrice = slPrice
    
End Function

Private Function mFindItemByName(slName As String, cbcList As CSI_ComboBoxList) As Integer
    Dim ilLoop As Integer
    mFindItemByName = -1
    For ilLoop = 0 To cbcList.ListCount - 1 Step 1
        If slName = cbcList.GetName(ilLoop) Then
            mFindItemByName = ilLoop
            Exit For
        End If
    Next ilLoop
End Function
Private Function mFindItemByCode(llCode As Long, cbcList As CSI_ComboBoxList) As Integer
    Dim ilLoop As Integer
    mFindItemByCode = -1
    For ilLoop = 0 To cbcList.ListCount - 1 Step 1
        If llCode = cbcList.GetItemData(ilLoop) Then
            mFindItemByCode = ilLoop
            Exit For
        End If
    Next ilLoop
End Function
Private Sub mSpotsPer()
    cbcPer.AddItem "Day"
    cbcPer.SetItemData = 0
    cbcPer.AddItem "Week"
    cbcPer.SetItemData = 1
    cbcPer.AddItem "All"
    cbcPer.SetItemData = 2
End Sub
Private Sub mComputeTotals()
    Dim llRow As Long
    Dim slTotalSpots As String
    Dim slTotalPrice As String
    Dim slTotalGRPPct As String
    Dim slTotalGRP As String
    Dim slRowPrice As String
    Dim slNumberWeeks As String
    Dim slPrice As String
    Dim slNumberAirDays As String
    Dim slStr As String
    Dim llGrImp As Long
    Dim slPop As String
    Dim slGrImp As String
    
    gSetMousePointer grdHeader(imTokenIdx), grdLines(imTokenIdx), vbHourglass
    slTotalSpots = "0"
    slTotalPrice = "0"
    smTtlPrice = "0"
    slRowPrice = "0"
    slTotalGRPPct = "0"
    smTtlPercentGRP = "0"
    slTotalGRP = "0"
    smTtlGRP = "0"
    smTtlCPP = "0"
    slPop = ""
    slGrImp = "0"
    slNumberWeeks = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, NUMBERWEEKSINDEX)
    For llRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
        If grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX) <> "" And grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX) <> "" Then
            If grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX) <> "" Then
                If (imTabIndex = 0) Or (imTabIndex = 1) Or ((imTabIndex = 2) And (grdLines(imTokenIdx).TextMatrix(llRow, SPOTSPERINDEX) <> "")) Then
                    slPrice = grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX)
                    slPrice = Replace(slPrice, ",", "")
                    If (imTabIndex = 0) Or (imTabIndex = 1) Then
                        If slPop = "" Then
                            slPop = grdLines(imTokenIdx).TextMatrix(llRow, POPULATIONINDEX)
                        ElseIf slPop <> grdLines(imTokenIdx).TextMatrix(llRow, POPULATIONINDEX) Then
                            slPop = "0"
                        End If
                        slGrImp = gAddStr(slGrImp, gMulStr(grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX), grdLines(imTokenIdx).TextMatrix(llRow, AUDIENCEINDEX)))
                        slTotalSpots = gAddStr(slTotalSpots, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX))
                        slRowPrice = gMulStr(slPrice, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX))
                        slTotalGRPPct = gAddStr(slTotalGRPPct, grdLines(imTokenIdx).TextMatrix(llRow, GRPGRIMPPCTINDEX))
                        slTotalGRP = gAddStr(slTotalGRP, grdLines(imTokenIdx).TextMatrix(llRow, LNGRPINDEX))
                    Else
                        Select Case Trim$(grdLines(imTokenIdx).TextMatrix(llRow, SPOTSPERINDEX))
                            Case "All"
                                slTotalSpots = gAddStr(slTotalSpots, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX))
                                slRowPrice = gMulStr(slPrice, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX))
                            Case "Day"
                                slNumberAirDays = mAirDayCount(llRow)
                                slTotalSpots = gAddStr(slTotalSpots, gMulStr(gMulStr(slNumberWeeks, slNumberAirDays), grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX)))
                                slRowPrice = gMulStr(slPrice, gMulStr(slNumberAirDays, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX)))
                                slRowPrice = gMulStr(slRowPrice, slNumberWeeks)
                                grdLines(imTokenIdx).TextMatrix(llRow, TOTALPRICEINDEX) = slRowPrice
                            Case "Week"
                                slTotalSpots = gAddStr(slTotalSpots, gMulStr(slNumberWeeks, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX)))
                                slRowPrice = gMulStr(slPrice, grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX))
                                slRowPrice = gMulStr(slRowPrice, slNumberWeeks)
                                grdLines(imTokenIdx).TextMatrix(llRow, TOTALPRICEINDEX) = slRowPrice
                                
                        End Select
                    End If
                    slStr = slRowPrice
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    
                    'This is the price per row for price tab
                    grdLines(imTokenIdx).TextMatrix(llRow, TOTALPRICEINDEX) = slStr
                    smPriceRowPrice = slStr
                    'this is the grand total of all prices
                    slTotalPrice = gAddStr(slTotalPrice, slRowPrice)
                End If
            End If
        End If
    Next llRow
    If slTotalSpots <> "0" Then
        'Total rows at bottom for price tab
        lacTotalSpots.Caption = slTotalSpots
        smTtlSpots = slTotalSpots
        
        slStr = slTotalPrice
        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        lacTotalPrice.Caption = slStr
        
        'Total price at bottom for price tab
        smTtlPrice = slStr
        If (imTabIndex = 0) Or (imTabIndex = 1) Then
            lacTotalGrpPct.Caption = slTotalGRPPct
            smTtlPercentGRP = slTotalGRPPct
            lacTotalGRP.Caption = slTotalGRP
            smTtlGRP = slTotalGRP
            If imTabIndex = 0 Then
                If (slPop = "") Or (slPop = "0") Then
                    slStr = gDivStr(slTotalPrice, slTotalGRP)
                Else
                    slStr = gDivStr(gRoundStr(gDivStr(gMulStr(slTotalPrice, slPop), slGrImp), ".50", 0), "100.00")
                End If
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                lacTotalCPP.Caption = slStr
                smTtlCPP = slStr
            Else
                slStr = gDivStr(slTotalPrice, slTotalGRP)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                lacTotalCPP.Caption = slStr
                smTtlCPP = lacTotalCPP.Caption
            End If
        Else
            lacTotalGrpPct.Caption = ""
            smTtlPercentGRP = ""
            lacTotalGRP.Caption = ""
            smTtlGRP = ""
            lacTotalCPP.Caption = ""
            smTtlCPP = lacTotalCPP.Caption
            smTtlCPP = ""
        End If
    Else
        lacTotalSpots.Caption = ""
        smTtlSpots = ""
            lacTotalPrice.Caption = ""
        smTtlPrice = ""
        lacTotalGrpPct.Caption = ""
        smTtlPercentGRP = ""
        lacTotalGRP.Caption = ""
        smTtlGRP = ""
        lacTotalCPP.Caption = ""
        smTtlCPP = ""
    End If
    gSetMousePointer grdHeader(imTokenIdx), grdLines(imTokenIdx), vbDefault
End Sub

Private Sub mAdjustPrice(ilLen As Integer, slPrice As String)
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slDollarAdj As String
    Dim slIndexAdj As String
    
    ilIndex = -1
    If ilLen > 0 Then
        For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
            If tmRcf.iLen(ilLoop) = ilLen Then
                ilIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    If ilIndex <> -1 Then
        Select Case tmRcf.sUseLen
            Case "D"    'Delta
                slDollarAdj = gLongToStrDec(tmRcf.lValue(ilIndex), 2)
                slPrice = gAddStr(slPrice, slDollarAdj)
            Case "I"    'Index
                slIndexAdj = gLongToStrDec(tmRcf.lValue(ilIndex), 2)
                slPrice = gMulStr(slPrice, slIndexAdj)
        End Select
    End If

End Sub

Private Function mAirDayCount(llRow As Long) As String
    Dim slDayPart As String
    Dim ilDp As Integer
    Dim ilDay As Integer
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim blDay(0 To 6) As Boolean
    
    For ilDay = 0 To 6 Step 1
        blDay(ilDay) = False
    Next ilDay
    slDayPart = grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX)
    For ilDp = 0 To UBound(tmRdf) - 1 Step 1
        If Trim$(tmRdf(ilDp).sName) = slDayPart Then
            For ilLoop = LBound(tmRdf(ilDp).sWkDays, 1) To UBound(tmRdf(ilDp).sWkDays, 1) Step 1
                For ilDay = 0 To 6 Step 1
                    If tmRdf(ilDp).sWkDays(ilLoop, ilDay) = "Y" Then
                        blDay(ilDay) = True
                    End If
                Next ilDay
            Next ilLoop
            Exit For
        End If
    Next ilDp
    ilCount = 0
    For ilDay = 0 To 6 Step 1
        If blDay(ilDay) Then
            ilCount = ilCount + 1
        End If
    Next ilDay
    mAirDayCount = Trim$(Str$(ilCount))
End Function


Private Sub mPopLineGrid(llStartRow As Long)
    Dim ilLen As Integer
    Dim ilRif As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilRdfCode As Integer
    Dim ilTest As Integer
    Dim llRow As Long
    Dim ilVef As Integer
    Dim slNetwork As String
    Dim ilRdf As Integer
    Dim slDayPart As String
    Dim llCount As Long
    Dim slStr As String
    Dim llCPM As Long
    
    'D.S. 7/18/18 - Moved the below from under the IF statement to above it.
    llRow = llStartRow
    If grdLines(imTokenIdx).TextMatrix(llRow, LENGTHINDEX) = "" Then
        Exit Sub
    End If

    ilLen = Val(grdLines(imTokenIdx).TextMatrix(llStartRow, LENGTHINDEX))
    If ilLen = 0 Then
        Exit Sub
    End If
    gSetMousePointer grdHeader(imTokenIdx), grdLines(imTokenIdx), vbHourglass
    For ilRif = 0 To UBound(tmRif) - 1 Step 1
        ilVefCode = tmRif(ilRif).iVefCode
        ilVpfIndex = gBinarySearchVpf(ilVefCode)
        ilRdfCode = tmRif(ilRif).iRdfCode
        If ilVpfIndex <> -1 Then
            For ilTest = 0 To 9 Step 1
                If (tgVpf(ilVpfIndex).iSLen(ilTest) = ilLen) Then
                    'Add row
                    ilVef = mFindItemByCode(CLng(ilVefCode), cbcNetwork)
                    If ilVef <> -1 Then
                        slNetwork = cbcNetwork.GetName(ilVef)
                    End If
                    slDayPart = ""
                    For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                        If ilRdfCode = tmRdf(ilRdf).iCode Then
                            slDayPart = Trim$(tmRdf(ilRdf).sName)
                            Exit For
                        End If
                    Next ilRdf
                    If (ilVef <> -1) And (slDayPart <> "") Then
                        If llRow >= grdLines(imTokenIdx).Rows Then
                            grdLines(imTokenIdx).AddItem ""
                        End If
                        grdLines(imTokenIdx).TextMatrix(llRow, LENGTHINDEX) = grdLines(imTokenIdx).TextMatrix(llStartRow, LENGTHINDEX)
                        grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX) = slNetwork
                        grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX) = slDayPart
                        grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX) = mGetPrice(ilVefCode, ilRdfCode, ilLen)
                        grdLines(imTokenIdx).TextMatrix(llRow, AVAILINDEX) = mGetAvailCount(llRow)
                        slStr = mGetCPP(llRow)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        grdLines(imTokenIdx).TextMatrix(llRow, CPPCPMINDEX) = slStr
                        grdLines(imTokenIdx).TextMatrix(llRow, STATIONINDEX) = mGetStationCount()
                        llRow = llRow + 1
                    End If
                End If
            Next ilTest
        End If
    Next ilRif
    If llRow >= grdLines(imTokenIdx).Rows Then
        grdLines(imTokenIdx).AddItem ""
    End If
    gSetMousePointer grdHeader(imTokenIdx), grdLines(imTokenIdx), vbDefault
End Sub


Private Function mGetAvailCount(llRow As Long) As String
    Dim ilVefCode As Integer
    Dim slStartDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llWkStartDate As Long
    Dim llWkEndDate As Long
    Dim ilWk As Integer
    Dim llDate As Long
    ReDim ilWkCount(0 To 5) As Integer
    Dim ilVef As Integer
    Dim ilPVF As Integer
    Dim slNetwork As String
    Dim slSQLQuery As String
    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim ilSec As Integer
    Dim ilUnits As Integer
    Dim slStr As String
    Dim slUnits As String
    Dim slSpotLen As String
    Dim ilSpot As Integer
    Dim ilLen As Integer
    Dim ilVpfIndex As Integer
    Dim ilBuyerSpotLen As Integer
    Dim ilCount As Integer
    Dim ilRdf As Integer
    Dim ilRdfIndex As Integer
    Dim slDayPart As String
    Dim ilDp As Integer
    Dim ilDay As Integer
    Dim ilRow As Integer
    Dim ilLoop As Integer
    Dim slInOut As String
    Dim ilAnfCode As Integer
    Dim blDay(0 To 6, 0 To 6) As Boolean
    Dim llStartTime(0 To 6) As Long
    Dim llEndTime(0 To 6) As Long
    Dim llTime As Long
    Dim blDayFd As Boolean
    Dim blCountEqOrGt10 As Boolean
    Dim blFound As Boolean
    Dim blInRC As Boolean
    Dim ilRif As Integer
    Dim ilPkgVefCode As Integer
    
    mGetAvailCount = ""
    ReDim imVefCode(0 To 0) As Integer
    If (grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX) = "") Or (grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX) = "") Or (grdLines(imTokenIdx).TextMatrix(llRow, LENGTHINDEX) = "") Then
        gSetMousePointer grdHeader, grdLines(imTokenIdx), vbDefault
        Exit Function
    End If
    For ilWk = 0 To 5 Step 1
        ilWkCount(ilWk) = -1
    Next ilWk
    slNetwork = grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX)
    ilVef = mFindItemByName(slNetwork, cbcNetwork)
    If ilVef = -1 Then
        Exit Function
    End If
    ilVefCode = cbcNetwork.GetItemData(ilVef)
    ilPkgVefCode = ilVefCode
    slSQLQuery = "Select vefPvfCode From vef_Vehicles Where vefCode = " & ilVefCode
    Set vef_rst = cnn.Execute(slSQLQuery)
    If vef_rst.EOF Then
        Exit Function
    End If
    ilBuyerSpotLen = Val(grdLines(imTokenIdx).TextMatrix(llRow, LENGTHINDEX))
    For ilRow = 0 To 6 Step 1
        For ilDay = 0 To 6 Step 1
            blDay(ilRow, ilDay) = False
        Next ilDay
        llStartTime(ilRow) = 0
        llEndTime(ilRow) = -1
    Next ilRow
    slDayPart = grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX)
    For ilDp = 0 To UBound(tmRdf) - 1 Step 1
        If Trim$(tmRdf(ilDp).sName) = slDayPart Then
            slInOut = Trim$(tmRdf(ilDp).sInOut)
            ilAnfCode = tmRdf(ilDp).ianfCode
            ilRow = 0
            For ilLoop = LBound(tmRdf(ilDp).sWkDays, 2) To UBound(tmRdf(ilDp).sWkDays, 2) Step 1
                If (tmRdf(ilDp).iStartTime(0, ilLoop) <> 1) Or (tmRdf(ilDp).iStartTime(1, ilLoop) <> 0) Then
                    blDayFd = False
                    For ilDay = 0 To 6 Step 1
                        'D.S.
                        'If tmRdf(ilDp).sWkDays(ilLoop, ilDay + 1) = "Y" Then
                        If tmRdf(ilDp).sWkDays(ilLoop, ilDay) = "Y" Then
                            blDay(ilRow, ilDay) = True
                            blDayFd = True
                        End If
                    Next ilDay
                    If blDayFd Then
                        gUnpackTimeLong tmRdf(ilDp).iStartTime(0, ilLoop), tmRdf(ilDp).iStartTime(1, ilLoop), False, llStartTime(ilRow)
                        gUnpackTimeLong tmRdf(ilDp).iEndTime(0, ilLoop), tmRdf(ilDp).iEndTime(1, ilLoop), True, llEndTime(ilRow)
                        ilRow = ilRow + 1
                    End If
                End If
            Next ilLoop
            Exit For
        End If
    Next ilDp

    
    slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & vef_rst!vefPvfCode
    Set pvf_rst = cnn.Execute(slSQLQuery)
    Do While Not pvf_rst.EOF
        mCreateUDTForPvf
        For ilPVF = 0 To UBound(tmPvf.iVefCode) Step 1
            blInRC = False
            If tmPvf.iVefCode(ilPVF) > 0 Then
                ilVpfIndex = gBinarySearchVpf(tmPvf.iVefCode(ilPVF))
                ilRdfIndex = -1
                For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                    If tmPvf.iRdfCode(ilPVF) = tmRdf(ilRdf).iCode Then
                        ilRdfIndex = ilRdf
                        Exit For
                    End If
                Next
                If ilRdfIndex <> -1 Then
                    For ilRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
                        If tmRif(ilRif).iVefCode = tmPvf.iVefCode(ilPVF) Then
                            If tmRif(ilRif).iRdfCode = tmPvf.iRdfCode(ilPVF) Then
                                blInRC = True
                            End If
                        End If
                    Next ilRif
                End If
            End If
            If (tmPvf.iVefCode(ilPVF) > 0) And (ilVpfIndex <> -1) And (ilRdfIndex <> -1 And (blInRC)) Then
                ilVefCode = tmPvf.iVefCode(ilPVF)
                blFound = False
                For ilVef = 0 To UBound(imVefCode) - 1 Step 1
                    If imVefCode(ilVef) = ilVefCode Then
                        blFound = True
                        Exit For
                    End If
                Next ilVef
                If Not blFound Then
                    imVefCode(UBound(imVefCode)) = ilVefCode
                    ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                End If
                slStartDate = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, STARTDATEINDEX)
                llWkStartDate = gDateValue(slStartDate)
                llWkEndDate = llWkStartDate + 6
                For ilWk = 0 To 5 Step 1
                    ilCount = 0
                    blCountEqOrGt10 = False
                    For llDate = llWkStartDate To llWkEndDate Step 1
                        For ilRdf = 0 To ilRow - 1 Step 1
                            ilDay = gWeekDayLong(llDate)
                            If blDay(ilRdf, ilDay) Then
                                gPackDateLong llDate, ilDate0, ilDate1
                                tmSsfSrchKey2.iVefCode = ilVefCode
                                tmSsfSrchKey2.iDate(0) = ilDate0
                                tmSsfSrchKey2.iDate(1) = ilDate1
                                ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode)
                                    If (ilDate0 = tmSsf.iDate(0)) And (ilDate1 = tmSsf.iDate(1)) Then
                                        ilEvt = 1
                                        Do While ilEvt <= tmSsf.iCount
                                            'LSet tmAvail = tmSsf.tPas(ilEvt)
                                            LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                                                If ((slInOut = "") Or (slInOut = "N")) Or ((slInOut = "I") And (tmAvail.ianfCode = ilAnfCode)) Or ((slInOut = "O") And (tmAvail.ianfCode <> ilAnfCode)) Then
                                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                                    If (llTime >= llStartTime(ilRdf)) And (llTime < llEndTime(ilRdf)) Then
                                                        If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Or (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                                                            ilUnits = tmAvail.iAvInfo And &H1F
                                                            slUnits = Trim$(Str$(ilUnits)) & ".0"   'For units as thirty
                                                            ilSec = tmAvail.iLen
                                                        Else
                                                            ilUnits = tmAvail.iAvInfo And &H1F
                                                            slUnits = Trim$(Str$(ilUnits)) & ".0"   'For units as thirty
                                                            ilSec = 0
                                                        End If
                                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                            ilEvt = ilEvt + 1
                                                            LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                            ilLen = (tmSpot.iPosLen And &HFFF)
                                                            If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                                                                ilUnits = ilUnits - 1
                                                                ilSec = ilSec - ilLen   '(tmSpot.iPosLen And &HFFF)
                                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                                ilUnits = ilUnits - 1
                                                                ilSec = ilSec - ilLen   '(tmSpot.iPosLen And &HFFF)
                                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                                slSpotLen = Trim$(Str$(ilLen))  'tmSpot.iPosLen And &HFFF))
                                                                slStr = gDivStr(slSpotLen, "30.0")
                                                                slUnits = gSubStr(slUnits, slSpotLen)
                                                            End If
                                                        Next ilSpot
                                                        If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Or (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                                                            Do While ilUnits > 0 And ilSec >= ilBuyerSpotLen
                                                                ilCount = ilCount + 1
                                                                ilUnits = ilUnits - 1
                                                                ilSec = ilSec - ilBuyerSpotLen
                                                            Loop
                                                        Else
                                                            slSpotLen = Trim$(Str$(ilBuyerSpotLen))
                                                            slStr = gDivStr(slSpotLen, "30.0")
                                                            Do While gSubStr(slUnits, slSpotLen) >= 0
                                                                ilCount = ilCount + 1
                                                                slUnits = gSubStr(slUnits, slSpotLen)
                                                            Loop
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If ilCount \ tmPvf.iNoSpot(ilPVF) >= 10 Then
                                                blCountEqOrGt10 = True
                                                Exit For
                                            End If
                                            ilEvt = ilEvt + 1
                                        Loop
                                    Else
                                        Exit Do
                                    End If
                                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                        Next ilRdf
                        If blCountEqOrGt10 Then
                            Exit For
                        End If
                    Next llDate
                    ilCount = ilCount \ tmPvf.iNoSpot(ilPVF)
                    If ilWkCount(ilWk) = -1 Then
                        ilWkCount(ilWk) = ilCount
                    ElseIf ilCount < ilWkCount(ilWk) Then
                        ilWkCount(ilWk) = ilCount
                    End If
                    llWkStartDate = llWkEndDate + 1
                    llWkEndDate = llWkStartDate + 6
                Next ilWk
            End If
        Next ilPVF
        If tmPvf.lLkPvfCode <= 0 Then
            Exit Do
        End If
        slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & tmPvf.lLkPvfCode
        Set pvf_rst = cnn.Execute(slSQLQuery)
    Loop
    slStr = ""
    For ilWk = 0 To 5 Step 1
        If ilWkCount(ilWk) = -1 Then
            slStr = slStr & "0"
        ElseIf ilWkCount(ilWk) < 10 Then
            slStr = slStr & Trim(Str(ilWkCount(ilWk)))
        Else
            slStr = slStr & "+"
        End If
    Next ilWk
    mGetAvailCount = slStr
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mGetPkgAud                      *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Package Audience            *
'*                                                     *
'*******************************************************
Private Function mGetCPP(llRow As Long) As String

    Dim slNetwork As String
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim ilMnfDemo As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfSocEco As Integer
    Dim ilRdf As Integer
    Dim ilRdfCode As Integer
    Dim llPop As Long
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    ReDim ilDays(0 To 6) As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slSQLQuery As String
    Dim ilPVF As Integer
    Dim llPopEst As Long
    Dim llPkgPop As Long
    Dim llAvgAud As Long
    Dim llPkgAud As Long
    Dim slCPP As String
    Dim slCPM As String
    Dim llRafCode As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilDay As Integer
    Dim tlRdf As RDF
    Dim ilTime As Integer
    Dim slStartDate As String
    Dim llWkStartDate As Long
    Dim llWkEndDate As Long
    Dim llGrImp As Long
    Dim ilRif As Integer
    Dim blInRC As Boolean
    
    
    If (grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX) = "") Or (grdLines(imTokenIdx).TextMatrix(llRow, DAYPARTINDEX) = "") Or (grdLines(imTokenIdx).TextMatrix(llRow, LENGTHINDEX) = "") Then
        Exit Function
    End If

    slNetwork = grdLines(imTokenIdx).TextMatrix(llRow, NETWORKINDEX)
    ilVef = mFindItemByName(slNetwork, cbcNetwork)
    If ilVef = -1 Then
        Exit Function
    End If
    
    ilVefCode = cbcNetwork.GetItemData(ilVef)

    slSQLQuery = "Select vefPvfCode From vef_Vehicles Where vefCode = " & ilVefCode
    Set vef_rst = cnn.Execute(slSQLQuery)
    If vef_rst.EOF Then
        Exit Function
    End If

    llOvStartTime = 0
    llOvEndTime = 0
    If cbcDemo.ListIndex < 0 Then
        ilMnfDemo = 0
    Else
        slNameCode = tgDemoCode(cbcDemo.ListIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfDemo = Val(slCode)
    End If
    imHeaderDemo = ilMnfDemo
    llPkgPop = -1
    llGrImp = 0
    
    slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & vef_rst!vefPvfCode
    Set pvf_rst = cnn.Execute(slSQLQuery)
    Do While Not pvf_rst.EOF
        mCreateUDTForPvf
        For ilPVF = 0 To UBound(tmPvf.iVefCode) Step 1
            ilVefCode = -1
            ilRdfCode = -1
            ilDnfCode = -1
            For ilDay = 0 To 6 Step 1
                ilDays(ilDay) = False
            Next ilDay
            If tmPvf.iVefCode(ilPVF) > 0 Then
                ilVefCode = tmPvf.iVefCode(ilPVF)
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    ilDnfCode = tgMVef(ilVef).iDnfCode
                End If
                For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                    If tmPvf.iRdfCode(ilPVF) = tmRdf(ilRdf).iCode Then
                        ilRdfCode = tmRdf(ilRdf).iCode
                        
                        LSet tlRdf = tmRdf(ilRdf)
                        For ilTime = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1  'Row
                            If (tlRdf.iStartTime(0, ilTime) <> 1) Or (tlRdf.iStartTime(1, ilTime) <> 0) Then
                                'D.S. 12/26/17 For ilDay = 1 To 7 Step 1
                                For ilDay = 0 To 6 Step 1
                                    If tlRdf.sWkDays(ilTime, ilDay) = "Y" Then
                                        ilDays(ilDay) = True
                                    End If
                                Next ilDay
                            End If
                        Next ilTime
                        Exit For
                    End If
                Next
            End If
            blInRC = False
            For ilRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
                If tmRif(ilRif).iVefCode = ilVefCode Then
                    If tmRif(ilRif).iRdfCode = ilRdfCode Then
                        blInRC = True
                    End If
                End If
            Next ilRif
            If (blInRC) And (tmPvf.iVefCode(ilPVF) > 0) And (ilVefCode <> -1) And (ilRdfCode <> -1) And (ilDnfCode <> -1) Then
                slStartDate = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, STARTDATEINDEX)
                llWkStartDate = gDateValue(slStartDate)
                llWkEndDate = llWkStartDate + 6
            
                ilMnfSocEco = 0
                If (ilDnfCode > 0) And (ilMnfDemo > 0) Then
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
                Else
                    llPop = 0
                End If
                If llPkgPop = -1 Then
                    llPkgPop = llPop
                ElseIf llPkgPop <> llPop Then
                    llPkgPop = 0
                End If
                llRafCode = 0
                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, ilMnfDemo, llWkStartDate, llWkEndDate, tlRdf.iCode, llOvStartTime, llOvEndTime, ilDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                If ilRet Then
                    llGrImp = llGrImp + llAvgAud * tmPvf.iNoSpot(ilPVF)
                End If
            End If
        Next ilPVF
        If tmPvf.lLkPvfCode <= 0 Then
            Exit Do
        End If
        slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & tmPvf.lLkPvfCode
        Set pvf_rst = cnn.Execute(slSQLQuery)
    Loop
    llPkgAud = llGrImp / 1
    grdLines(imTokenIdx).TextMatrix(llRow, AUDIENCEINDEX) = llPkgAud
    grdLines(imTokenIdx).TextMatrix(llRow, POPULATIONINDEX) = llPkgPop
    grdLines(imTokenIdx).TextMatrix(llRow, RATINGINDEX) = 0
    If imTabIndex = 0 Then
        'Compute CPP from Total Cost / #Spots*Rating = Total Cost * pop / #Spots*Aud = Total Cost * Pop / Audinence ( Assume 1 spot)
        slCPP = gDivStr(gMulStr(gDivStr(grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX), "100.00"), Str(llPkgPop)), Str(llPkgAud))
        mGetCPP = slCPP
    Else
        If tgSpf.sSAudData = "H" Then
            slCPM = gDivStr(gMulStr(grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX), "10.00"), Str(llGrImp))
        ElseIf tgSpf.sSAudData = "N" Then
            slCPM = gDivStr(gMulStr(grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX), "100.00"), Str(llGrImp))
        ElseIf tgSpf.sSAudData = "U" Then
            slCPM = gDivStr(gMulStr(grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX), "1000.00"), Str(llGrImp))
        Else
            slCPM = gDivStr(grdLines(imTokenIdx).TextMatrix(llRow, SPOTPRICEINDEX), Str(llGrImp))
        End If
        mGetCPP = slCPM
    End If
    
End Function

Private Sub mDemoPop()
    Dim ilRet As Integer
    'gDemoPop cbcDemo
    ilRet = gPopMnfPlusFieldsBox(ProgrammaticBuy, cbcDemo, tgDemoCode(), sgDemoCodeTag, "D")
End Sub

Private Sub mComputeSpots(llRow As Long)
    'Once % entered: Entered GRP in header * % GRP.  # Spots = GRP * % GRP/Rating = GRP * $ GRP * Pop/ Audience
    Dim slHdGRP As String
    Dim slGRP As String
    Dim slGRPPct As String
    Dim slHdGrImp As String
    Dim slGrImp As String
    Dim llGrImp As Long
    Dim slGrImpPct As String
    Dim slPop As String
    Dim slAud As String
    Dim slSpots As String
    
    If imTabIndex = 2 Then
        Exit Sub
    End If
    slPop = grdLines(imTokenIdx).TextMatrix(llRow, POPULATIONINDEX)
    slAud = grdLines(imTokenIdx).TextMatrix(llRow, AUDIENCEINDEX)
    If tgSpf.sSAudData = "H" Then
        slPop = gLongToStrDec(Val(slPop), 1)
        slAud = gLongToStrDec(Val(slAud), 1)
    ElseIf tgSpf.sSAudData = "N" Then
        slPop = gLongToStrDec(Val(slPop), 2)
        slAud = gLongToStrDec(Val(slAud), 2)
    ElseIf tgSpf.sSAudData = "U" Then
        slPop = gLongToStrDec(Val(slPop), 3)
        slAud = gLongToStrDec(Val(slAud), 3)
    Else
        'slGrImp = Trim$(Str$(llGrimp))
    End If
    grdLines(imTokenIdx).TextMatrix(llRow, RATINGINDEX) = gDivStr(gAddStr(slAud, ".00"), slPop)
    If imTabIndex = 0 Then
        slHdGRP = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, TARGETINDEX)
        slGRPPct = gDivStr(grdLines(imTokenIdx).TextMatrix(llRow, GRPGRIMPPCTINDEX), "100.00")
        slGRP = gMulStr(slHdGRP, slGRPPct)
        'GRP = 1000*GrImp/Pop= 1000*#Spots*Audience/Population
        '#Spots = (GRP*Popuation)/(1000*Audience)
        slSpots = gRoundStr(gDivStr(gMulStr(slGRP, slPop), gMulStr(slAud, "100")), ".50", 0)
        slGRP = gDivStr(gMulStr(gMulStr(gAddStr(slSpots, ".00"), slAud), "100"), slPop)
        grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX) = slSpots
        grdLines(imTokenIdx).TextMatrix(llRow, LNGRPINDEX) = slGRP
    Else
        slHdGrImp = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, TARGETINDEX)
        slGrImpPct = gDivStr(grdLines(imTokenIdx).TextMatrix(llRow, GRPGRIMPPCTINDEX), "100.00")
        slGrImp = gMulStr(slHdGrImp, slGrImpPct)
        slSpots = gRoundStr(gDivStr(slGrImp, slAud), ".50", 0)
        slGrImp = gMulStr(slSpots, slAud)
        
'        If tgSpf.sSAudData = "H" Then
'            slGrImp = gLongToStrDec(llGrimp, 1)
'        ElseIf tgSpf.sSAudData = "N" Then
'            slGrImp = gLongToStrDec(llGrimp, 2)
'        ElseIf tgSpf.sSAudData = "U" Then
'            slGrImp = gLongToStrDec(llGrimp, 3)
'        Else
'            slGrImp = Trim$(Str$(llGrimp))
'        End If
        grdLines(imTokenIdx).TextMatrix(llRow, NUMBERSPOTSINDEX) = slSpots
        grdLines(imTokenIdx).TextMatrix(llRow, LNGRPINDEX) = slGrImp
        slGRP = gDivStr(gMulStr(gMulStr(gAddStr(slSpots, ".00"), slAud), "100"), slPop)
    End If
    grdLines(imTokenIdx).TextMatrix(llRow, RATINGINDEX) = gDivStr(gAddStr(slGRP, ".00"), slSpots)
End Sub

Private Sub mTypePop()
    cbcType.AddItem "CPP/GRP"
    cbcType.AddItem "CPM/GrImp"
End Sub

Private Function mGetStationCount() As Long
    Dim slStr As String
    Dim ilVef As Integer
    Dim ilVefIndex As Integer
    Dim ilAir As Integer
    Dim ilTest As Integer
    Dim tlVef As VEF
    Dim blFound As Boolean
    Dim slSQLQuery As String
    Dim slStartDate As String
    Dim llCount As Long
    ReDim ilVefCode(0 To 0) As Integer
    
    slStartDate = grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, STARTDATEINDEX)
    For ilVef = 0 To UBound(imVefCode) - 1 Step 1
        ilVefIndex = gBinarySearchVef(imVefCode(ilVef))
        If ilVefIndex <> -1 Then
            If tgMVef(ilVefIndex).sType = "S" Then
                LSet tlVef = tgMVef(ilVefIndex)
                gBuildLinkArray hmVLF, tlVef, slStartDate, ilVefCode()
                For ilAir = 0 To UBound(ilVefCode) - 1 Step 1
                    blFound = False
                    For ilTest = 0 To UBound(imVefCode) - 1 Step 1
                        If imVefCode(ilTest) = ilVefCode(ilAir) Then
                            blFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not blFound Then
                        imVefCode(UBound(imVefCode)) = ilVefCode(ilAir)
                        ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                    End If
                Next ilAir
            End If
        End If
    Next ilVef
    slStr = ""
    For ilVef = 0 To UBound(imVefCode) - 1 Step 1
        If slStr = "" Then
            slStr = Str(imVefCode(ilVef))
        Else
            slStr = slStr & "," & Str(imVefCode(ilVef))
        End If
    Next ilVef
    llCount = 0
    If slStr <> "" Then
        slSQLQuery = "Select Count(distinct attShfCode) as shttCount from att Where attOnAir <= '" & Format(slStartDate, sgSQLDateForm) & "' And  attOffAir >= '" & Format(slStartDate, sgSQLDateForm) & "'" & " And attVefCode In (" & slStr & ")"
        Set att_rst = cnn.Execute(slSQLQuery)
        If Not att_rst.EOF Then
            llCount = att_rst!shttCount
        Else
            llCount = 0
        End If
    End If
    mGetStationCount = llCount
End Function

Private Sub tscProgrammaticBuy_Click()
    If imTerminate Then
        Exit Sub
    End If
    mSetShow
    'If wanting to retain grid and totals, this is the place to add the code
    If imTabIndex = 0 Then
    ElseIf imTabIndex = 1 Then
    Else
    End If
    grdHeader(imTokenIdx).Clear
    grdLines(imTokenIdx).Clear
    lacTotalSpots.Caption = ""
    lacTotalPrice.Caption = ""
    lacTotalGrpPct.Caption = ""
    lacTotalGRP.Caption = ""
    lacTotalCPP.Caption = ""
    smTtlCPP = ""
    imTabIndex = tscProgrammaticBuy.SelectedItem.Index - 1
    If imTabIndex <= 2 Then
        mHeaderColumnNames
        mLineColumnNames
        mHeaderColumnWidths
        mLineColumnWidths
        mSetTitles
        mSetTotals
        frcTab.Visible = True
    Else
        frcTab.Visible = False
    End If
    'If grid and totals saved, restore code goes here
    If imTabIndex = 0 Then
    ElseIf imTabIndex = 1 Then
    Else
    End If
    mSetCommands
End Sub

Public Sub mSetTitles()
    If imTabIndex = 0 Then
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = "Target GRP's"
        grdLines(imTokenIdx).TextMatrix(0, CPPCPMINDEX) = "CPP"
        grdLines(imTokenIdx).TextMatrix(0, GRPGRIMPPCTINDEX) = "% GRP"
        grdLines(imTokenIdx).TextMatrix(0, LNGRPINDEX) = "GRP"
    ElseIf imTabIndex = 1 Then
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = "Target GrImp's"
        grdLines(imTokenIdx).TextMatrix(0, CPPCPMINDEX) = "CPM"
        grdLines(imTokenIdx).TextMatrix(0, GRPGRIMPPCTINDEX) = "% GrImp"
        grdLines(imTokenIdx).TextMatrix(0, LNGRPINDEX) = "GrImp"
    Else
        grdHeader(imTokenIdx).TextMatrix(0, TARGETINDEX) = ""
        grdLines(imTokenIdx).TextMatrix(0, CPPCPMINDEX) = "CPM"
        grdLines(imTokenIdx).TextMatrix(0, GRPGRIMPPCTINDEX) = "% GrImp"
        grdLines(imTokenIdx).TextMatrix(0, LNGRPINDEX) = "GrImp"
    End If

End Sub

Public Sub mSetTotals()
    Dim llColPos As Long
    Dim ilCol As Integer
    
    If (imTabIndex = 0) Or (imTabIndex = 1) Then
        lacTotalCPP.Visible = True
        lacTotalGrpPct.Visible = True
        lacTotalGRP.Visible = True
        llColPos = 0
        For ilCol = 0 To NUMBERSPOTSINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalSpots.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalSpots.Width = grdLines(imTokenIdx).ColWidth(NUMBERSPOTSINDEX) - 30
        llColPos = 0
        For ilCol = 0 To TOTALPRICEINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalPrice.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalPrice.Width = grdLines(imTokenIdx).ColWidth(TOTALPRICEINDEX) - 30
        
        llColPos = 0
        For ilCol = 0 To CPPCPMINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalCPP.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalCPP.Width = grdLines(imTokenIdx).ColWidth(CPPCPMINDEX) - 30
        
        llColPos = 0
        For ilCol = 0 To GRPGRIMPPCTINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalGrpPct.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalGrpPct.Width = grdLines(imTokenIdx).ColWidth(GRPGRIMPPCTINDEX) - 30
        
        llColPos = 0
        For ilCol = 0 To LNGRPINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalGRP.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalGRP.Width = grdLines(imTokenIdx).ColWidth(LNGRPINDEX) - 30
    Else
        lacTotalCPP.Visible = False
        lacTotalGrpPct.Visible = False
        lacTotalGRP.Visible = False
        llColPos = 0
        For ilCol = 0 To NUMBERSPOTSINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalSpots.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalSpots.Width = grdLines(imTokenIdx).ColWidth(NUMBERSPOTSINDEX) - 30
        llColPos = 0
        For ilCol = 0 To TOTALPRICEINDEX - 1 Step 1
            llColPos = llColPos + grdLines(imTokenIdx).ColWidth(ilCol) + 15
        Next ilCol
        lacTotalPrice.Left = llColPos + grdLines(imTokenIdx).Left
        lacTotalPrice.Width = grdLines(imTokenIdx).ColWidth(TOTALPRICEINDEX) - 30
    End If
End Sub


Private Sub mProgTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    xgInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub

Private Sub btn_Stop_Click()
    sckClient.Close
    lblStatus.Caption = "Stopped"
    lblStatus.ForeColor = vbRed
    gLogMsg "<ProgMsg> Socket Services Closed", "ProgrammaticBuy.Txt", False
    cmdGo.Enabled = True
    btn_Stop.Enabled = False
End Sub

'*******************************************************************************
' Connect to the server.
'*******************************************************************************
Public Sub cmdGo_Click()
    
    cmdGo.Enabled = False
    btn_Stop.Enabled = True
    ilFirstResponse = 0
    gLogMsg "", "ProgrammaticBuy.Txt", False
    gLogMsg "<ProgMsg> *** Start New Connection to Comm Center ***", "ProgrammaticBuy.Txt", False
    sckClient.Connect sgCommServerIP, sgCommServerPort
    gLogMsg "<ProgMsg> Connecting To: IP: " & sgCommServerIP & " Port: " & sgCommServerPort, "ProgrammaticBuy.Txt", False
    
End Sub

Public Sub btn_Close_Click()
    sckClient.Close
    gLogMsg "<ProgMsg> Closing", "ProgrammaticBuy.Txt", False
    End
End Sub




Private Sub mGetSiteOptions()

    Dim slSQLQuery As String
    Dim ilRet As Integer
    
    slSQLQuery = "Select * From SAF_Schd_Attributes"
    Set rst = gSQLSelectCall(slSQLQuery)
    
    If Not rst.EOF Then
        ilRet = Asc(rst!safFeatures5)
        smPrgmmaticAllow = "N"
        If (ilRet And PROGRAMMATICALLOWED) = PROGRAMMATICALLOWED Then
            smPrgmmaticAllow = "Y"
         End If
        smShowAvailCount = "N"
        If (ilRet And SHOWAVAILCOUNT) = SHOWAVAILCOUNT Then
            smShowAvailCount = "Y"
        End If
        smShowCPPTab = "N"
        If (ilRet And SHOWCPPTAB) = SHOWCPPTAB Then
            smShowCPPTab = "Y"
        End If
        smShowCPMTab = "N"
        If (ilRet And SHOWCPMTAB) = SHOWCPMTAB Then
            smShowCPMTab = "Y"
        End If
        smShowPriceTab = "N"
        If (ilRet And SHOWPRICETAB) = SHOWPRICETAB Then
            smShowPriceTab = "Y"
        End If
    End If
End Sub

'*******************************************************************************
' If the connect worked, this function is called.
' Send the data to identitify ourselves.
'*******************************************************************************
Public Sub sckClient_Connect()

    smNetworkName = sgCustomerName
    smArgs = Split(Command$, "=")
    If UBound(smArgs) > 0 Then
        If smArgs(0) = "TOKEN" Then
            smNetworkName = smNetworkName & "_" & smArgs(1)
            gLogMsg "NetworkName = " & smNetworkName, "ProgrammaticBuy.Txt", False
        End If
    End If
    sckClient.SendData "IDENTITY=" & smNetworkName & "<EOF>"
    lb_Messages.AddItem "IDENTITY=" & smNetworkName & "<EOF>"
    gLogMsg "<ProgMsg> IDENTITY=" & smNetworkName & "<EOF>", "ProgrammaticBuy.Txt", False
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = "Socket Connection Established"
    gLogMsg "<ProgMsg> Socket Connection Established", "ProgrammaticBuy.Txt", False
End Sub

'*******************************************************************************
' Load all of the text from a file into a string.
'*******************************************************************************
Public Function FileText(fileName$) As String
    Dim handle As Integer
    handle = FreeFile
    Open fileName$ For Input As #handle
    FileText = Input$(LOF(handle), handle)
    Close #handle
End Function

'*******************************************************************************
' This code is called anytime we get a socket error.
'*******************************************************************************
Public Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckClient.Close
    lblStatus.Caption = Description
    gLogMsg "<Socket>  Error: " & Description, "ProgrammaticBuy.Txt", False
End Sub


Private Function mSubmit() As Boolean

    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    Dim slDate As String
    Dim slExportType As String
    Dim slFile As String
    Dim slCntr1or2PlaceRating As String
    Dim slTemp As String
    Dim ilPos As Integer
    Dim blAdvExists As Boolean
    Dim blProdExists As Boolean
    Dim slStr As String
    Dim llIdx As Long
    
    'D.S. 02/27/15 need to set transaction here
    
    mSubmit = False
    gObtainAgency
    igAgfCode = tgTokenInfo(UBound(tgTokenInfo) - 1).iAgfCode
    blAdvExists = mAdvertiserExists(smHeaderAdvt)
    If Not blAdvExists Then
        ilRet = AdvNew()
    End If
    smHeaderProd = gFixQuote(smHeaderProd)
    blProdExists = mProductExists(smHeaderProd)
    If Not blProdExists Then
        ilRet = ProdNew()
    End If
    ilRet = AdvProdUpdate()
    If Not mBuildHeader Then
        gLogMsg "ProgrammaticBuy.frm-mBuildHeader returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If
    If Not mInsertHdrComment() Then
        gLogMsg "ProgrammaticBuy.frm-mInsertHdrComment returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If
    'D.S. 2/27/18 on insert do loop to retry in the event that the cntr no is no longer available
    llIdx = 0
    Do
        llIdx = llIdx + 1
        slSQLQuery = "Select spfCNextNo From SPF_Site_Options"
        Set rst_Temp = gSQLSelectCall(slSQLQuery)
        If Not rst_Temp.EOF Then
            tmChf.lCntrNo = rst_Temp!spfCNextNo + llIdx
        Else
            ilRet = ilRet
            'message jeff failure and exit
        End If
        ilRet = btrInsert(hmCHF, tmChf, imCHFRecLen, INDEXKEY0)
    Loop While ilRet = BTRV_ERR_DUPLICATE_KEY

    If ilRet <> BTRV_ERR_NONE Then
        'send failure to jeff
        ilRet = ilRet
    End If
    slSQLQuery = "update SPF_Site_Options set spfCNextNo = " & tmChf.lCntrNo & " where spfcnextno < " & tmChf.lCntrNo
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
       'error
        ilRet = ilRet
    End If
    ilRet = mBuildLines
    If Not ilRet Then
        gLogMsg "ProgrammaticBuy.frm-mBuildLines returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If
    ilRet = mBuildFlights
    If Not ilRet Then
        gLogMsg "ProgrammaticBuy.frm-mBuildFlights returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If
    ilRet = mBuildHiddenLinesAndFlights
    If Not ilRet Then
        gLogMsg "ProgrammaticBuy.frm-mBuildHiddenLinesAndFlights returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If
    
    ReDim tgClfCntr(0 To UBound(tmClf)) As CLFLIST
    For ilClf = 0 To UBound(tmClf) - 1 Step 1
        tgClfCntr(ilClf).ClfRec = tmClf(ilClf)
    Next ilClf
    mCLFListSet
    
    ReDim tgCffCntr(0 To UBound(tmCff)) As CFFLIST
    For ilCff = 0 To UBound(tmCff) - 1 Step 1
        tgCffCntr(ilCff).CffRec = tmCff(ilCff)
    Next ilCff
    mCFFListSet
    '7/4/18
    mDistributeSpotsByWk
    ilRet = mSetHiddenLinePrice
    If Not ilRet Then
        gLogMsg "ProgrammaticBuy.frm-mSetHiddenLinePrice returned an error.", "ProgrammaticBuy.Txt", False
        Exit Function
    End If

    tgChfCntr = tmChf
    tgChf = tgChfCntr
    slCntr1or2PlaceRating = gSet1or2PlaceRating(tgChfCntr.iAgfCode)
    ''''''ReDim tgClfCntr(0 To UBound(tmClf)) As CLFLIST
    
    '7/3/18
'    For ilClf = 0 To UBound(tmClf) - 1 Step 1
'        tgClfCntr(ilClf).ClfRec = tmClf(ilClf)
'    Next ilClf
'    mCLFListSet
'
'    ReDim tgCffCntr(0 To UBound(tmCff)) As CFFLIST
'    For ilCff = 0 To UBound(tmCff) - 1 Step 1
'        tgCffCntr(ilCff).CffRec = tmCff(ilCff)
'    Next ilCff
'    mCFFListSet
    
    '7/4/18
    'mDistributeSpotsByWk
    
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        tgClfCntr(ilClf).ClfRec.lCode = 0
        ilRet = btrInsert(hmClf, tgClfCntr(ilClf).ClfRec, imClfRecLen, INDEXKEY2)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = ilRet
            'handle error
        End If
    Next ilClf
    
    For ilCff = 0 To UBound(tgCffCntr) - 1 Step 1
        tgCffCntr(ilCff).CffRec.lCode = 0
        ilRet = btrInsert(hmCff, tgCffCntr(ilCff).CffRec, imCffRecLen, INDEXKEY1)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = ilRet
            'handle error
        End If
    Next ilCff
    'these are for Darlene setting up the BR
    ReDim tgClf(0 To UBound(tgClfCntr)) As CLFLIST
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        tgClf(ilClf) = tgClfCntr(ilClf)
    Next ilClf
    'mCFFListSet
    'these are for Darlene setting up the BR
    ReDim tgCff(0 To UBound(tgCffCntr)) As CFFLIST
    For ilCff = 0 To UBound(tgCffCntr) - 1 Step 1
        tgCff(ilCff) = tgCffCntr(ilCff)
    Next ilCff
    gObtainRcfRifRdf
    For ilClf = 0 To UBound(imGrdClfRow) - 1 Step 1
        If imTabIndex <> 2 Then
            If slCntr1or2PlaceRating <> "2" Then
                slStr = gRoundStr(gAddStr(gMulStr(grdLines(imTokenIdx).TextMatrix(imGrdClfRow(ilClf), RATINGINDEX), "10."), ".5"), "0", 0)
            Else
                slStr = gRoundStr(gAddStr(gMulStr(grdLines(imTokenIdx).TextMatrix(imGrdClfRow(ilClf), RATINGINDEX), "100."), ".5"), "0", 0)
            End If
            If slStr <> "" Then
                tgClfCntr(ilClf).iAvgRating = slStr
            Else
                tgClfCntr(ilClf).iAvgRating = 0
            End If
            tgClfCntr(ilClf).lAvgAud = grdLines(imTokenIdx).TextMatrix(imGrdClfRow(ilClf), AUDIENCEINDEX)
            tgClf(ilClf).lAvgAud = tgClfCntr(ilClf).lAvgAud
            tgClf(ilClf).iAvgRating = tgClfCntr(ilClf).iAvgRating
        Else
            tgClfCntr(ilClf).iAvgRating = 0
            tgClfCntr(ilClf).lAvgAud = 0
            tgClf(ilClf).lAvgAud = 0
            tgClf(ilClf).iAvgRating = 0
        End If
    Next ilClf

    igStationXmlChoice = AGENCYXML
    slExportType = "X"  'Trim$(tmAgf.sCntrExptForm)
    slFile = mExportProposal(slExportType, False, False)
    ReDim tgFBSbf(0 To 0) As SBFLIST
    tgFBSbf(0).iStatus = -1 'Not Used
    tgFBSbf(0).lRecPos = 0
    ReDim tgIBSbf(0 To 0) As SBFLIST
    tgIBSbf(0).iStatus = -1 'Not Used
    tgIBSbf(0).lRecPos = 0
    ReDim tgMBSbf(0 To 0) As SBFLIST
    tgMBSbf(0).iStatus = -1 'Not Used
    tgMBSbf(0).lRecPos = 0
    ReDim tgPBSbf(0 To 0) As SBFLIST
    tgPBSbf(0).iStatus = -1 'Not Used
    tgPBSbf(0).lRecPos = 0
    ReDim tgCgfCntr(0 To 0) As CGFLIST
    tgCgfCntr(0).iStatus = -1 'Not Used
    tgCgfCntr(0).iNextCgf = -1
    ReDim tgMsfCntr(0 To 0) As MSFLIST
    tgMsfCntr(0).iStatus = -1 'Not Used
    tgMsfCntr(0).iFirstMgf = -1
    tgMsfCntr(0).iCxfIndex = -1
    ReDim tgMgfCntr(0 To 0) As MGFLIST
    tgMgfCntr(0).iStatus = -1 'Not Used
    ReDim sgMsfCntrCxf(0 To 0) As String
    mGenPDF
    'mSendAgencyEmail
    slDate = Format$(gNow(), "m/d/yy")
    ilRet = gAlertAdd("C", "P", tmChf.lCode, 0, slDate)
    If Len(Trim(tmChf.lCntrNo)) > 0 Then
        smResponse = "Proposal #" & Trim(tmChf.lCntrNo) & " was successfully processed!" & vbCrLf & vbCrLf & "An e-mail has been sent to you. Attached to the e-mail is a copy of this proposal and XML file."
    Else
        smResponse = "No Proposal Number was Generated"
    End If
    mSubmit = True
End Function


Private Function mBuildHeader() As Boolean

    Dim rst_Temp As ADODB.Recordset
    Dim slSQLQuery As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim ilMnfDemo As Integer
    Dim slCode As String
    Dim ilIdx As Integer
    Dim ilRow As Integer
    Dim slStr As String
    
    mBuildHeader = False
    tmChf = gDefaultCHF
    tmChf.sAgyEstNo = grdHeader(imTokenIdx).TextMatrix(1, ESTIMATENUMBERINDEX)
    mGetAdfCode grdHeader(imTokenIdx).TextMatrix(1, ADVERTISERINDEX)
    tmChf.iAdfCode = imHeaderAdvtCode
    lgAdfCode = tmChf.iAdfCode
    tmChf.sProduct = Trim$(grdHeader(imTokenIdx).TextMatrix(1, PRODUCTINDEX))
    tmChf.iMnfComp(0) = imMnfComp(0)
    tmChf.iMnfComp(1) = imMnfComp(1)
    tmChf.iMnfExcl(0) = imMnfExcl(0)
    tmChf.iMnfExcl(1) = imMnfExcl(1)
    For ilIdx = LBound(tgTokenInfo) To UBound(tgTokenInfo) Step 1
        If Trim(tgTokenInfo(ilIdx).sCommToken) = Trim(smCurrentToken) Then
            mBuyerExists ilIdx, tmChf.iPnfBuyer, tmChf.sPhone
            'Rate card used
            tmChf.iRcfCode = tgTokenInfo(ilIdx).iRcfCode
            slSQLQuery = "select agfCode from AGF_Agencies where agfName = " & "'" & Trim(tgTokenInfo(ilIdx).sAgency) & "'"
            Set rst_Temp = cnn.Execute(slSQLQuery)
            If Not rst_Temp.EOF Then
                tmChf.iAgfCode = rst_Temp!agfCode
            Else
                'Error condition? Add the agency?  - failure exit
                tmChf.iAgfCode = 0
            End If
            Exit For
        End If
    Next ilIdx
    slSQLQuery = "Select mnfCode from MNF_Multi_Names where mnfName = " & "'" & grdHeader(imTokenIdx).TextMatrix(1, DEMOINDEX) & "'" & " and mnftype = 'D'"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        If imTabIndex = 0 Then
            tmChf.sCppCpm = "P"
        ElseIf imTabIndex = 1 Then
            tmChf.sCppCpm = "M"
        Else
            tmChf.sCppCpm = ""
        End If
        
        tmChf.iMnfDemo(0) = rst_Temp!mnfCode
    Else
        tmChf.iMnfDemo(0) = 0
        tmChf.sCppCpm = ""
    End If
    gPackDate grdHeader(imTokenIdx).TextMatrix(1, STARTDATEINDEX), tmChf.iStartDate(0), tmChf.iStartDate(1)
    gPackDate grdHeader(imTokenIdx).TextMatrix(1, ENDDATEINDEX), tmChf.iEndDate(0), tmChf.iEndDate(1)
    'C = complete, W = Working
    tmChf.sStatus = "C"
    tmChf.iEDSSentExtRevNo = -2
    tmChf.sSchStatus = "P"
    tmChf.sSource = "P"
    tmChf.iSlfCode(0) = imSlfCode
    igSlfCode = imSlfCode
    tmChf.lComm(0) = 1000000
    tmChf.iUrfCode = imUrfCode
    'D.S. 3/1/18
    gUnformatStr smTtlPrice, 0, slStr
    tmChf.lInputGross = gStrDecToLong(slStr, 2)
    tmChf.lAirTimeGross = gStrDecToLong(slStr, 2)
    mBuildHeader = True
    
End Function

Private Function mBuildLines() As Boolean

    Dim ilRow As Integer
    Dim ilLineIdx As Integer
    Dim ilLineCount As Integer
    Dim ilVefCode As Integer
    Dim ilRdf As Integer
    Dim ilRdfCode As Integer
    Dim ilVef As Integer

    mBuildLines = False
    imWhichGrid = 1
    ilLineCount = 0
    For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
        If Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)) <> "" And Val(grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)) > 0 Then
            ilLineCount = ilLineCount + 1
        End If
    Next ilRow
    If ilLineCount <= 0 Then
        mBuildLines = False
        Exit Function
    End If
    ReDim tmClf(0 To ilLineCount)
    ReDim imGrdClfRow(0 To ilLineCount)
    ilLineIdx = 0
    For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
        If grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX) <> "" And Val(grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)) > 0 Then
            tmClf(ilLineIdx) = mDefaultCLF
            tmClf(ilLineIdx).lCode = ilRow
            tmClf(ilLineIdx).lChfCode = tmChf.lCode
            tmClf(ilLineIdx).iLine = ilLineIdx + 1
            ilVefCode = gGetVehCodeByVefName(Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, NETWORKINDEX)), False)
            tmClf(ilLineIdx).iVefCode = ilVefCode
            tmClf(ilLineIdx).iMnfDemo = tmChf.iMnfDemo(0)
            
            ilVef = gBinarySearchVef(ilVefCode)
            If ilVef <> -1 Then
                tmClf(ilLineIdx).iDnfCode = tgMVef(ilVef).iDnfCode
            Else
                tmClf(ilLineIdx).iDnfCode = 0
            End If
            
            'D.S. maybe later set these
            If smTabType = "CPP" Then
                tmClf(ilLineIdx).lCPM = 0
                tmClf(ilLineIdx).lCPP = 0
                tmClf(ilLineIdx).lGrImp = 0
            ElseIf smTabType = "CPM" Then
                tmClf(ilLineIdx).lCPM = 0
                tmClf(ilLineIdx).lCPP = 0
                tmClf(ilLineIdx).lGrImp = 0
             Else
                tmClf(ilLineIdx).lCPM = 0
                tmClf(ilLineIdx).lCPP = 0
                tmClf(ilLineIdx).lGrImp = 0
            End If
            
            imGrdClfRow(ilLineIdx) = ilRow
            
            tmClf(ilLineIdx).iRdfCode = gGetRdfCode(Trim$(grdLines(imTokenIdx).TextMatrix(ilRow, DAYPARTINDEX)))
            tmClf(ilLineIdx).iStartDate(0) = tmChf.iStartDate(0)
            tmClf(ilLineIdx).iStartDate(1) = tmChf.iStartDate(1)
            tmClf(ilLineIdx).iEndDate(0) = tmChf.iEndDate(0)
            tmClf(ilLineIdx).iEndDate(1) = tmChf.iEndDate(1)
            tmClf(ilLineIdx).iPkLineNo = 0
            'below 1 = weekly and 0 = daily if tab is proce type
            tmClf(ilLineIdx).sCntPct = 1
            tmClf(ilLineIdx).iLen = grdLines(imTokenIdx).TextMatrix(ilRow, LENGTHINDEX)
            'live copy = R ask jim
            tmClf(ilLineIdx).sLiveCopy = "R"
            'verfify -2
            tmClf(ilLineIdx).iEDSSentExtRevNo = -2
            tmClf(ilLineIdx).iUrfCode = 1
            tmClf(ilLineIdx).sType = "O"
            ilLineIdx = ilLineIdx + 1
        End If
    Next ilRow
    mBuildLines = True
End Function

Private Function mBuildFlights() As Boolean

    Dim ilRow As Integer
    Dim ilIdx As Integer
    Dim llCurrRowNum As Integer
    Dim ilLineCount As Integer
    Dim ilNoWks As Integer

    mBuildFlights = False
    imWhichGrid = 1
    ilLineCount = 0
    ilNoWks = Val(grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, NUMBERWEEKSINDEX))
    ReDim tmCff(0 To UBound(tmClf))
    For ilRow = grdLines(imTokenIdx).FixedRows To grdLines(imTokenIdx).Rows - 1 Step 1
        If grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX) <> "" And Val(grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)) > 0 Then
            tmCff(ilLineCount) = mDefaultCFF
            tmCff(ilLineCount).lChfCode = tmChf.lCode
            tmCff(ilLineCount).iClfLine = ilLineCount + 1
            tmCff(ilLineCount).iStartDate(0) = tmChf.iStartDate(0)
            tmCff(ilLineCount).iStartDate(1) = tmChf.iStartDate(1)
            tmCff(ilLineCount).iEndDate(0) = tmChf.iEndDate(0)
            tmCff(ilLineCount).iEndDate(1) = tmChf.iEndDate(1)
            'D.S. 02/27/18 below for CPP/CPM is weeklyd and Price can be daily or weekly
            mDPAllowedDays tmClf(ilLineCount).iRdfCode
            'D.S. 02/27/18 below for CPP/CPM is weekly and price can be daily or weekly
            If imTabIndex <> 2 Or (imTabIndex = 2 And grdLines(imTokenIdx).TextMatrix(ilRow, SPOTSPERINDEX) <> "Day") Then
                tmCff(ilLineCount).sDyWk = "W"      'set to D for daily W for weekly
                If imTabIndex <> 2 Then
                    tmCff(ilLineCount).iSpotsWk = grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)
                Else
                    tmCff(ilLineCount).iSpotsWk = grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX) * ilNoWks
                End If
                For ilIdx = 0 To 6
                    If bmAllowDays(ilIdx) Then
                        tmCff(ilLineCount).iDay(ilIdx) = 1        'set all of these for weekly 1 or daily use the num spots per day based
                    Else
                        tmCff(ilLineCount).iDay(ilIdx) = 0
                    End If
                Next ilIdx                              'on the daypart days only valid days
            Else
                tmCff(ilLineCount).sDyWk = "D"      'set to D for daily W for weekly
                tmCff(ilLineCount).iSpotsWk = 0
                For ilIdx = 0 To 6
                    If bmAllowDays(ilIdx) Then
                        'set all of these for weekly 1 or daily use the num spots per day based
                        tmCff(ilLineCount).iDay(ilIdx) = grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)
                    Else
                        tmCff(ilLineCount).iDay(ilIdx) = 0
                    End If
                Next ilIdx                              'on the daypart days only valid days

            End If
            'D.S. 3/1/18
            'tmCff(ilLineCount).lActPrice = CLng((grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX) / grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)))
            If grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX) <> ".00" Then
                If imTabIndex <> 2 Then
                    tmCff(ilLineCount).lActPrice = gStrDecToLong((grdLines(imTokenIdx).TextMatrix(ilRow, TOTALPRICEINDEX) / grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX)), 2)
                    tmCff(ilLineCount).lPropPrice = tmCff(ilLineCount).lActPrice
                Else
                    tmCff(ilLineCount).lActPrice = gStrDecToLong(grdLines(imTokenIdx).TextMatrix(ilRow, SPOTPRICEINDEX), 2)
                    tmCff(ilLineCount).lPropPrice = tmCff(ilLineCount).lActPrice
                End If
            End If
            ilLineCount = ilLineCount + 1
        End If
    Next ilRow
    mBuildFlights = True
End Function

Private Function mBuildHiddenLinesAndFlights() As Boolean
    
    Dim ilLine As Integer
    Dim ilVef As Integer
    Dim llPVFCode As Long
    Dim ilPkgLnNo As Integer
    Dim slSQLQuery As String
    Dim ilPVF As Integer
    Dim blInRC As Boolean
    Dim ilLnUpper As Integer
    Dim ilFlUpper As Integer
    Dim ilLineNo As Integer
    Dim ilRdfIndex As Integer
    Dim ilRdf As Integer
    Dim ilRif As Integer
    Dim ilRow As Integer
    Dim ilIdx As Integer
    
    mBuildHiddenLinesAndFlights = False
    ilPkgLnNo = UBound(tmClf) - 1
    ilLineNo = tmClf(ilPkgLnNo).iLine
    For ilLine = 0 To ilPkgLnNo Step 1
        'D.S determine if pkg line
        ilVef = gBinarySearchVef(tmClf(ilLine).iVefCode)
        If ilVef = -1 Then
            'failed to find the vehicle
            Exit Function
        End If
        If tgMVef(ilVef).sType = "P" And tgMVef(ilVef).lPvfCode > 0 Then
            'Create the hidden lines and flights
            llPVFCode = tgMVef(ilVef).lPvfCode
            slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & llPVFCode
            Set pvf_rst = cnn.Execute(slSQLQuery)
            Do While Not pvf_rst.EOF
                mCreateUDTForPvf
                For ilPVF = 0 To UBound(tmPvf.iVefCode) Step 1
                    blInRC = False
                    If tmPvf.iVefCode(ilPVF) > 0 Then
                        ilRdfIndex = -1
                        For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                            If tmPvf.iRdfCode(ilPVF) = tmRdf(ilRdf).iCode Then
                                ilRdfIndex = ilRdf
                                Exit For
                            End If
                        Next
                        If ilRdfIndex <> -1 Then
                            For ilRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
                                If tmRif(ilRif).iVefCode = tmPvf.iVefCode(ilPVF) Then
                                    If tmRif(ilRif).iRdfCode = tmPvf.iRdfCode(ilPVF) Then
                                        blInRC = True
                                        Exit For
                                    End If
                                End If
                            Next ilRif
                        End If
                        If (ilRdfIndex <> -1) And blInRC Then
                            'D.S. 02/27/18 We failed, no hidden line daypart infomation
                            ilLnUpper = UBound(tmClf)
                            ReDim Preserve tmClf(0 To UBound(tmClf) + 1)
                            'copying the package line into the hidden line image
                            tmClf(ilLnUpper) = tmClf(ilLine)
                            'setting unique values in the hidden lines
                            tmClf(ilLnUpper).iPkLineNo = tmClf(ilLine).iLine
                            ilLineNo = ilLineNo + 1
                            tmClf(ilLnUpper).iLine = ilLineNo
                            tmClf(ilLnUpper).iVefCode = tmPvf.iVefCode(ilPVF)
                            ilVef = gBinarySearchVef(tmPvf.iVefCode(ilPVF))
                            If ilVef <> -1 Then
                                tmClf(ilLnUpper).iDnfCode = tgMVef(ilVef).iDnfCode
                            Else
                                tmClf(ilLnUpper).iDnfCode = 0
                            End If
                            tmClf(ilLnUpper).iRdfCode = tmPvf.iRdfCode(ilPVF)
                            tmClf(ilLnUpper).sType = "H"
                            tmClf(ilLnUpper).lCode = tmPvf.iPctRate(ilPVF)
                            'create the unique flight info
                            ilFlUpper = UBound(tmCff)
                            ReDim Preserve tmCff(0 To UBound(tmCff) + 1)
                            'copying the package flight into the hidden line flight
                            tmCff(ilFlUpper) = tmCff(ilLine)
                            'Setting the unique values in the hidden flights
                            tmCff(ilFlUpper).iClfLine = ilLineNo
                            'D.S. 02/27/18 below the days are based on the daypart definition
                            mDPAllowedDays tmClf(ilLnUpper).iRdfCode
                            ilRow = tmClf(ilLine).lCode
                            If imTabIndex <> 2 Or imTabIndex = 2 And grdLines(imTokenIdx).TextMatrix(ilRow, SPOTSPERINDEX) <> "Day" Then
                                'Weekly
                                tmCff(ilFlUpper).sDyWk = "W"      'set to D for daily W for weekly
                                'tmCff(ilFlUpper).iSpotsWk = grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX) * tmPvf.iNoSpot(ilPVF)
                                tmCff(ilFlUpper).iSpotsWk = tmPvf.iNoSpot(ilPVF)
                                For ilIdx = 0 To 6
                                    If bmAllowDays(ilIdx) Then
                                        tmCff(ilFlUpper).iDay(ilIdx) = 1        'set all of these for weekly 1 or daily use the num spots per day based
                                    Else
                                        tmCff(ilFlUpper).iDay(ilIdx) = 0
                                    End If
                                Next ilIdx
                                'on the daypart days only valid days
                            Else
                                'Daily
                                tmCff(ilFlUpper).sDyWk = "D"      'set to D for daily W for weekly
                                'tmCff(ilFlUpper).iSpotsWk = grdLines(imTokenIdx).TextMatrix(ilRow, NUMBERSPOTSINDEX) * tmPvf.iNoSpot(ilPVF)
                                tmCff(ilFlUpper).iSpotsWk = tmPvf.iNoSpot(ilPVF)
                                For ilIdx = 0 To 6
                                    If bmAllowDays(ilIdx) Then
                                        'set all of these for weekly 1 or daily use the num spots per day based
                                        tmCff(ilFlUpper).iDay(ilIdx) = 1 'tmPvf.iNoSpot(ilPVF)
                                    Else
                                        tmCff(ilFlUpper).iDay(ilIdx) = 0
                                    End If
                                Next ilIdx
                                'on the daypart days only valid days
                            
                            End If
                            tmCff(ilFlUpper).lPropPrice = gStrDecToLong(mGetPrice(tmPvf.iVefCode(ilPVF), tmPvf.iRdfCode(ilPVF), tmClf(ilLine).iLen), 2)
                            tmCff(ilFlUpper).lActPrice = tmCff(ilFlUpper).lPropPrice
                        End If
                    End If
                Next ilPVF
                If tmPvf.lLkPvfCode <= 0 Then
                    Exit Do
                End If
                slSQLQuery = "Select * From pvf_Package_Vehicle Where pvfCode = " & tmPvf.lLkPvfCode
                'D.S. 02/28/18  change method for sql call to gsqnowait...
                Set pvf_rst = cnn.Execute(slSQLQuery)
            Loop
        End If
    Next ilLine
    mBuildHiddenLinesAndFlights = True
    
End Function


Sub mDPAllowedDays(ilRdfCode As Integer)

    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim ilLIndex As Integer
    ilLIndex = LBound(bmAllowDays)
    
    For ilLoop = LBound(bmAllowDays) To UBound(bmAllowDays) Step 1
        bmAllowDays(ilLoop) = False
    Next ilLoop
    
    ilRdf = -1
    For ilLoop = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
        If tmRdf(ilLoop).iCode = ilRdfCode Then
            ilRdf = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilRdf <> -1 Then
        For ilIndex = LBound(tmRdf(ilRdf).iStartTime, 2) To UBound(tmRdf(ilRdf).iStartTime, 2) Step 1
            If (tmRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tmRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                ilLIndex = LBound(bmAllowDays)
                For ilDay = 1 To 7 Step 1
                    If tmRdf(ilRdf).sWkDays(ilIndex, ilDay - 1) = "Y" Then
                        bmAllowDays(ilLIndex) = True
                    End If
                    ilLIndex = ilLIndex + 1
                Next ilDay
            End If
        Next ilIndex
        Exit Sub
    End If
    Exit Sub
End Sub


Private Function mOpenFiles() As Boolean

    Dim ilRet As Integer
    
    mOpenFiles = False
    
    ReDim tmClf(0 To 0)
    ReDim tmCff(0 To 0)
    
    imCHFRecLen = Len(tmChf)
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gLogMsg "Issue opening Chf Table", "ProgrammaticBuy.Txt", False
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    
    imClfRecLen = Len(tmClf(0))
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "CLF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gLogMsg "Issue opening Clf Table", "ProgrammaticBuy.Txt", False
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Function
    End If
    
    imCffRecLen = Len(tmCff(0))
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "CFF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gLogMsg "Issue opening Cff Table", "ProgrammaticBuy.Txt", False
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Function
    End If
    
    imCxfRecLen = Len(tmCxf)
    hmCxf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "CXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gLogMsg "Issue opening Cff Table", "ProgrammaticBuy.Txt", False
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        Exit Function
    End If
    
    imPvfRecLen = Len(tmPvf)
    hmPvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gLogMsg "Issue opening Pvf Table", "ProgrammaticBuy.Txt", False
        ilRet = btrClose(hmPvf)
        btrDestroy hmPvf
        Exit Function
    End If
    
    mOpenFiles = True
End Function


Private Sub mBuyerExists(ilIdx As Integer, iPnfCode As Integer, sPhone As String)
    
    Dim rst_Temp As ADODB.Recordset
    Dim slSQLQuery As String
    Dim ilAgfCode As String
    Dim slEmail As String
    Dim slPhone As String
    Dim slPhonePak As String
    Dim slName As String
    Dim blAddBuyer As Boolean
    Dim slCefCode As String
    Dim llPnfCode As Long
    Dim slSyncDate As String
    Dim slSyncTime As String
    
    'aDD BUYER
    
    
    iPnfCode = 0
    sPhone = ""
    slEmail = Trim(tgTokenInfo(ilIdx).sEMail)
    slPhone = Trim(tgTokenInfo(ilIdx).sPhone)
    slName = Trim(tgTokenInfo(ilIdx).sBuyer)
    ilAgfCode = tgTokenInfo(ilIdx).iAgfCode
    
    blAddBuyer = False
    slSQLQuery = "select pnfCode,pnfPhone, pnfemailcefcode from Pnf_Personnel_Names left outer join cef_comments_events on pnfemailcefcode = cefcode where pnfAgfCode = " & ilAgfCode & " and pnfType = 'B' AND cefComment = '" & slEmail & "'"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        iPnfCode = rst_Temp!pnfCode
        sPhone = rst_Temp!pnfPhone
    Else
        slSQLQuery = "select pnfCode,pnfPhone, pnfemailcefcode from Pnf_Personnel_Names where pnfAgfCode = " & ilAgfCode & " and pnfType = 'B' AND pnfName = '" & slName & "'"
        Set rst_Temp = gSQLSelectCall(slSQLQuery)
        If Not rst_Temp.EOF Then
            iPnfCode = rst_Temp!pnfCode
            sPhone = rst_Temp!pnfPhone
        Else
            blAddBuyer = True
        End If
    End If
    If InStr(1, slPhone, "Ext", vbTextCompare) Then
        mkcPhone.Mask = "(AAA) AAA-AAAA Ext(AAAA)"
    Else
        mkcPhone.Mask = "(AAA) AAA-AAAA"
    End If
    mkcPhone.Text = slPhone
    gGetPhoneNo mkcPhone, slPhonePak
    If blAddBuyer Then
    
        slSQLQuery = "Insert Into CEF_Comments_Events ( "
        slSQLQuery = slSQLQuery & "cefCode ,"
        slSQLQuery = slSQLQuery & "cefComment "
        slSQLQuery = slSQLQuery & ") "
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & "Replace" & ", "
        slSQLQuery = slSQLQuery & "'" & Trim(tgTokenInfo(imTokenIdx).sEMail) & "' "
        slSQLQuery = slSQLQuery & ") "
        slCefCode = gInsertAndReturnCode(slSQLQuery, "CEF_Comments_Events", "CefCode", "Replace")
        

        slSQLQuery = "Insert Into PNF_Personnel_Names ( "
        slSQLQuery = slSQLQuery & "pnfCode ,"                       'Code
        slSQLQuery = slSQLQuery & "pnfName ,"                       'Name
        slSQLQuery = slSQLQuery & "pnfPhone ,"                      'Phone
        slSQLQuery = slSQLQuery & "pnfEmailcefCode ,"               'Email Cef code
        slSQLQuery = slSQLQuery & "pnfAgfCode ,"                    'Agf code
        slSQLQuery = slSQLQuery & "pnfState ,"                      'State
        slSQLQuery = slSQLQuery & "pnfType "                        'Type
        slSQLQuery = slSQLQuery & ")"
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & "Replace" & ", "                           'Code
        slSQLQuery = slSQLQuery & "'" & Trim(slName) & "',"          'Name
        slSQLQuery = slSQLQuery & "'" & Trim(slPhonePak) & "',"      'Phone
        slSQLQuery = slSQLQuery & slCefCode & ","                    'Email Cef code
        slSQLQuery = slSQLQuery & ilAgfCode & ","                    'Agf code
        slSQLQuery = slSQLQuery & "'" & "A" & "',"                   'State
        slSQLQuery = slSQLQuery & "'" & "B" & "'"                    'Type
        slSQLQuery = slSQLQuery & ")"
        llPnfCode = gInsertAndReturnCode(slSQLQuery, "PNF_Personnel_Names", "PnfCode", "Replace")

        slSQLQuery = "UPDATE PNF_Personnel_Names SET "
        slSQLQuery = slSQLQuery & "pnfRemoteID = " & tgUrf(0).iRemoteUserID & ","
        slSQLQuery = slSQLQuery & "pnfAutoCode = " & llPnfCode & ","
        slSQLQuery = slSQLQuery & "pnfSyncDate = " & "'" & Format$(gNow, sgSQLDateForm) & "',"
        slSQLQuery = slSQLQuery & "pnfSynctime = " & "'" & Format$(gNow, sgSQLTimeForm) & "'"
        slSQLQuery = slSQLQuery & " WHERE (pnfCode = " & llPnfCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "ProgrammaticBuy.Txt", "ProgrammaticBuy-mBuyerExists3"
            gLogMsg "Update Error SQL =" & "[" & slSQLQuery & "]", "ProgrammaticBuy.Txt", False
        End If
    Else
        slSQLQuery = "UPDATE CEF_Comments_Events SET "
        slSQLQuery = slSQLQuery & "cefComment = " & "'" & Trim(tgTokenInfo(imTokenIdx).sEMail) & "'"
        slSQLQuery = slSQLQuery & " WHERE (cefCode = " & "'" & slCefCode & "')"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "ProgrammaticBuy.Txt", "ProgrammaticBuy-mBuyerExists2"
        End If
        
        slSQLQuery = "UPDATE PNF_Personnel_Names SET "
        slSQLQuery = slSQLQuery & "pnfName = " & "'" & slName & "',"
        slSQLQuery = slSQLQuery & "pnfPhone = " & "'" & slPhonePak & "'"
        slSQLQuery = slSQLQuery & " WHERE (pnfCode = " & iPnfCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "ProgrammaticBuy.Txt", "ProgrammaticBuy-mBuyerExists3"
        End If
    End If
    'talk to dick about updating the agf
End Sub

Private Function mSetHiddenLinePrice() As Boolean

    Dim ilPkClf As Integer
    Dim ilPkCff As Integer
    Dim ilPkVef As Integer
    Dim ilHClf As Integer
    Dim ilHCff As Integer
    Dim ilHVef As Integer
    Dim ilHStartIndex As Integer
    Dim ilHEndIndex As Integer
    Dim ilFlStartIndex As Integer
    Dim ilFlEndIndex As Integer
    Dim blPropPriceExist As Boolean
    Dim slStr As String
    Dim llRafCode As Long
    Dim ilRet As Integer
    Dim ilPassDnfCode As Integer
    Dim llAvgAud As Long
    Dim llAPrice As Long
    Dim llPPrice As Long
    Dim ilAnyRated As Integer
    Dim ilDay As Integer
    Dim slPkLineType As String
    Dim ilLp As Integer
    Dim llStartDate As Long
    Dim llDate As Long
    Dim ilSpots As Integer
    Dim ilSpotCount As Integer
    Dim ilVef As Integer
    
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    ReDim ilAudDays(0 To 6) As Integer
    Dim ilMnfDemo As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfSocEco As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim llReallDate As Long
    ReDim tlVef(0 To UBound(tgMVef)) As VEF
    Dim ilClf As Integer
    Dim ilCff As Integer
        
    mSetHiddenLinePrice = False
    
    ReDim tmClf(0 To UBound(tgClfCntr)) As CLF
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        tmClf(ilClf) = tgClfCntr(ilClf).ClfRec
    Next ilClf

    ReDim tmCff(0 To UBound(tgCffCntr)) As CFF
    For ilCff = 0 To UBound(tgCffCntr) - 1 Step 1
        tmCff(ilCff) = tgCffCntr(ilCff).CffRec
    Next ilCff
    
    For ilVef = 0 To UBound(tgMVef) Step 1
        tlVef(ilVef) = tgMVef(ilVef)
        If imTabIndex = 2 And tlVef(ilVef).sStdPrice = "A" Then
            tlVef(ilVef).sStdPrice = ""
        End If
    Next ilVef
    
    llRafCode = 0
    ilMnfSocEco = 0
    llOvStartTime = 0
    llOvEndTime = 0
    If cbcDemo.ListIndex < 0 Then
        ilMnfDemo = 0
    Else
        slNameCode = tgDemoCode(cbcDemo.ListIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfDemo = Val(slCode)
    End If
    gUnpackDateLong tgSpf.iReallDate(0), tgSpf.iReallDate(1), llReallDate

    For ilPkClf = 0 To UBound(tmClf) Step 1
        slPkLineType = tmClf(ilPkClf).sType
        If (slPkLineType = "O") Or (slPkLineType = "A") Then
            ilPkVef = gBinarySearchVef(tmClf(ilPkClf).iVefCode)
            ilHStartIndex = -1
            For ilHClf = 0 To UBound(tmClf) Step 1
                If tmClf(ilPkClf).iLine = tmClf(ilHClf).iPkLineNo Then
                    If ilHStartIndex = -1 Then
                        ilHStartIndex = ilHClf
                    End If
                    ilHEndIndex = ilHClf
                End If
            Next ilHClf
            
            
            'The following code taken from mChkForStdPkg: First step to distributing Package dollars to Hidden lines
            blPropPriceExist = False
            ilAnyRated = False
            
            For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                ilHVef = gBinarySearchVef(tmClf(ilHClf).iVefCode)
                For ilHCff = 0 To UBound(tmCff) Step 1
                    If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                        tmCff(ilHCff).sPriceType = "T"
                        tmCff(ilHCff).lPropPrice = gStrDecToLong(mGetPrice(tmClf(ilHClf).iVefCode, tmClf(ilHClf).iRdfCode, tmClf(ilHClf).iLen), 2)
                        If tmCff(ilHCff).lPropPrice > 0 Then
                            blPropPriceExist = True
                        End If
                        If tlVef(ilPkVef).sStdPrice = "P" Then   'Percent
                            'Fix  The percent obtain from PVF
                            tmCff(ilHCff).lActPrice = tmClf(ilHClf).lCode
                        ElseIf tlVef(ilPkVef).sStdPrice = "A" Then   'Audience
                            llAvgAud = 0
                            If ilHVef <> -1 Then
                                '4/4/06-Hidden line property not set until later, don't use GetLineType
                                slStr = "H" 'mGetLineType(imLnRowNo)
                                llRafCode = 0
                                gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                For ilDay = 0 To 6 Step 1
                                    If tmCff(ilHCff).iDay(ilDay) > 0 Then
                                        ilAudDays(ilDay) = True
                                    Else
                                        ilAudDays(ilDay) = False
                                    End If
                                Next ilDay

                                If tgSpf.iReallMnfDemo <= 0 Then
                                    If (tlVef(ilHVef).iReallDnfCode > 0) And (ilMnfDemo > 0) Then
                                        ilPassDnfCode = tlVef(ilHVef).iReallDnfCode
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmClf(ilHClf).iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmClf(ilHClf).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                    ElseIf (tlVef(ilHVef).iDnfCode > 0) And (ilMnfDemo > 0) And (llReallDate <= 0) Then
                                        ilPassDnfCode = tlVef(ilHVef).iDnfCode
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmClf(ilHClf).iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmClf(ilHClf).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), slStr, llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                    End If
                                Else
                                    If (tlVef(ilHVef).iReallDnfCode > 0) And (tgSpf.iReallMnfDemo > 0) Then
                                        ilPassDnfCode = tlVef(ilHVef).iReallDnfCode
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmClf(ilHClf).iVefCode, ilMnfSocEco, tgSpf.iReallMnfDemo, llDate, llDate, tmClf(ilHClf).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), slStr, llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                    End If
                                End If
                            End If
                            tmCff(ilHCff).lActPrice = llAvgAud
                        ElseIf tlVef(ilPkVef).sStdPrice = "S" Then   'Spot Count
                            'Test if daily or weeky
                            If tmCff(ilHCff).sDyWk = "D" Then
                                'Count spots for each day
                                tmCff(ilHCff).lActPrice = 0
                                For ilDay = 0 To 6 Step 1
                                    tmCff(ilHCff).lActPrice = tmCff(ilHCff).lActPrice + tmCff(ilHCff).iDay(ilDay)
                                Next ilDay
                            Else
                                tmCff(ilHCff).lActPrice = tmCff(ilHCff).iSpotsWk
                            End If
                        Else
                            'Set price temporary
                            tmCff(ilHCff).lActPrice = tmCff(ilHCff).lPropPrice 'tmCff(ilCff).lActPrice
                            llAPrice = llAPrice + tmCff(ilHCff).lPropPrice
                        End If
                        If ilHVef <> -1 Then
                            If tlVef(ilHVef).iMnfVehGp5Rsch > 0 Then
                                For ilLp = LBound(tmRatedGp) To UBound(tmRatedGp) - 1 Step 1
                                    If tlVef(ilHVef).iMnfVehGp5Rsch = tmRatedGp(ilLp).iMnfVehGp5Rsch Then
                                        If tmRatedGp(ilLp).sRated <> "Y" Then
                                            tmCff(ilHCff).lActPrice = -tmCff(ilHCff).lActPrice
                                        Else
                                            ilAnyRated = True
                                        End If
                                        Exit For
                                    End If
                                Next ilLp
                            End If
                        End If
                    End If
                Next ilHCff
            Next ilHClf
            
            
            'Step 2: Check if by rate and values zero
            If (llAPrice = 0) And (tlVef(ilPkVef).sStdPrice = "R") And (Not ilAnyRated) Then
                For ilPkCff = 0 To UBound(tmCff) Step 1
                    If tmClf(ilPkClf).iLine = tmCff(ilPkCff).iClfLine Then
                        If (tlVef(ilPkVef).sStdPrice <> "S") Or ((slPkLineType = "A") And (tlVef(ilPkVef).sStdPrice = "S")) Then
                            If tmCff(ilPkCff).sDyWk = "D" Then
                                ilSpots = 0
                                For ilDay = 0 To 6 Step 1
                                    ilSpots = ilSpots + tmCff(ilPkCff).iDay(ilDay)
                                Next ilDay
                            Else
                                ilSpots = tmCff(ilPkCff).iSpotsWk
                            End If
                            llPPrice = ilSpots * tmCff(ilPkCff).lActPrice
                        Else
                            llPPrice = tmCff(ilPkCff).lActPrice
                        End If
                        If llPPrice > 0 Then
                            gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                            For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                                For ilHCff = 0 To UBound(tmCff) Step 1
                                    If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                        gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                        If llStartDate = llDate Then
                                            tmCff(ilHCff).lActPrice = 1
                                            Exit For
                                        End If
                                     End If
                                Next ilHCff
                            Next ilHClf
                        End If
                    End If
                Next ilPkCff
            End If
            
            
            'Step 3: Distribute dollars
            For ilPkCff = 0 To UBound(tmCff) Step 1
                If tmClf(ilPkClf).iLine = tmCff(ilPkCff).iClfLine Then
                    llAPrice = 0
                    gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                    For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                        For ilHCff = 0 To UBound(tmCff) Step 1
                            If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                        
                                gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                If llStartDate = llDate Then
                                    If ilAnyRated Then
                                        If tmCff(ilHCff).lActPrice < 0 Then
                                            tmCff(ilHCff).lActPrice = 0
                                        End If
                                    Else
                                        If tmCff(ilHCff).lActPrice < 0 Then
                                            tmCff(ilHCff).lActPrice = -tmCff(ilHCff).lActPrice
                                        End If
                                    End If
                                    If (tlVef(ilPkVef).sStdPrice = "P") Or (tlVef(ilPkVef).sStdPrice = "S") Then
                                        llAPrice = llAPrice + tmCff(ilHCff).lActPrice
                                    Else
                                        If tmCff(ilHCff).sDyWk = "D" Then
                                            ilSpots = 0
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tmCff(ilHCff).iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tmCff(ilHCff).iSpotsWk
                                        End If
                                        llAPrice = llAPrice + ilSpots * tmCff(ilHCff).lActPrice
                                    End If
                                    Exit For
                                End If
                            End If
                        Next ilHCff
                    Next ilHClf
                    
                    
                    If (tlVef(ilPkVef).sStdPrice <> "S") Or ((slPkLineType = "A") And (tlVef(ilPkVef).sStdPrice = "S")) Or ((slPkLineType = "O") And (tlVef(ilPkVef).sStdPrice = "S")) Then
                        If tmCff(ilPkCff).sDyWk = "D" Then
                            llPPrice = 0
                            For ilDay = 0 To 6 Step 1
                                llPPrice = llPPrice + tmCff(ilPkCff).iDay(ilDay) * tmCff(ilPkCff).lActPrice
                            Next ilDay
                        Else
                            llPPrice = tmCff(ilPkCff).iSpotsWk * tmCff(ilPkCff).lActPrice
                        End If
                    Else
                        llPPrice = tmCff(ilPkCff).lActPrice
                    End If
                    'Distribute dollars
                    If llAPrice > 0 Then
                        ilSpotCount = 0
                        For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                            For ilHCff = 0 To UBound(tmCff) Step 1
                                If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                    gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        If tlVef(ilPkVef).sStdPrice <> "S" Then
                                            tmCff(ilHCff).lActPrice = (CSng(llPPrice) * tmCff(ilHCff).lActPrice) / llAPrice
                                            If tmCff(ilHCff).sDyWk = "D" Then
                                                ilSpots = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilSpots = ilSpots + tmCff(ilHCff).iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilSpots = tmCff(ilHCff).iSpotsWk
                                            End If
                                            If (ilSpots > 0) And (tlVef(ilPkVef).sStdPrice = "P") Then
                                                tmCff(ilHCff).lActPrice = tmCff(ilHCff).lActPrice / ilSpots
                                            End If
                                        Else
                                            If tmCff(ilHCff).lActPrice > 0 Then
                                                tmCff(ilHCff).lActPrice = CSng(llPPrice) / llAPrice
                                            Else
                                                tmCff(ilHCff).lActPrice = 0
                                            End If
                                            If tmCff(ilHCff).sDyWk = "D" Then
                                                For ilDay = 0 To 6 Step 1
                                                    ilSpotCount = ilSpotCount + tmCff(ilHCff).iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilSpotCount = ilSpotCount + tmCff(ilHCff).iSpotsWk
                                            End If
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next ilHCff
                        Next ilHClf
                        
                        gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                        For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                            For ilHCff = 0 To UBound(tmCff) Step 1
                                If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                    gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        If tmCff(ilHCff).sDyWk = "D" Then
                                            ilSpots = 0
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tmCff(ilHCff).iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tmCff(ilHCff).iSpotsWk
                                        End If
                                        llPPrice = llPPrice - ilSpots * tmCff(ilHCff).lActPrice
                                        Exit For
                                    End If
                                End If
                            Next ilHCff
                        Next ilHClf
                        
                        
                        
                        If llPPrice <> 0 Then
                            'If tgCffCntr(ilCff).CffRec.iSpotsWk = 1 Then
                            '    tgCffCntr(ilCff).CffRec.lActPrice = tgCffCntr(ilCff).CffRec.lActPrice + llPPrice
                            '    llPPrice = 0
                            'End If
                            gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                            For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                                If llPPrice = 0 Then
                                    Exit For
                                End If
                                For ilHCff = 0 To UBound(tmCff) Step 1
                                    If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                        gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                        If llStartDate = llDate Then
                                            If tmCff(ilHCff).sDyWk = "D" Then
                                                ilSpots = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilSpots = ilSpots + tmCff(ilHCff).iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilSpots = tmCff(ilHCff).iSpotsWk
                                            End If
                                            If tmCff(ilPkCff).sDyWk = "D" Then
                                                ilRet = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilRet = ilRet + tmCff(ilPkCff).iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilRet = tmCff(ilPkCff).iSpotsWk
                                            End If
                                            If (ilSpots = ilRet) And (tmCff(ilHCff).lActPrice + (llPPrice / ilSpots) >= 0) Then
                                                If (tmCff(ilHCff).lPropPrice > 0) Or (Not blPropPriceExist) Then
                                                    tmCff(ilHCff).lActPrice = tmCff(ilHCff).lActPrice + (llPPrice / ilSpots)
                                                    llPPrice = 0
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next ilHCff
                            Next ilHClf
                            
                            
                            '12/18/15: Test for 1 spot
                            If llPPrice <> 0 Then
                                gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                                For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                                    If llPPrice = 0 Then
                                        Exit For
                                    End If
                                    For ilHCff = 0 To UBound(tmCff) Step 1
                                        If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                            gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                            If llStartDate = llDate Then
                                                If tmCff(ilHCff).sDyWk = "D" Then
                                                    ilSpots = 0
                                                    For ilDay = 0 To 6 Step 1
                                                        ilSpots = ilSpots + tmCff(ilHCff).iDay(ilDay)
                                                    Next ilDay
                                                Else
                                                    ilSpots = tmCff(ilHCff).iSpotsWk
                                                End If
                                                If (ilSpots = 1) And (tmCff(ilHCff).lActPrice + (llPPrice / ilSpots) >= 0) Then
                                                    If (tmCff(ilHCff).lPropPrice > 0) Or (Not blPropPriceExist) Then
                                                        tmCff(ilHCff).lActPrice = tmCff(ilHCff).lActPrice + (llPPrice / ilSpots)
                                                        llPPrice = 0
                                                    End If
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    Next ilHCff
                                Next ilHClf
                            End If
                        End If
                    Else
                        tmCff(ilPkCff).lActPrice = 0
                        gUnpackDateLong tmCff(ilPkCff).iStartDate(0), tmCff(ilPkCff).iStartDate(1), llStartDate
                        For ilHClf = ilHStartIndex To ilHEndIndex Step 1
                            For ilHCff = 0 To UBound(tmCff) Step 1
                                If tmClf(ilHClf).iLine = tmCff(ilHCff).iClfLine Then
                                    gUnpackDateLong tmCff(ilHCff).iStartDate(0), tmCff(ilHCff).iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        tmCff(ilHCff).lActPrice = 0
                                        Exit For
                                    End If
                                End If
                            Next ilHCff
                        Next ilHClf
                    End If
                End If
            Next ilPkCff
        End If
    Next ilPkClf
    
    For ilClf = 0 To UBound(tmClf) - 1 Step 1
        tgClfCntr(ilClf).ClfRec = tmClf(ilClf)
    Next ilClf
    
    For ilCff = 0 To UBound(tmCff) - 1 Step 1
        tgCffCntr(ilCff).CffRec = tmCff(ilCff)
    Next ilCff
    
    mSetHiddenLinePrice = True
End Function

Private Sub mPopRateGp()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slCode As String
    
    smVehGp5CodeTag = ""
    ilRet = gPopMnfPlusFieldsBox(ProgrammaticBuy, lbcVefGp5, tmVehGp5Code(), smVehGp5CodeTag, "H5")
    Screen.MousePointer = vbHourglass
    ReDim tmRatedGp(LBound(tmVehGp5Code) To UBound(tmVehGp5Code)) As RATEDGP
    For ilLoop = LBound(tmVehGp5Code) To UBound(tmVehGp5Code) - 1 Step 1
        ilRet = gParseItem(tmVehGp5Code(ilLoop).sKey, 2, "\", slCode)
        tmRatedGp(ilLoop).iMnfVehGp5Rsch = Val(slCode)
        tmMnfSrchKey.iCode = tmRatedGp(ilLoop).iMnfVehGp5Rsch
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If StrComp(Trim$(tmMnf.sRPU), "N", 1) = 0 Then
                tmRatedGp(ilLoop).sRated = "N"
            Else
                tmRatedGp(ilLoop).sRated = "Y"
            End If
        Else
            tmRatedGp(ilLoop).sRated = "Y"
        End If
    Next ilLoop
End Sub


Private Function mInsertHdrComment() As Boolean

    Dim ilRet As Integer
    Dim llCxfCode As Long
    Dim slSQLQuery As String
    Dim max_rst As ADODB.Recordset

    mInsertHdrComment = False
    tmCxf = mDefaultCXF()
    Do
        slSQLQuery = "select max(cxfCode) from CXF_Hdr_Ln_Comments"
        Set max_rst = gSQLSelectCall(slSQLQuery)
        If IsNull(max_rst(0).Value) Then
            'error
            llCxfCode = llCxfCode
        Else
            llCxfCode = max_rst(0).Value + 1
            tmCxf.lCode = llCxfCode
            tmChf.lCxfCode = llCxfCode
            tmCxf.lAutoCode = llCxfCode
            tmCxf.sComment = Trim(edcComment.Text)
            ilRet = btrInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            'send failure to jeff
            ilRet = ilRet
        End If
    Loop While ilRet = BTRV_ERR_DUPLICATE_KEY
    mInsertHdrComment = True
    Exit Function

End Function

Private Sub mProgrammaticUser()

    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
        
    imUrfCode = 0
    slSQLQuery = "Select urfCode, urfPrgmmaticAlert From URF_User_Options Where urfPrgmmaticAlert = 'I' And urfCode > 2"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        imUrfCode = rst_Temp!urfCode
        smUrfPrgmmaticAlert = rst_Temp!urfPrgmmaticAlert
    End If
    
    rst_Temp.Close
End Sub
Private Sub mProgrammaticSalesPerson()

    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
        
    imSlfCode = 0
    slSQLQuery = "Select slfCode From SLF_Salespeople Where slfFirstName = 'Programmatic' and slfLastName = 'Buy'"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        imSlfCode = rst_Temp!slfCode
    End If
    
    rst_Temp.Close
End Sub


Function SendBrochure(sBrochure As String)
    
    Dim sPFN As String
    Dim lFileHandle As Long
    Dim lFSize As Long
    Dim iTotalBytesSent As Long
    Dim Buffer() As Byte
    Dim PacketSize As Long
    Dim bSendComplete As Boolean
    Dim BytesLeft As Long
    
    On Error GoTo ErrorHandler
    
    SendBrochure = "SUCCESS"
    
    'sPFN = CurDir & "\" & sBrochure
    sPFN = sgSalesBrochure & "\" & sBrochure
    lFSize = FileLen(sPFN)
    PacketSize = 4098
    
    If lFSize < PacketSize Then
        PacketSize = lFSize
    End If
    
    ReDim Buffer(1 To PacketSize) As Byte
    
    lFileHandle = FreeFile
    Open sPFN For Binary Access Read As #lFileHandle
    bSendComplete = False
    iTotalBytesSent = 0
    Do Until (iTotalBytesSent >= lFSize Or bSendComplete)
        
        BytesLeft = lFSize - Loc(lFileHandle)
        If iTotalBytesSent > 0 And BytesLeft < PacketSize Then
            If BytesLeft < 1 Then
                Close #lFileHandle
                Exit Function
            End If
            ReDim Buffer(1 To BytesLeft) As Byte
        End If
        
        Get #lFileHandle, , Buffer()
        sckClient.SendData Buffer()
        iTotalBytesSent = iTotalBytesSent + PacketSize
        DoEvents
    Loop
    Close #lFileHandle

    Exit Function
    
ErrorHandler:
    bSendComplete = True
    Resume Next
End Function

Private Function mGatherAgencyInfo() As String

    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    Dim slSQLQuery As String
    Dim ilCount As Integer
    Dim ilMaxRecs As Integer
    
    SQLQuery = "SELECT count(agfName) FROM AGF_Agencies"
    Set rst = gSQLSelectCall(SQLQuery)
    ilMaxRecs = rst(0).Value
    
    ilCount = 0
    smResponse = """" & "agfCode" & """, " & """" & "agfName" & """, " & """" & "agfAbbr" & """, " & """" & "agfCity" & """, " & """" & "agfCntrAddr1" & """, " & """" & "agfCntrAddr2" & """, " & """" & "agfCntrAddr3" & """" & vbCrLf
    SQLQuery = "SELECT agfName, agfAbbr, agfCode, agfCity, agfCntrAddr1, agfCntrAddr2, agfCntrAddr3"
    SQLQuery = SQLQuery + " FROM AGF_Agencies"
    SQLQuery = SQLQuery + " ORDER BY agfName, agfCity"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        ilCount = ilCount + 1
        If ilCount <> ilMaxRecs Then
            smResponse = smResponse & """" & rst!agfCode & """, " & """" & Trim(rst!agfName) & """, " & """" & Trim(rst!agfAbbr) & """, " & """" & Trim(rst!agfCity) & """, " & """" & Trim(rst!agfCntrAddr1) & """, " & """" & Trim(rst!agfCntrAddr2) & """," & """" & Trim(rst!agfCntrAddr3) & """" & vbCrLf
        Else
            smResponse = smResponse & """" & rst!agfCode & """, " & """" & Trim(rst!agfName) & """, " & """" & Trim(rst!agfAbbr) & """, " & """" & Trim(rst!agfCity) & """, " & """" & Trim(rst!agfCntrAddr1) & """, " & """" & Trim(rst!agfCntrAddr2) & """," & """" & Trim(rst!agfCntrAddr3) & """"
        End If
        rst.MoveNext
    Wend
End Function


Private Function mGetDataNoQuotes(sDataStr As String) As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slNewStr As String
    Dim clOneChar As String
    
    ilLen = Len(sDataStr)
    slNewStr = ""
    For ilLoop = 1 To ilLen
        clOneChar = Mid(sDataStr, ilLoop, 1)
        If clOneChar <> """" Then
            slNewStr = slNewStr + clOneChar
        End If
    Next
    mGetDataNoQuotes = Trim(slNewStr)
End Function
Private Sub cbcAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcProduct_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSignOff_Click()
    If bmShutdownNeeded Then
        Unload ProgrammaticBuy
    End If
End Sub

Private Sub cmcSignOff_GotFocus()
    mSetShow
End Sub

Private Sub mCFFListSet()
    
    Dim ilCff As Integer
    
    For ilCff = 0 To UBound(tgCffCntr) - 1 Step 1
        With tgCffCntr(ilCff)
            .lRecPos = 0
            .iStatus = 1  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
            .iNextCff = -1
            .lNextCff = -1    'Used to filter out lines/flights that have more then 32000 records
            gUnpackDateLong .CffRec.iStartDate(0), .CffRec.iStartDate(1), .lStartDate
            gUnpackDateLong .CffRec.iEndDate(0), .CffRec.iEndDate(1), .lEndDate
            'Need to find how to set the 3 below
            .lAvgAud = 0     'Average audience for user selected demo on demo bar
            .lPriDemoAvgAud = 0  'Average audience used for Contract Totals
                                'that show at bottom Right for primary demo
            .lPriDemoPop = 0
        
            .iGameNo = 0
        End With
    Next ilCff

End Sub

Private Sub mCLFListSet()

    Dim ilClf As Integer
    
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        With tgClfCntr(ilClf)
            .sKey = ""
            .lRecPos = 0
            .lCxfRecPos = 0
            .iStatus = 1  '0=New; 1=old and retain; 2=old and delete; -1=Not used
            .iFirstCff = ilClf
            .iFirstCgf = -1
            .lFirstCff = ilClf       'Used to filter out lines/flights that have more then 32000 records
            .lFirstCgf = -1       'Used to filter out lines/flights that have more then 32000 records
            .sGameLayout = ""
            .iCancel = False  'True=Canceled before start; False = not canceled
            .iOverride = False    'True=Allow subset of times to be defined
            .iGame = False    'True = Game buy; False = Non-Game buy
            .iLibBuy = False  'True= Library buy; False = Time buy
            .iInPast = False 'True=Line has flights in the past which are scheduled
            .lUnbilledDate = 0   'date of first unbilled spot in past
            .lEndBilledDate = 0  'Billing End Date
            .iResearch = True   'True=Research data is valid; False=Research data must be computed
            .iSPChg = False     'Line Spot or Price changed(Y= Don't show totals if using Compute Button)
            'need to get total cost from lines
            .lCost = 0       'Total Cost (used when computing Research)
            'need to set next 7 fields
            .lAvgAud = 0     'Avg Aud
            .iAvgRating = 0   'Average rating
            .lGrImp = 0    'Gross Impressions
            .lGRP = 0
            .lCPM = 0
            .lCPP = 0
            .lPop = 0   'Population
            'lRCCost As Long       'Rate Card Total Cost (used when computing Research)
            'iRCAvgRating As Integer   'Rate Card Average rating
            'lRCGrImp As Long      'Rate Card Gross Impressions
            'lRCGRP As Long        'Rate CardGRP
            'lRTCost As Long      'Request Total Cost
            'need to set total cost
            .lTSpots = 0 'Total number of spots
            .iSaveIndex = 0  'Save and Comment Index
            .iOrigAnySpots = 1   'Line have any spots originally(-1 = No Spots & no Price; 0=No Spots & Price defined, 1=Spots defined)
            .iLineSchd = False 'True= Line scheduled; False = New Line and not scheduled
            .sCurrentPrice = "" 'Current way price is to be set for a week (from Line price field)
            .iPriceSet = 2   '0=This flag not set; 1=Price not set; 2=Price set
            'need to set next 9 fields
            .lPriDemoPop = 0    'Population used for Contract Totals
                                    'that show at bottom Right for primary demo
            .iPriDemoBook = 0 'dnf reference used for Contract Totals
                                    'that show at bottom Right for primary demo
            .iPriChgd = 0     'If Vehicle; Daypart or Overrides changed
            .lPriDemoAvgAud = 0 'Primary Demo Average audience
            .iPriDemoAvgRating = 0 'Primary Demo Average Rating
            .lPriDemoGRP = 0
            .lPriDemoCPP = 0
            .lPriDemoCPM = 0
            .lPriDemoGrImp = 0
            .lLLD = 0
            .iMonThruToDaysSpots = 0 'Number of spots in past within current week (Monday thru Today)
            .iChgBilledPrice = 0 'Allow billed price to be changed
            '12/16/14: Add Hide/Show Hidden Lines
            .bShowLine = False
            .bPkExpanded = False
            .iAcqCostSpotRate = 0 '0=Indepentent; 1=Acquisition always zero; 2=Acquisition Cost must match the Spot Rate
            .sOwnership = ""    'A=Owned-Network; B=Owned-Station: C=Unowned-Network; D=Unowned-Station.  Defaul A or Blank. Test for B, C and D
                                        'Proposal/order acquisition rules
                                        'If Network(Owned or unowned): Use current acquisition rules
                                        'If Station-Owned: Bypass Acquisition cost and set to zero
                                        'If Station-Unowned: Acquisition Cost and Spot rate will match
                                        'Proposal/Order package overrides copy to hidden line rules
                                        'If Network(Owned or Unowned): Retain current rules (don't violated the hidden times/days)
                                        'If Station (Owned or Unowned): Copy package overrides to hidden lines (allow hidden line times/days to be violated)
            .sLineChgd = ""   'Line changed (Vehicle, Daypart, Overrides, Length, Spots, Acquisition, Spot rate)
        End With
    Next ilClf

End Sub
Private Sub mLoadExportFacts()

'ilRet = mXmlDemoInfo(.hmDrf, .hmMnf, .hmDpf, .hmDef, .hmRaf, .tmRegionCode(), .imStep1Index)

   With tgExportFacts
       .hmDef = hmDef
       .hmDpf = hmDpf
       .hmDrf = hmDrf
       .hmMnf = hmMnf
       .hmRaf = hmRaf
       .hmCHF = hmCHF
       .hmClf = hmClf
       .imChgMode = 0 'imChgMode
       '.imCppCpm = imCppCpm
       '.imCPriDemoRating = imCPriDemoRating
       .imMnfRecLen = imMnfRecLen
        .imStep1Index = 0 'imStep1Index  in the proposal dropdown in step 1   0=create from scratch
       .imTerminate = 0 'imTerminate
       '.imVefCode = imCntrVefCode
       '.imVpfIndex = imVpfIndex
       '.lmCPriDemoCPM = lmCPriDemoCPM
       '.lmCPriDemoCPP = lmCPriDemoCPP
       '.lmCPriDemoGrImp = lmCPriDemoGrImp
       .tmMnf = tmMnf
       .tmMnfSrchKey = tmMnfSrchKey
       '.tmRegionCode = tmRegionCode
       .smAgencyXMLPath = sgExportPath
       '.smStationXMLWOPath = smStationXMLWOPath
       .imLB1Or2 = 1 'imLB1Or2
       '.smLnSave = smLnSave
       .imDiffMode = 0 'imDiffMode
       '.smStationXMLBasePath = smStationXMLBasePath
   End With
   'tgExportAgf = tmAgf
   'tgExportAgf = tgCommAgf
End Sub

Private Function mGenPDF() As Boolean

    Dim ilLoop As Integer
    Dim slStr As String

    Dim ilRet As Integer
    Dim llCurrCode As Long
    Dim llPrevCode As Long
    Dim llPrevSpots As Long
    Dim llPrevGross As Long
    Dim ilCurrTotalWks As Integer
    Dim ilCurrAirWks As Integer
    Dim slDate As String            'generation date
    Dim slTime As String            'generation time
    Dim slDetSumBoth As String      '0=det, 1 = sum, 2 = both
    Dim slInclRates As String       'Show rates on printed contract Y/N
    Dim slInclResearch As String    'show research data on printed contracts Y/N
    Dim slInclProof As String       'show hidden lines on printed contracts Y/N
    Dim ilChfRecLen As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slOutputTo As String * 1
    Dim slSaveToFileName As String
    Dim slExportIndex As String
    Dim slInclSplits As String      '2-14-04 show slsp comm splits
    Dim slInclNTRBillSummary As String  '2-2-10 show NTR bill summary with air time summary
    Dim slShowNetOnProps As String      '2-3-10
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilErr As Integer
    Dim tlSbfList() As SBF
    Dim llStdStartDates(0 To 37) As Long            'standard start dates for bdcst calendar, 4 years, Index 0 ignored
    Dim slShowProdProt As String
    Dim tlBR As BRSELECTIONS        'BR options
    Dim ilGenTime(0 To 1) As Integer    '10-30-01
    Dim slAdvName As String
    Dim slProduct As String
    Dim slToFile As String
    mGenPDF = True
    tlBR.sSnapshot = " "     'flag to indicate snapshot rather than from reports
    tlBR.iShowProof = False     'force to not show Proof mode
    tlBR.iCorpOrStd = True      'forc standard monthly summaries
    tlBR.iPrintables = False
    tlBR.iDiffOnly = False
    tlBR.iThisCntMod = False
    tlBR.iShowSplits = True     'always assume to show the splits on summaries
    tlBR.iShowNTRBillSummary = True 'assume to show ntr monthly billing with air time
    tlBR.iShowNetAmtOnProps = True  'assume to show net amts on proposals
    tlBR.iShowProdProt = False          '8-25-15 default to exclude showing Prod Protection categories
    tlBR.iShowRates = True
    If imTabIndex <> 2 Then
        tlBR.iShowResearch = True
    Else
        tlBR.iShowResearch = False
    End If
    slInclProof = "N"
    tlBR.iDetail = True
    tlBR.iSummary = True
    tlBR.iAllDemos = False
    tlBR.iShowSplits = False
    tlBR.iSocEcoMnfCode = 0
    tlBR.iWhichSort = 0         'force to sort by advt, contr#

    llCurrCode = tgChfCntr.lCode
    llPrevCode = 0
    llPrevSpots = 0
    llPrevGross = 0
    ilCurrTotalWks = 0
    ilCurrAirWks = 0
    tlBR.iPropOrOrder = False               'assume order (HOGN)
    If tgChfCntr.sStatus = "W" Or tgChfCntr.sStatus = "D" Or tgChfCntr.sStatus = "C" Or tgChfCntr.sStatus = "I" Then
        tlBR.iPropOrOrder = True
    End If
    'ProcessFlag 0 = open files, then process cnt, 1 = close files and exit
    tlBR.lPrevSpots = llPrevSpots
    tlBR.lPrevGross = llPrevGross
    tlBR.iCurrTotWks = ilCurrTotalWks
    tlBR.iCurrAirWks = ilCurrAirWks
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tlBR.iGenDate(0), tlBR.iGenDate(1)
    igNowDate(0) = tlBR.iGenDate(0)
    igNowDate(1) = tlBR.iGenDate(1)
    slTime = Format$(gNow(), "h:mm:ssAM/PM")
    gPackTime slTime, ilGenTime(0), ilGenTime(1)    '10-30-01
    gUnpackTimeLong ilGenTime(0), ilGenTime(1), False, lgNowTime
    ilRet = gProcessBR(tlBR, 0, CONTRACTSJOB)             'open files
    If ilRet <> 0 Then
        ilRet = gProcessBR(tlBR, 2, CONTRACTSJOB)       'close files and terminate
            Screen.MousePointer = vbDefault
            'mTerminate
            mGenPDF = False
        Exit Function
    End If
    ilRet = gProcessBR(tlBR, 1, CONTRACTSJOB)            'process one contract
    ilRet = gProcessBR(tlBR, 2, CONTRACTSJOB)            'close files
    If igTestSystem Then
        slStr = "Traffic^Test^NOHELP\" & sgUserName
    Else
        slStr = "Traffic^Prod^NOHELP\" & sgUserName
    End If
    'Set up the command line string to pass to print the contract
    If tlBR.iDetail And tlBR.iSummary Then
        slDetSumBoth = "2"
    ElseIf tlBR.iDetail And Not tlBR.iSummary Then
        slDetSumBoth = "0"              'detail only
    Else
        slDetSumBoth = "1"              'summary only
    End If
    If tlBR.iShowRates Then
        slInclRates = "Y"
    Else
        slInclRates = "N"
    End If
    If tlBR.iShowResearch Then
        slInclResearch = "Y"
    Else
        slInclResearch = "N"
    End If
    If tlBR.iShowProof Then
        slInclProof = "Y"
    Else
        slInclProof = "N"
    End If

    If tlBR.iShowSplits Then     '2-14-04  show slsp splits on summary
        slInclSplits = "Y"
    Else
        slInclSplits = "N"
    End If
    
    If tlBR.iShowNTRBillSummary Then     '2-2-10 show ntr bill summary with air time
        slInclNTRBillSummary = "Y"
    Else
        slInclNTRBillSummary = "N"
    End If
    
    If tlBR.iShowNetAmtOnProps Then     '2-3-10 show net amt on proposal?
        slShowNetOnProps = "Y"
    Else
        slShowNetOnProps = "N"
    End If
    
    If tlBR.iShowProdProt Then     '8-25-15 Show prod protection categories
        slShowProdProt = "Y"
    Else
        slShowProdProt = "N"
    End If
    slOutputTo = "2"
    slExportIndex = "0"
    slToFile = tgChf.lCntrNo & "-P" & tgChf.iPropVer
    slAdvName = grdHeader(imTokenIdx).TextMatrix(1, ADVERTISERINDEX)
    slProduct = Trim$(grdHeader(imTokenIdx).TextMatrix(1, PRODUCTINDEX))
    smFileName = Trim$(sgExportPath) & slToFile & "-" & gFileNameFilter(slAdvName) & "-" & gFileNameFilter(slProduct) & "-" & gFileNameFilter(Trim$(sgClientName))
    slToFile = slToFile & "-" & gFileNameFilter(slAdvName) & "-" & gFileNameFilter(slProduct) & "-" & gFileNameFilter(Trim$(sgClientName)) & ".PDF"
    slSaveToFileName = Trim$(sgExportPath & slToFile)      'save to location
    'parm 1:  Application name^Test or Prod
    'parm 2:  user name
    'parm 3:  Generation date
    'parm 4:  Generation time
    'parm 5:  Detail(0), Summary(1), Both (2)
    'parm 6: Includes Rates Y/N
    'Parm 7: include Research Y/n
    'Parm 7A: include sls splits (y/n) 2-14-04
    'Parm 7B: Include NTR bill summary with air time summary 2-2-10
    'Parm 7C: Show Net Amts on Proposals  2-3-10
    'Parm 7D: show prod prot codes 8-25-15
    'Parm 8: include Proof
    'Parm 9: Display(0) or Print(1) (always 1)
    'Parm 10: File export index
    'Parm 11: filename for export
    slStr = slStr & "\" & slDate & "\" & slTime & "\" & slDetSumBoth & "\" & slInclRates & "\" & slInclResearch & "\" & slInclSplits & "\" & slInclNTRBillSummary & "\" & slShowNetOnProps & "\" & slShowProdProt & "\" & slInclProof & "\" & slOutputTo & "\" & slExportIndex & "\|" & slSaveToFileName
    sgCommandStr = slStr
    RptSelBR.mInitReport
    mSendAgencyEmail
    Unload RptSelBR
End Function

Private Sub mSendAgencyEmail()

    Dim slSubject As String
    Dim slFileNamePDF As String
    Dim slFileNameXML As String
    Dim slAttachment As String
    
    
    If slSubject = "" Then
        slSubject = "Automated message from a client"
    End If
    Set ogEmailer = New CEmail
    
    slFileNamePDF = smFileName & ".pdf"
    slFileNameXML = smFileName & ".xml"
    slAttachment = slFileNamePDF & ";" & slFileNameXML
    
    With ogEmailer
        .FromAddress = "AClient@Counterpoint.net"
        .FromName = Trim$(sgClientName)
        .AddTOAddress Trim(tgTokenInfo(imTokenIdx).sEMail), Trim(tgTokenInfo(imTokenIdx).sBuyer)
        .Subject = "Programmatic Buy Proposal"
        .Message = "This is an automated email.  Please do not respond."
        .Attachment = slAttachment
        .SetHost "smtpauth.hosting.earthlink.net", 587, "emailsender@counterpoint.net", "Csi44Sic", True
        '.SetHost "smtpauth.hosting.earthlink.net", 25, "172.16.11.10", "Csi44Sic", True
        If Not .Send() Then
            gLogMsg "Email could not be sent from frmProgrammaticBuy-mSendAgencyEmail: ", "ProgrammaticBuy.Txt", False
        End If
    End With
    Set ogEmailer = Nothing
End Sub

Private Sub mDistributeSpotsByWk()
    Dim ilTotalSpots As Integer
    Dim ilNoWks As Integer
    Dim ilNoPerWk As Integer
    Dim ilNoExtraWks As Integer
    Dim ilCffNext As Integer
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilCffPrevNext As Integer
    Dim ilCffUpper As Integer
    Dim ilWk As Integer
    Dim ilClf As Integer
    Dim ilPkg As Integer
    Dim ilMultFactor As Integer
    Dim blFirstHidden As Boolean
    Dim llActPrice As Long
    Dim llPropPrice As Long
    ReDim ilDays(0 To 6) As Integer
    Dim ilSpotCount As Integer
    
    'If multi-weeks split spots across weeks
    'Only weekly buy defined with imTabIndex 0 and 1.
    ilNoWks = Val(grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, NUMBERWEEKSINDEX))
    If ilNoWks <= 1 Then
'        Exit Sub
    End If
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        If tgClfCntr(ilClf).ClfRec.sType <> "H" Then
            ilNoWks = Val(grdHeader(imTokenIdx).TextMatrix(grdHeader(imTokenIdx).FixedRows, NUMBERWEEKSINDEX))
            ilCffNext = tgClfCntr(ilClf).iFirstCff
            ilTotalSpots = tgCffCntr(ilCffNext).CffRec.iSpotsWk
            'The divide \ will truncate the division
            ilNoPerWk = ilTotalSpots \ ilNoWks
            ilNoExtraWks = ilTotalSpots Mod ilNoWks
            tgCffCntr(ilCffNext).CffRec.iSpotsWk = ilNoPerWk
            If ilNoExtraWks > 0 Then
                ilCffPrevNext = ilCffNext
                gUnpackDateLong tgCffCntr(ilCffNext).CffRec.iStartDate(0), tgCffCntr(ilCffNext).CffRec.iStartDate(1), llStartDate
                llEndDate = llStartDate + 6
                gPackDateLong llEndDate, tgCffCntr(ilCffNext).CffRec.iEndDate(0), tgCffCntr(ilCffNext).CffRec.iEndDate(1)
                tgCffCntr(ilCffNext).lStartDate = llStartDate
                tgCffCntr(ilCffNext).lEndDate = llEndDate
                If ilNoPerWk = 0 Then
                    ilNoWks = ilNoExtraWks
                    ilNoExtraWks = 0
                    ilNoPerWk = 1
                    tgCffCntr(ilCffNext).CffRec.iSpotsWk = ilNoPerWk
                End If
                For ilWk = 2 To ilNoWks Step 1
                    ilCffUpper = UBound(tgCffCntr)
                    tgCffCntr(ilCffUpper) = tgCffCntr(ilCffNext)
                    tgCffCntr(ilCffPrevNext).iNextCff = ilCffUpper
                    tgCffCntr(ilCffPrevNext).lNextCff = ilCffUpper
                    tgCffCntr(ilCffUpper).iNextCff = -1
                    tgCffCntr(ilCffUpper).lNextCff = -1
                    llStartDate = llStartDate + 7
                    llEndDate = llStartDate + 6
                    gPackDateLong llStartDate, tgCffCntr(ilCffUpper).CffRec.iStartDate(0), tgCffCntr(ilCffUpper).CffRec.iStartDate(1)
                    gPackDateLong llEndDate, tgCffCntr(ilCffUpper).CffRec.iEndDate(0), tgCffCntr(ilCffUpper).CffRec.iEndDate(1)
                    tgCffCntr(ilCffUpper).lStartDate = llStartDate
                    tgCffCntr(ilCffUpper).lEndDate = llEndDate
                    ReDim Preserve tgCffCntr(0 To ilCffUpper + 1) As CFFLIST
                    ilCffPrevNext = ilCffUpper
                Next ilWk
                'Create extra flights
                Do While ilNoExtraWks > 0
                    tgCffCntr(ilCffNext).CffRec.iSpotsWk = tgCffCntr(ilCffNext).CffRec.iSpotsWk + 1
                    ilNoExtraWks = ilNoExtraWks - 1
                    ilCffNext = tgCffCntr(ilCffNext).iNextCff
                    If ilCffNext = -1 Then
                        ilCffNext = tgClfCntr(ilClf).iFirstCff
                    End If
                Loop
            End If
        Else
            'Find the package line associated with the hidden lines
            'Still need Daily
            ilMultFactor = tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.iSpotsWk
            llActPrice = tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.lActPrice
            llPropPrice = tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.lPropPrice
            'Critical point
            'llActPrice = llPropPrice
            For ilDay = 0 To 6 Step 1
                ilDays(ilDay) = tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.iDay(ilDay)
            Next ilDay
            For ilPkg = 0 To ilClf - 1 Step 1
                If tgClfCntr(ilPkg).ClfRec.iLine = tgClfCntr(ilClf).ClfRec.iPkLineNo Then
                    ilCffNext = tgClfCntr(ilPkg).iFirstCff
                    Do While ilCffNext <> -1
                        If ilCffNext = tgClfCntr(ilPkg).iFirstCff Then
                            ilCffUpper = tgClfCntr(ilClf).iFirstCff
                            blFirstHidden = True
                        Else
                            ilCffUpper = UBound(tgCffCntr)
                            blFirstHidden = False
                        End If
                        tgCffCntr(ilCffUpper) = tgCffCntr(ilCffNext)
                        tgCffCntr(ilCffUpper).CffRec.iClfLine = tgClfCntr(ilClf).ClfRec.iLine
                        tgCffCntr(ilCffUpper).CffRec.iSpotsWk = ilMultFactor * tgCffCntr(ilCffNext).CffRec.iSpotsWk
                        If tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.sDyWk <> "D" Then
                            ilSpotCount = tgCffCntr(ilCffUpper).CffRec.iSpotsWk
                        Else
                            ilSpotCount = 0
                        End If
                        'moved to use the daily or weekly spot count
                        'tgCffCntr(ilCffUpper).CffRec.lActPrice = llActPrice / tgCffCntr(ilCffUpper).CffRec.iSpotsWk
                        'tgCffCntr(ilCffUpper).CffRec.lPropPrice = llPropPrice
                        For ilDay = 0 To 6 Step 1
                            If tgCffCntr(ilCffUpper).CffRec.sDyWk <> "D" Then
                                tgCffCntr(ilCffUpper).CffRec.iDay(ilDay) = ilDays(ilDay)
                            Else
                                If ilDays(ilDay) <> 0 Then
                                    tgCffCntr(ilCffUpper).CffRec.iDay(ilDay) = ilMultFactor * tgCffCntr(ilCffNext).CffRec.iDay(ilDay)
                                    ilSpotCount = ilSpotCount + tgCffCntr(ilCffUpper).CffRec.iDay(ilDay)
                                Else
                                    tgCffCntr(ilCffUpper).CffRec.iDay(ilDay) = 0
                                End If
                            End If
                        Next ilDay
                        tgCffCntr(ilCffUpper).CffRec.lActPrice = llActPrice / ilSpotCount
                        tgCffCntr(ilCffUpper).CffRec.lPropPrice = llPropPrice
                        If tgCffCntr(tgClfCntr(ilClf).iFirstCff).CffRec.sDyWk = "D" Then
                            tgCffCntr(ilCffUpper).CffRec.iSpotsWk = 0
                        End If
                        tgCffCntr(ilCffUpper).lNextCff = -1
                        If Not blFirstHidden Then
                            tgCffCntr(ilCffPrevNext).iNextCff = ilCffUpper
                            tgCffCntr(ilCffPrevNext).lNextCff = ilCffUpper
                            ReDim Preserve tgCffCntr(0 To ilCffUpper + 1) As CFFLIST
                        End If
                        ilCffPrevNext = ilCffUpper
                        ilCffNext = tgCffCntr(ilCffNext).iNextCff
                    Loop
                    Exit For
                End If
            Next ilPkg
        End If
    Next ilClf
    
End Sub

Private Sub mSignOff()
    
    Dim ilIdx As Integer
    
    For ilIdx = LBound(tgTokenInfo) To UBound(tgTokenInfo) Step 1
        If Trim(tgTokenInfo(ilIdx).sCommToken) = Trim(tgTokenInfo(imTokenIdx).sCommToken) Then
            tgTokenInfo(ilIdx).iAgfCode = 0
            tgTokenInfo(ilIdx).iAgfCode = 0
            tgTokenInfo(ilIdx).iTokenIdx = 0
            tgTokenInfo(ilIdx).sAgency = ""
            tgTokenInfo(ilIdx).sBuyer = ""
            tgTokenInfo(ilIdx).sCommToken = ""
            tgTokenInfo(ilIdx).sDate = ""
            tgTokenInfo(ilIdx).sInUseStatus = "N"
            tgTokenInfo(ilIdx).sTime = ""
            Exit For
        End If
    Next ilIdx
    cmcUndo_Click
End Sub

Private Sub mSetCurrGridVisable()

    Dim ilIdx As Integer
    
    For ilIdx = 0 To UBound(tgTokenInfo) - 1
        If ilIdx <> imTokenIdx Then
            grdHeader(ilIdx).Visible = False
            grdLines(ilIdx).Visible = False
        End If
    Next ilIdx
    grdHeader(imTokenIdx).Visible = True
    grdLines(imTokenIdx).Visible = True
End Sub



Private Function mAdvertiserExists(sName As String) As Boolean

    Dim adf_rst As ADODB.Recordset
    Dim slSQLQuery As String
    
    
    mAdvertiserExists = False
    slSQLQuery = "Select adfCode from ADF_Advertisers where adfName = " & "'" & gFixQuote(Trim(sName)) & "'"
    Set adf_rst = cnn.Execute(slSQLQuery)
    If adf_rst.EOF Then
        Exit Function
    End If
    lgAdfCode = adf_rst!adfCode
    mAdvertiserExists = True
                    
End Function


Private Function mProductExists(sName As String) As Boolean

    Dim prf_rst As ADODB.Recordset
    Dim slSQLQuery As String
    
    mProductExists = False
    imMnfComp(0) = 0
    imMnfExcl(0) = 0
    imMnfComp(1) = 0
    imMnfExcl(1) = 0
    lgPrfCode = 0
    slSQLQuery = "Select * from PRF_Product_Names where prfadfCode = " & lgAdfCode & " And UCase(prfName) = " & " '" & UCase(Trim(sName) & "'")
    Set prf_rst = cnn.Execute(slSQLQuery)
    If prf_rst.EOF Then
        Exit Function
    End If
    lgPrfCode = prf_rst!prfCode
    imMnfComp(0) = prf_rst!prfMnfComp1
    imMnfComp(1) = prf_rst!prfMnfComp2
    imMnfExcl(0) = prf_rst!prfMnfExcl1
    imMnfExcl(1) = prf_rst!prfMnfExcl2
    mProductExists = True
    
End Function

Private Function AdvNew() As Boolean

    Dim slSQLQuery As String
    Dim llPrfCode As Long
    
    AdvNew = False
    sgAdfname = Trim(smHeaderAdvt)
    sgAdfAbbrv = Mid$(Trim(smHeaderAdvt), 1, 7)
    sgProdName = Trim(smHeaderProd)
    slSQLQuery = mDefaultADF()
    lgAdfCode = gInsertAndReturnCode(slSQLQuery, "ADF_Advertisers", "AdfCode", "Replace")
    
    AdvNew = True
End Function
Private Function ProdNew() As Boolean

    Dim slSQLQuery As String
    Dim llPrfCode As Long
    Dim ilAdfCode
    
    ProdNew = False
    If Len(smHeaderProd) > 0 Then
        sgProdName = smHeaderProd
        slSQLQuery = mDefaultPRF()
        llPrfCode = gInsertAndReturnCode(slSQLQuery, "PRF_Product_Names", "prfCode", "Replace")
        slSQLQuery = "UPDATE PRF_Product_Names SET "
        slSQLQuery = slSQLQuery + "prfAutoCode = " & llPrfCode
        slSQLQuery = slSQLQuery + " WHERE (prfCode = " & llPrfCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "ProgrammaticBuy.Txt", "modAgmnt-mAbfInsert"
        End If
    End If
    ProdNew = True
    
End Function


Private Function AdvProdUpdate() As Boolean

    Dim slSQLQuery As String
    
    AdvProdUpdate = False
    lgAdfCode = mGetAdfCode(smHeaderAdvt)
    sgProdName = Trim(smHeaderProd)
    If Len(sgProdName) > 0 Then
        slSQLQuery = "UPDATE ADF_Advertisers SET "
        slSQLQuery = slSQLQuery & "adfProd = " & "'" & sgProdName & "',"
        slSQLQuery = slSQLQuery & "adfagfCode = " & "'" & igAgfCode & "',"
        slSQLQuery = slSQLQuery & "adfSlfCode= " & "'" & igSlfCode & "'"
        slSQLQuery = slSQLQuery + " WHERE adfCode = " & lgAdfCode
        If Not gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            slSQLQuery = slSQLQuery
        End If
    End If
    AdvProdUpdate = True
    
End Function


Private Sub mRemDormantVeh()
    
    ReDim tlVef(0 To UBound(tgMVef)) As VEF
    Dim ilIdx As Integer
    Dim ilIdx2 As Integer
    Dim ilCount As Integer
    
    ilCount = 0
    For ilIdx = 0 To UBound(tgMVef) Step 1
        tlVef(ilIdx) = tgMVef(ilIdx)
        If tlVef(ilIdx).sState <> "D" Then
            ilCount = ilCount + 1
        End If
    Next ilIdx
    ReDim tgMVef(0 To ilCount) As VEF
    
    ilIdx2 = 0
    For ilIdx = 0 To UBound(tlVef) Step 1
        If tlVef(ilIdx).sState <> "D" Then
            tgMVef(ilIdx2) = tlVef(ilIdx)
            ilIdx2 = ilIdx2 + 1
        End If
    Next ilIdx
End Sub
