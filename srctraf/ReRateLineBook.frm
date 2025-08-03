VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ReRateLineBook 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   13095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ReRateLineBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbcTooltip 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   560
      Left            =   9480
      ScaleHeight     =   555
      ScaleWidth      =   5055
      TabIndex        =   57
      Top             =   5160
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lbcTooltip 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vehicle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10
         TabIndex        =   61
         Top             =   280
         Width           =   5025
      End
      Begin VB.Label lbcTooltip 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Daypart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2600
         TabIndex        =   60
         Top             =   15
         Width           =   2445
      End
      Begin VB.Label lbcTooltip 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1610
         TabIndex        =   59
         Top             =   10
         Width           =   975
      End
      Begin VB.Label lbcTooltip 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Contract"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10
         TabIndex        =   58
         Top             =   15
         Width           =   1575
      End
   End
   Begin VB.PictureBox pbcDropShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   560
      Left            =   9720
      ScaleHeight     =   555
      ScaleWidth      =   5055
      TabIndex        =   62
      Top             =   5280
      Visible         =   0   'False
      Width           =   5055
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcPurchaseBook 
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   450
      BorderStyle     =   1
   End
   Begin V81TrafficReports.CSI_ComboBoxMS cbcBook 
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   6240
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   450
      BorderStyle     =   1
   End
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "ReRateLineBook.frx":08CA
      Left            =   0
      List            =   "ReRateLineBook.frx":08CC
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   5070
   End
   Begin V81TrafficReports.CSI_Calendar edcGRBEnd 
      Height          =   330
      Left            =   4200
      TabIndex        =   56
      Top             =   6840
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      Text            =   "2/9/2022"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81TrafficReports.CSI_Calendar edcGRBStart 
      Height          =   330
      Left            =   2040
      TabIndex        =   55
      Top             =   6840
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      Text            =   "2/9/2022"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   80
      Picture         =   "ReRateLineBook.frx":08CE
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcAssign 
      Caption         =   "Assign"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10440
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcMap 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IntegralHeight  =   0   'False
      ItemData        =   "ReRateLineBook.frx":0BD8
      Left            =   7800
      List            =   "ReRateLineBook.frx":0BDA
      TabIndex        =   48
      Top             =   6600
      Width           =   2520
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcAssignBookName 
      Height          =   330
      Left            =   1680
      TabIndex        =   24
      Top             =   570
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   582
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin V81TrafficReports.CSI_Calendar edcGBEnd 
      Height          =   330
      Left            =   3840
      TabIndex        =   20
      Top             =   1100
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      Text            =   "2/9/2022"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81TrafficReports.CSI_Calendar edcGBStart 
      Height          =   330
      Left            =   1695
      TabIndex        =   21
      Top             =   1100
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      Text            =   "2/9/2022"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcContract 
      Height          =   330
      Left            =   7800
      TabIndex        =   35
      Top             =   5880
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcVehicle 
      Height          =   330
      Left            =   7800
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   582
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.Frame frcByBook 
      BorderStyle     =   0  'None
      Height          =   1340
      Left            =   480
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   7200
      Begin VB.Frame frcFilterReRateBook 
         Caption         =   "Filter ReRate Book Names:"
         Height          =   640
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   7200
         Begin VB.CommandButton cmcGetReRateBook 
            Appearance      =   0  'Flat
            Caption         =   "Apply Filter"
            Height          =   285
            Left            =   5760
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   " to"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3000
            TabIndex        =   53
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label1 
            Caption         =   "Book Dates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   52
            Top             =   280
            Width           =   1290
         End
      End
      Begin VB.CommandButton cmcFilterReRateBooks 
         Caption         =   "Filter Books.."
         Height          =   285
         Left            =   5880
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.PictureBox pbcExcludeToInclude 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   180
         Left            =   6150
         Picture         =   "ReRateLineBook.frx":0BDC
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   380
         Width           =   180
      End
      Begin VB.CommandButton cmcFromMap 
         Appearance      =   0  'Flat
         Caption         =   "    Mo&ve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5910
         TabIndex        =   27
         Top             =   330
         Width           =   1185
      End
      Begin VB.PictureBox pbcIncludeToExclude 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   180
         Left            =   6720
         Picture         =   "ReRateLineBook.frx":0CB6
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   80
         Width           =   180
      End
      Begin VB.CommandButton cmcToMap 
         Appearance      =   0  'Flat
         Caption         =   "M&ove   "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5910
         TabIndex        =   25
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lacReRateBook 
         Caption         =   "ReRate Book"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lacPurchaseBook 
         Caption         =   "Purchase Book"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.ComboBox cbcGBVehicle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "ReRateLineBook.frx":0D90
      Left            =   1680
      List            =   "ReRateLineBook.frx":0D92
      TabIndex        =   19
      Top             =   1520
      Width           =   4365
   End
   Begin VB.ListBox lbcPurchaseBook 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "ReRateLineBook.frx":0D94
      Left            =   10680
      List            =   "ReRateLineBook.frx":0D96
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox lbcBook 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "ReRateLineBook.frx":0D98
      Left            =   11025
      List            =   "ReRateLineBook.frx":0D9A
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox lbcVehicle 
      Height          =   255
      ItemData        =   "ReRateLineBook.frx":0D9C
      Left            =   9480
      List            =   "ReRateLineBook.frx":0D9E
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   5865
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmcClear 
      Caption         =   "Clear Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7875
      TabIndex        =   3
      Top             =   5190
      Width           =   1335
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -30
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   4860
      Width           =   45
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   5190
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5190
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBookGrid 
      Height          =   2250
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2055
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   3969
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcLnBookName 
      Height          =   165
      Left            =   9525
      TabIndex        =   5
      Top             =   5655
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   291
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   0
   End
   Begin VB.ListBox lbcApplyLog 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "ReRateLineBook.frx":0DA0
      Left            =   120
      List            =   "ReRateLineBook.frx":0DA2
      TabIndex        =   14
      Top             =   4320
      Width           =   12855
   End
   Begin VB.Frame frcOptions 
      Caption         =   "Research Books by Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmcFilterBooks 
         Caption         =   "Filter Books.."
         Height          =   285
         Left            =   6000
         TabIndex        =   49
         Top             =   540
         Width           =   1215
      End
      Begin VB.PictureBox pbcView 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox ckcDontOverwriteByLine 
         Caption         =   "Retain previous assignments"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7440
         TabIndex        =   29
         Top             =   1580
         Width           =   2175
      End
      Begin VB.CommandButton cmcApply 
         Caption         =   "Assign"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11480
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Frame frcByLineOpts 
         Height          =   1935
         Left            =   7320
         TabIndex        =   31
         Top             =   0
         Width           =   5535
         Begin VB.OptionButton rbcLine 
            Caption         =   "MG/Bonus only"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   1020
            TabIndex        =   39
            Top             =   525
            Width           =   1320
         End
         Begin VB.OptionButton rbcLine 
            Caption         =   "All except MG/Bonus"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1020
            TabIndex        =   38
            Top             =   810
            Width           =   1680
         End
         Begin VB.OptionButton rbcLine 
            Caption         =   "All Lines"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1020
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.TextBox edcLines 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1600
            TabIndex        =   36
            Top             =   600
            Visible         =   0   'False
            Width           =   3795
         End
         Begin VB.Label lacLines 
            Caption         =   "Lines (3-9,12,14)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lbcApplyOption 
            Caption         =   "Lines"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame frcFilter 
         Caption         =   "Filter Book Names:"
         Height          =   1080
         Left            =   120
         TabIndex        =   40
         Top             =   855
         Visible         =   0   'False
         Width           =   7335
         Begin VB.CommandButton cmcGetBook 
            Appearance      =   0  'Flat
            Caption         =   "Apply Filter"
            Height          =   285
            Left            =   5880
            TabIndex        =   41
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lacGBStart 
            Caption         =   "Book Dates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   44
            Top             =   285
            Width           =   1290
         End
         Begin VB.Label lacGBEnd 
            Caption         =   " to"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3120
            TabIndex        =   43
            Top             =   285
            Width           =   690
         End
         Begin VB.Label lacGBVehicle 
            Caption         =   "Book Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   42
            Top             =   705
            Width           =   1260
         End
      End
      Begin VB.Label lacViewOption 
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblBookName 
         Caption         =   "Book Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   270
      Picture         =   "ReRateLineBook.frx":0DA4
      Top             =   5310
      Width           =   480
   End
End
Attribute VB_Name = "ReRateLineBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************


'******************************************************
'*  ReRateLineBook - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private smNowDate As String
Private lmNowDate As Long

'2/23/1:Filter information
Dim smFilterStartDate As String
Dim lmFilterStartDate As Long
Dim smFilterEndDate As String
Dim lmFilterEndDate As Long
Dim imFilterVefCode As Integer

Dim smRangeStartDate As String
Dim smRangeEndDate As String

Dim hmCHF As Integer            'Contract header file handle
Dim tmChf As CHF
Dim imCHFRecLen As Integer
Dim tmChfSrchKey0 As LONGKEY0    'Key record image
Dim tmChfSrchKey1 As CHFKEY1
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim tmChfReRate As CHF            'CHF record image
Dim tmClfReRate() As CLFLIST      'CLF record image

Dim smLineNo As String
Dim bmInGrid As Boolean

Const CNTRNOINDEX = 0
Const LINENOINDEX = 1
Const VEHICLEINDEX = 2
Const DAYPARTINDEX = 3
Const LENGTHINDEX = 4
Const FLIGHTDATE = 5
Const PURCHASEBOOKNAMEINDEX = 6
Const RERATEBOOKNAMEINDEX = 7
Const DNFCODEINDEX = 8
Const CHFINDEXINDEX = 9
Const CLFINDEXINDEX = 10
Const ASSIGNMETHODINDEX = 11

Dim smCurrentCntrNo As String '3/3/21 - Bonus improvements: log Apply activity - Keeps track of the last Contract # being examined
Dim imApplied As Integer '3/3/21 - Bonus improvements: log Apply activity - Keeps track of how many contract lines an reRate was applied to

Enum imViews
    VIEWBYBOOK = 0
    VIEWBYLINE = 1
    VIEWBYVEHICLE = 2
    VIEWBYCNTR = 3
End Enum


Dim imView As imViews '3/9/21 - Toggle

Private Sub cbcAssignBookName_GotFocus()
    mSetShow
    'If cbcAssignBookName.ListCount <= 0 Then
        'MsgBox "Press 'Get Books' to obtain the Book names to Assign to lines prior to selecting the Book Name", vbInformation + vbOKOnly, "Click Get Books First"
        'cmcGetBook.SetFocus
    'End If
End Sub

Private Sub cbcAssignBookName_OnChange()
    mSetCommands
End Sub

Private Sub cbcBook_OnChange()
    mSetByBookCommands
End Sub

Private Sub cbcContract_GotFocus()
    mSetShow
End Sub

Private Sub cbcContract_OnChange()
    mSetCommands
End Sub

Private Sub cbcGBVehicle_Click()
    If cbcGBVehicle.ListIndex < 1 And edcGBStart.Text = "" And edcGBEnd.Text = "" Then
        cmcGetBook.Caption = "Cancel"
    Else
        cmcGetBook.Caption = "Apply Filter"
    End If
End Sub

Private Sub cbcGBVehicle_GotFocus()
    mSetShow
End Sub

Private Sub cbcPurchaseBook_OnChange()
    mSetByBookCommands
End Sub

Private Sub cbcVehicle_GotFocus()
    mSetShow
End Sub

Private Sub cbcVehicle_OnChange()
    mSetCommands
End Sub

Private Sub cmcApply_Click()
    Dim llRow As Long
    Dim ilClf As Integer
    Dim slType As String
    Dim llTopRow As Long
    Dim ilApplicable As Integer '3/3/21 - Bonus improvements: log Apply activity
    Dim ilTotalApplied As Integer
    Dim ilRemain As Integer
    lbcApplyLog.Clear
    imApplied = 0

    If imView = VIEWBYBOOK Then
        'Assign by Book
        cmcAssign_Click
    Else
        'Assign by Veh, Contr, Line, etc..
        llTopRow = grdBookGrid.TopRow
        If lmLastClickedRow <> -1 Then
            grdBookGrid.Row = lmLastClickedRow
            grdBookGrid.Col = RERATEBOOKNAMEINDEX
            grdBookGrid.CellBackColor = WHITE
            lmLastClickedRow = -1
        End If
        smCurrentCntrNo = ""
        For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
            If grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then
                If (grdBookGrid.TextMatrix(llRow, CNTRNOINDEX) <> "") Then smCurrentCntrNo = grdBookGrid.TextMatrix(llRow, CNTRNOINDEX)
                ilClf = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
                slType = tgBookByLineAssigned(ilClf).sType
                If slType <> "O" And slType <> "A" And slType <> "E" Then
                    mAssign llRow
                End If
            End If
        Next llRow
        grdBookGrid.TopRow = llTopRow
        ilApplicable = 0 '3/3/21 - Bonus improvements: log Apply activity
        ilTotalApplied = 0
        ilApplicable = 0
        ilRemain = 0
        For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
            'TTP 10172 - 7/1/21 - JW - Issue 2 - check RERATE BOOKNAME column to make sure it's not Read only (yellow) - to know if It needs to be saved
            If (grdBookGrid.TextMatrix(llRow, VEHICLEINDEX) <> "") Then
                grdBookGrid.Row = llRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                If grdBookGrid.CellBackColor = -2147483643 Then
                    If (grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) <> "") Then ilTotalApplied = ilTotalApplied + 1
                    If (grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS") And (grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Then ilRemain = ilRemain + 1
                    If (grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS") Then ilApplicable = ilApplicable + 1
                End If
            End If
        Next llRow
        lbcApplyLog.AddItem "Applied to " & imApplied & " of " & ilApplicable & ". " & IIF(ilRemain < 1, "All Lines assigned.", ilRemain & " Line" & IIF(ilRemain > 1, "s", "") & " remain.")
        lbcApplyLog.Selected(lbcApplyLog.ListCount - 1) = True
    End If
    grdBookGrid_Scroll
End Sub

Private Sub cmcApply_GotFocus()
    mSetShow
End Sub

Private Sub cmcAssign_Click()
    Dim llRow As Long
    Dim ilClf As Integer
    Dim slType As String
    Dim llTopRow As Long
    Dim ilToMap As Integer
    Dim slName As String
    Dim slItemData As String
    Dim ilPurchaseDnfCode As Integer
    Dim ilReRateDnfCode As Integer
    Dim slPurchaseName As String
    Dim slReRateName As String
    Dim ilBook As Integer
    Dim ilPos As Integer
    Dim ilVefCode As Integer
    Dim llNext As Long
    Dim llSvNext As Long   '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
    Dim ilFound As Integer '3/3/21 - Bonus improvements: log Apply activity
    Dim ilApplicable As Integer
    Dim ilTotalApplied As Integer
    Dim ilRemain As Integer
    ilTotalApplied = 0
    imApplied = 0
    ilApplicable = 0
    lbcApplyLog.Clear
    llTopRow = grdBookGrid.TopRow
    For ilToMap = 0 To lbcMap.ListCount - 1 Step 1
        slName = lbcMap.List(ilToMap)
        ilPos = InStr(1, slName, "->")
        If ilPos > 0 Then
            slPurchaseName = Left(slName, ilPos - 1)
            slReRateName = Mid(slName, ilPos + 2)
            ilReRateDnfCode = -1
            For ilBook = 0 To UBound(tgBookInfo) - 1 Step 1
                If Trim$(tgBookInfo(ilBook).sName) = slReRateName Then
                    ilReRateDnfCode = tgBookInfo(ilBook).iDnfCode
                    '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                    llSvNext = tgBookInfo(ilBook).lFirst
                    Exit For
                End If
            Next ilBook
            
            smCurrentCntrNo = "" '3/3/21 - Bonus improvements: log Apply activity (Keep track of current Contract Number for log output)
            'If ilReRateDnfCode >= 0 And ilNext <> -1 Then
            If ilReRateDnfCode >= 0 And llSvNext <> -1 Then '3/2/21 - TTP 10086, ilNext would=-1 when No book found for vehicle, and prevent subsiquent mappings
                slItemData = lbcMap.ItemData(ilToMap)
                ilPurchaseDnfCode = Val(slItemData)
                For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
                    If (grdBookGrid.TextMatrix(llRow, CNTRNOINDEX) <> "") Then smCurrentCntrNo = grdBookGrid.TextMatrix(llRow, CNTRNOINDEX)
                    If (grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS") Then
                                                                        
                        ilClf = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
                        ilVefCode = tgBookByLineAssigned(ilClf).iVefCode
                        slType = tgBookByLineAssigned(ilClf).sType
                        If slType <> "O" And slType <> "A" And slType <> "E" Then
                            If (grdBookGrid.TextMatrix(llRow, PURCHASEBOOKNAMEINDEX) = slPurchaseName) Then
                                grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX) = "B"
                                If ckcDontOverwriteByLine.Value = 0 Then grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "" '3/3/21 - Bonus improvements: "don't overwrite previously assigned lines"
                                grdBookGrid.Row = llRow
                                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                                grdBookGrid.CellForeColor = BLACK
                                '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                                llNext = llSvNext
                                ilFound = 0
                                Do While llNext <> -1
                                    If ilVefCode = tgBookVehicle(llNext).iVefCode Then
                                        ilFound = 1
                                        If (ckcDontOverwriteByLine.Value = 1 And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Or ckcDontOverwriteByLine.Value = 0 Then '3/3/21 - Bonus improvements: "don't overwrite previously assigned lines"
                                            'TTP 10172 - 7/1/21 - JW - Issue #3
                                            If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then
                                                grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = slReRateName
                                                imApplied = imApplied + 1
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                    '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                                    llNext = tgBookVehicle(llNext).lNext
                                Loop
                                If ilFound = 0 Then
                                    lbcApplyLog.AddItem "Book:'" & Trim(slReRateName) & "' doesn't include research for Vehicle:'" & Trim(grdBookGrid.TextMatrix(llRow, VEHICLEINDEX)) & "'.  Contract #" & smCurrentCntrNo & ", Line:" & Trim(grdBookGrid.TextMatrix(llRow, LINENOINDEX))
                                    lbcApplyLog.ItemData(lbcApplyLog.NewIndex) = llRow
                                End If
                                grdBookGrid.CellAlignment = flexAlignLeftCenter
                            End If
                        End If
                    End If
                Next llRow
            End If
        End If
    Next ilToMap
    
    ilApplicable = 0 '3/3/21 - Bonus improvements: log Apply activity
    ilTotalApplied = 0
    ilApplicable = 0
    ilRemain = 0
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        If (grdBookGrid.TextMatrix(llRow, PURCHASEBOOKNAMEINDEX) <> "") Then
            ilTotalApplied = ilTotalApplied + 1
            If (grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS") And (grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Then ilRemain = ilRemain + 1
            If (grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS") Then ilApplicable = ilApplicable + 1
        End If
    Next llRow
    lbcApplyLog.AddItem "Applied to " & imApplied & " of " & ilApplicable & ". " & IIF(ilRemain < 1, "All Lines assigned.", ilRemain & " Line" & IIF(ilRemain > 1, "s", "") & " remain.")
    lbcApplyLog.Selected(lbcApplyLog.ListCount - 1) = True
    
    For ilBook = 0 To lbcMap.ListCount - 1 Step 1
        slName = lbcMap.List(ilBook)
        slItemData = lbcMap.ItemData(ilBook)
        ilPos = InStr(1, slName, "->")
        If ilPos > 0 Then
            slName = Left(slName, ilPos - 1)
            'lbcPurchaseBook.AddItem slName
            'lbcPurchaseBook.ItemData(lbcPurchaseBook.NewIndex) = slItemData
            cbcPurchaseBook.AddItem slName
            cbcPurchaseBook.SetItemData = slItemData
        End If
    Next ilBook
    lbcMap.Clear
    grdBookGrid.TopRow = llTopRow
    mSetByBookCommands
    cmcApply.Enabled = False
    
End Sub

Private Sub cmcCancel_Click()
    igTerminateReturn = 0
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcClear_Click()
    Dim llRow As Long
    
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = ""
    Next llRow
    lbcApplyLog.Clear
    pbcArrow.Visible = False
End Sub

Private Sub cmcClear_GotFocus()
    mSetShow
End Sub

Private Sub cmcDone_Click()
    Dim ilClf As Integer
    Dim llRow As Long
    Dim llClf As Long
    Dim slType As String
    Dim slBookName As String
    Dim ilVef As Integer
    Dim ilDnf As Integer
    Dim blFound As Boolean
    Dim blAllLinesAssignBook As Boolean
    Dim ilRes As Integer
    Dim llColor As Long
    Dim ilPos As Integer
    Dim slAssignedBookName As String
    
    blAllLinesAssignBook = True
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        'TTP 10172 - 7/1/21 - JW - Issue #2 - Items w/o Purch Book dont save
        If grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) <> "" Then
            ilClf = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
            slType = tgBookByLineAssigned(ilClf).sType
            If slType <> "O" And slType <> "A" And slType <> "E" Then
                If grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX) = "L" Then
                    slAssignedBookName = Trim$(cbcAssignBookName.Text)
                Else
                    slAssignedBookName = ""
                End If
                slBookName = grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX)
                ilPos = InStr(1, slBookName, ":")
                If ilPos > 0 Then
                    slBookName = Trim$(Left(slBookName, ilPos - 1))
                End If
                If (slBookName <> "") Or (slAssignedBookName = "[Closest to Air Date]") Then
                    grdBookGrid.Row = llRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    llColor = grdBookGrid.CellForeColor
                    If (slBookName = "[Vehicle Default]") Or (llColor = DARKGREEN) Then
                        ilVef = gBinarySearchVef(tgBookByLineAssigned(ilClf).iVefCode)
                        If ilVef <> -1 Then
                            '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                            tgBookByLineAssigned(ilClf).iReRateDnfCode = tgMVef(ilVef).iDnfCode
                        Else
                            tgBookByLineAssigned(ilClf).iReRateDnfCode = -1
                        End If
                    'TTP 10172 - 7/1/21 - JW - Issue #2 - Items w/o Purch Book dont Save
                    ElseIf (slBookName = "[Closest to Air Date]") Or (slAssignedBookName = "[Closest to Air Date]") Or (slAssignedBookName <> "" And llColor = BLUE) Then
                        tgBookByLineAssigned(ilClf).iReRateDnfCode = -2
                    ElseIf (slBookName = "[Purchase Book]") Or (llColor = ORANGE) Then
                        tgBookByLineAssigned(ilClf).iReRateDnfCode = -3
                        'Find match
                        blFound = False
                        slBookName = grdBookGrid.TextMatrix(llRow, PURCHASEBOOKNAMEINDEX)
                        For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
                            If Trim(tgBookInfo(ilDnf).sName) = slBookName Then
                                '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                                tgBookByLineAssigned(ilClf).iReRateDnfCode = tgBookInfo(ilDnf).iDnfCode
                                blFound = True
                                Exit For
                            End If
                        Next ilDnf
                        If Not blFound Then
                            If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then blAllLinesAssignBook = False
                            tgBookByLineAssigned(ilClf).iReRateDnfCode = 0
                        End If
                    Else
                        'Find match
                        blFound = False
                        For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
                            If Trim(tgBookInfo(ilDnf).sName) = slBookName Then
                                '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                                tgBookByLineAssigned(ilClf).iReRateDnfCode = tgBookInfo(ilDnf).iDnfCode
                                blFound = True
                                    Exit For
                            End If
                        Next ilDnf
                        If Not blFound Then
                            If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then blAllLinesAssignBook = False
                            tgBookByLineAssigned(ilClf).iReRateDnfCode = 0
                        End If
                    End If
                Else
                    If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then blAllLinesAssignBook = False
                    tgBookByLineAssigned(ilClf).iReRateDnfCode = 0
                End If
                grdBookGrid.Row = llRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                tgBookByLineAssigned(ilClf).lColor = grdBookGrid.CellForeColor
                tgBookByLineAssigned(ilClf).sDaypartName = grdBookGrid.TextMatrix(llRow, DAYPARTINDEX)
                tgBookByLineAssigned(ilClf).sAssignMethod = grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX)
            End If
        'TTP 10172 - 7/1/21 - JW - Issue #6 - wasnt warned when Some were still empty
        ElseIf grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "" Then
            ilClf = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
            tgBookByLineAssigned(ilClf).iReRateDnfCode = 0
            If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then
                grdBookGrid.Row = llRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                If grdBookGrid.CellBackColor = -2147483643 Then
                    blAllLinesAssignBook = False
                End If
            End If
        End If
    Next llRow
    If Not blAllLinesAssignBook Then
        ilRes = MsgBox("Lines missing Research Book Assignment, Continue with Ok", vbYesNo + vbExclamation, "Incomplete")
        If ilRes = vbNo Then
            Exit Sub
        End If
    End If
    igTerminateReturn = 1
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcFilterBooks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmcFilterBooks.Caption = "Filter Books.." Then
        'Show the Filters
        frcFilter.Visible = True
        edcGBStart.Visible = True
        edcGBEnd.Visible = True
        cbcGBVehicle.Visible = True
        cmcFilterBooks.Enabled = False
        cmcGetBook.Caption = "Cancel"
    Else
        'Clear the Filters
        edcGBStart.Text = ""
        edcGBEnd.Text = ""
        cbcGBVehicle.Text = "[All Vehicles]"
        cmcGetBook.Caption = "Apply Filter"
        cmcGetBook_Click
        
        frcFilter.Visible = False
        edcGBStart.Visible = False
        edcGBEnd.Visible = False
        cbcGBVehicle.Visible = False
        cmcFilterBooks.Enabled = True
        cmcFilterBooks.Caption = "Filter Books.."
    End If
    cbcAssignBookName.SetFocus
End Sub

Private Sub cmcFilterReRateBooks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'TTP 10143
    If cmcFilterReRateBooks.Caption = "Filter Books.." Then
        'Show the Filters
        If cbcAssignBookName.Visible = True Then cbcAssignBookName.SetFocus
        frcFilterReRateBook.Visible = True
        edcGRBStart.Visible = True
        edcGRBEnd.Visible = True
        'cmcGetReRateBook.Enabled = False
        cmcGetReRateBook.Caption = "Cancel"
    Else
        'Clear the Filters
        edcGRBStart.Text = ""
        edcGRBEnd.Text = ""
        cbcGBVehicle.Text = "[All Vehicles]"
        cmcGetReRateBook.Caption = "Apply Filter"
        cmcGetReRateBook_Click
        
        frcFilterReRateBook.Visible = False
        edcGRBStart.Visible = False
        edcGRBEnd.Visible = False
        cmcFilterReRateBooks.Enabled = True
        cmcFilterReRateBooks.Caption = "Filter Books.."
    End If
    cmcFilterReRateBooks.SetFocus
End Sub

Private Sub cmcFromMap_Click()
    mMoveName "FromMap"
    mSetCommands
End Sub

Private Sub cmcGetBook_Click()
    If cmcGetBook.Caption = "Cancel" Then
        cmcFilterBooks.Enabled = True
        cmcFilterBooks.Caption = "Filter Books.."
        'Hide the Filter inputs
        If cbcAssignBookName.Text = "" Then cbcAssignBookName.Text = "[Vehicle Default]"
        If cbcAssignBookName.Visible = True Then cbcAssignBookName.SetFocus
        frcFilter.Visible = False
        edcGBStart.Visible = False
        edcGBEnd.Visible = False
        cbcGBVehicle.Visible = False
        cbcAssignBookName.SetFocus
        Exit Sub
    End If
    If edcGBStart.Text <> "" Then
        If gIsDate(edcGBStart.Text) Then
            smFilterStartDate = edcGBStart.Text
            lmFilterStartDate = gDateValue(edcGBStart.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterStartDate = ""
        lmFilterStartDate = 0
    End If
    If edcGBEnd.Text <> "" Then
        If gIsDate(edcGBEnd.Text) Then
            smFilterEndDate = edcGBEnd.Text
            lmFilterEndDate = gDateValue(edcGBEnd.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterEndDate = ""
        lmFilterEndDate = 0
    End If
    If cbcGBVehicle.ListIndex > 0 Then
        imFilterVefCode = cbcGBVehicle.ItemData(cbcGBVehicle.ListIndex)
    Else
        imFilterVefCode = -1
    End If
    Screen.MousePointer = vbHourglass  'Wait
    mApplyFilter 0 ' mPopAssignBookNames
    Screen.MousePointer = vbDefault
    
    If cbcAssignBookName.Text = "" Then cbcAssignBookName.Text = "[Vehicle Default]"
    If cbcAssignBookName.Visible = True Then cbcAssignBookName.SetFocus
    
    'Filters applied
    frcFilter.Visible = False
    edcGBStart.Visible = False
    edcGBEnd.Visible = False
    cbcGBVehicle.Visible = False
    cmcFilterBooks.Caption = "Clear Filters"
    cmcFilterBooks.Enabled = True
End Sub

Private Sub cmcGetBook_GotFocus()
    mSetShow
End Sub

Private Sub cmcGetReRateBook_Click()
    'TTP 10143
    If cmcGetReRateBook.Caption = "Cancel" Then
        cmcFilterReRateBooks.Enabled = True
        cmcFilterReRateBooks.Caption = "Filter Books.."
        'Hide the Filter inputs
        'Filters applied
        If cbcBook.Visible = True Then cbcBook.SetFocus
    
        frcFilterReRateBook.Visible = False
        edcGRBStart.Visible = False
        edcGRBEnd.Visible = False
        cmcFilterReRateBooks.Visible = True
        cmcFilterReRateBooks.Enabled = True
        Exit Sub
    End If
    
    If edcGRBStart.Text <> "" Then
        If gIsDate(edcGRBStart.Text) Then
            smFilterStartDate = edcGRBStart.Text
            lmFilterStartDate = gDateValue(edcGRBStart.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterStartDate = ""
        lmFilterStartDate = 0
    End If
    If edcGRBEnd.Text <> "" Then
        If gIsDate(edcGRBEnd.Text) Then
            smFilterEndDate = edcGRBEnd.Text
            lmFilterEndDate = gDateValue(edcGRBEnd.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterEndDate = ""
        lmFilterEndDate = 0
    End If
    
    'no Vehicle Filter
    imFilterVefCode = -1
    
    Screen.MousePointer = vbHourglass  'Wait
    mApplyFilter 1
    Screen.MousePointer = vbDefault
   
    'Filters applied
    If cbcBook.Visible = True Then cbcBook.SetFocus
    
    frcFilterReRateBook.Visible = False
    edcGRBStart.Visible = False
    edcGRBEnd.Visible = False
    cmcFilterReRateBooks.Visible = True
    cmcFilterReRateBooks.Enabled = True
    cmcFilterReRateBooks.Caption = "Clear Filters"
End Sub

Private Sub cmcToMap_Click()
    mMoveName "ToMap"
    mSetCommands
End Sub

Private Sub edcGBEnd_Change()
    If cbcGBVehicle.ListIndex < 1 And edcGBStart.Text = "" And edcGBEnd.Text = "" Then
        cmcGetBook.Caption = "Cancel"
    Else
        cmcGetBook.Caption = "Apply Filter"
    End If
End Sub

Private Sub edcGBEnd_GotFocus()
    mSetShow
End Sub

Private Sub edcGBStart_Change()
    If cbcGBVehicle.ListIndex < 1 And edcGBStart.Text = "" And edcGBEnd.Text = "" Then
        cmcGetBook.Caption = "Cancel"
    Else
        cmcGetBook.Caption = "Apply Filter"
    End If
End Sub

Private Sub edcGBStart_GotFocus()
    mSetShow
End Sub

Private Sub edcGRBEnd_Change()
    'TTP 10143
    If edcGRBStart.Text = "" And edcGRBEnd.Text = "" Then
        cmcGetReRateBook.Caption = "Cancel"
    Else
        cmcGetReRateBook.Caption = "Apply Filter"
    End If
End Sub

Private Sub edcGRBEnd_GotFocus()
    mSetShow
End Sub

Private Sub edcGRBStart_Change()
    'TTP 10143
    If edcGRBStart.Text = "" And edcGRBEnd.Text = "" Then
        cmcGetReRateBook.Caption = "Cancel"
    Else
        cmcGetReRateBook.Caption = "Apply Filter"
    End If
End Sub

Private Sub edcGRBStart_GotFocus()
    mSetShow
End Sub

Private Sub edcLines_Change()
'    mSetCommands
    If imView = VIEWBYCNTR And edcLines.Text <> "" And cbcContract.Text <> "" Then
        If mGetLineNo <> "" Then
            cmcApply.Enabled = True
        Else
            cmcApply.Enabled = False
        End If
    Else
        cmcApply.Enabled = False
    End If
End Sub

Private Sub edcLines_GotFocus()
    mSetShow
End Sub

Private Sub edcLines_LostFocus()
    smLineNo = mGetLineNo()
End Sub

Private Sub Form_Activate()
    If imFirstTime Then
        imFirstTime = False
        Screen.MousePointer = vbDefault
        
        pbcView_Paint
    End If
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Dim ilRet As Integer
    Dim blSetBooks As Boolean
    Dim llRet As Long
    Dim illoop As Integer 'TTP 10172 - 7/1/21 - JW - New Behavior #1
    
    'Me.Width = (CLng(40) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = 5760
    Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    If (UBound(tgBookByLineAssigned) <= LBound(tgBookByLineAssigned)) Then
        blSetBooks = True
    Else
        blSetBooks = False
    End If
    mSetControls
    mPopCntr
    'mPopAssignBookNames
    mClearGrid
    mPopulate
    mPopVehicle
    
    mPopBook
    mPopPurchaseBook
    llRet = SendMessageByNum(lbcMap.HWnd, LB_SETHORIZONTALEXTENT, 0, 0)
    'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        If ExptReRate.rbcReRateBook(3).Value Then
            'rbcByBook.Value = True
            imView = VIEWBYBOOK
        Else
            'rbcByLine.Value = True
            imView = VIEWBYLINE
        End If
    'Else
    '    rbcByLine.Value = True
    '    rbcByLine.Visible = False
    '    rbcByBook.Visible = False
    'End If
    If (ExptReRate.rbcReRateBook(0).Value Or ExptReRate.rbcReRateBook(1).Value) And (blSetBooks) Then
        cmcGetBook_Click
        If ExptReRate.rbcReRateBook(0).Value Then
            cbcAssignBookName.SetListIndex = 0
        Else
            cbcAssignBookName.SetListIndex = 1
        End If
        'TTP 10172 - 7/1/21 - JW - New Behavior #1
        For illoop = 0 To UBound(tgBookByLineAssigned) - 1
            If tgBookByLineAssigned(illoop).iReRateDnfCode > 0 Then ckcDontOverwriteByLine.Value = 1: Exit For
        Next illoop
        cmcApply_Click
    Else
        'Reassign books to lines
        cmcGetBook_Click
        If ExptReRate.rbcReRateBook(0).Value Then
            cbcAssignBookName.SetListIndex = 0
        Else
            cbcAssignBookName.SetListIndex = 1
        End If
        mReassignBooksToLines
        'TTP 10172 - 7/1/21 - JW - New Behavior #1
        For illoop = 0 To UBound(tgBookByLineAssigned) - 1
            If tgBookByLineAssigned(illoop).iReRateDnfCode > 0 Then ckcDontOverwriteByLine.Value = 1: Exit For
        Next illoop
    End If
    gCenterStdAlone ReRateLineBook
    cmcFilterBooks.Caption = "Filter Books.."
    DoEvents
End Sub

Private Sub Form_Load()
    imView = VIEWBYBOOK 'Default View
    Screen.MousePointer = vbHourglass

    mInit
    Screen.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    
    Set ReRateLineBook = Nothing
End Sub


Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    'Blank rows within grid
'    gGrid_Clear grdBookGrid, True
    'Set color within cells
    grdBookGrid.RowHeight(0) = fgBoxGridH + 15
    gGrid_IntegralHeight grdBookGrid, fgBoxGridH + 15
    gGrid_FillWithRows grdBookGrid, fgBoxGridH + 15
    grdBookGrid.Height = grdBookGrid.Height + 60
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        For llCol = CNTRNOINDEX To PURCHASEBOOKNAMEINDEX Step 1
            grdBookGrid.Row = llRow
            grdBookGrid.Col = llCol
            grdBookGrid.CellBackColor = LIGHTYELLOW
        Next llCol
        grdBookGrid.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    grdBookGrid.ColAlignment(VEHICLEINDEX) = flexAlignLeftCenter
    grdBookGrid.ColAlignment(PURCHASEBOOKNAMEINDEX) = flexAlignLeftCenter
    grdBookGrid.ColAlignment(RERATEBOOKNAMEINDEX) = flexAlignLeftCenter
    
    grdBookGrid.ColAlignment(FLIGHTDATE) = flexAlignLeftCenter
    
    grdBookGrid.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer
    
    gSetMousePointer grdBookGrid, grdBookGrid, vbHourglass
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdBookGrid.FixedRows
    cbcVehicle.Clear
    cbcAssignBookName.Clear
    cbcContract.Clear
    cbcLnBookName.Clear
    
    smLineNo = ""
    'edcGBStart.Text = ""
    'edcGBEnd.Text = ""
    
    imcKey.Picture = IconTraf!imcKey.Picture
    mPopListKey
    
    mDetermineDateRange
        
    cbcLnBookName.BackColor = &HFFFF00
   
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ReRateLineBook
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ReRateLineBook
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ReRateLineBook
    imCffRecLen = Len(tmCff)
    
    'cbcPurchaseBook.Left = frcAssignBook.Left + lbcPurchaseBook.Left + 120
    'cbcPurchaseBook.Top = rbcByLine.Top + rbcByLine.Height + lbcPurchaseBook.Top
    'cbcBook.Left = frcAssignBook.Left + lbcBook.Left + 120
    cbcPurchaseBook.SetDropDownCharWidth 36
    cbcBook.SetDropDownCharWidth 36
    '3/9/21 new UI
    Me.Width = 13185
    pbcView.Top = 240
    pbcView.Left = 1560
    
    frcByBook.Left = 240
    frcByBook.Top = 560
    cbcPurchaseBook.Left = 1680
    cbcPurchaseBook.Top = 600
    edcGRBStart.Left = 1680
    edcGRBEnd.Left = 3825

    cbcBook.Left = 1680
    cbcBook.Top = 960
    
    edcGRBStart.Top = 1500
    edcGRBEnd.Top = 1500
    
    cbcAssignBookName.Left = 1680
    cbcAssignBookName.Top = 570
    cmcFilterBooks.Top = 570
    
    edcLines.Left = 1560
    edcLines.Top = 600
    cbcVehicle.Left = 8400
    cbcVehicle.Top = 240
    cbcContract.Left = 9000
    cbcContract.Top = 240
    edcGBStart.Visible = False
    edcGBEnd.Visible = False
    cbcGBVehicle.Visible = False
    edcGBStart.Top = 1100
    edcGBEnd.Top = 1100
    lacGBStart.Top = 285
    lacGBEnd.Top = 285
    lbcMap.Top = 210
    lbcMap.Left = 7560
    lbcMap.Width = 5305
    lbcMap.Height = 1100
    mSetCommands
    
    Screen.MousePointer = vbDefault
    gSetMousePointer grdBookGrid, grdBookGrid, vbDefault
    
    mSetOptions imView
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdBookGrid, grdBookGrid, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim ilChf As Integer
    Dim ilClf As Integer
    Dim llPrevNextIndex As Long
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim blFirstLine As Boolean
    Dim ilNext As Integer
    Dim ilIndex As Integer
    Dim ilDnf As Integer
    Dim ilLen As Integer

    On Error GoTo ErrHand:
    If UBound(tgBookByLineCntr) <= LBound(tgBookByLineCntr) Then
        Exit Sub
    End If
    grdBookGrid.Redraw = False
    grdBookGrid.RowHeight(0) = fgBoxGridH + 15
    llRow = grdBookGrid.FixedRows
    For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
        If tgBookByLineCntr(ilChf).sSelected = "1" Then
            If tgBookByLineCntr(ilChf).iFirst = -1 Then
                ReDim tmClfReRate(0 To 0) As CLFLIST
                'ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tgBookByLineCntr(ilChf).lChfCode, False, tmChfReRate, tmClfReRate(), tmCffReRate(), False)  '8-28-12 do not sort by special dp order
                'Switch over to:
                ilRet = gObtainChfClf(hmCHF, hmClf, tgBookByLineCntr(ilChf).lChfCode, False, tmChfReRate, tmClfReRate())
                
                'Build tgBookByLineAssigned
                llPrevNextIndex = -1
                For ilClf = 0 To UBound(tmClfReRate) - 1 Step 1
                    For ilLen = 0 To UBound(igReRateAllowedLengths) - 1 Step 1
                        If igReRateAllowedLengths(ilLen) = tmClfReRate(ilClf).ClfRec.iLen Then
                            llUpper = UBound(tgBookByLineAssigned)
                            tgBookByLineAssigned(llUpper).lClfCode = tmClfReRate(ilClf).ClfRec.lCode
                            tgBookByLineAssigned(llUpper).sType = tmClfReRate(ilClf).ClfRec.sType
                            tgBookByLineAssigned(llUpper).iPurchaseDnfCode = tmClfReRate(ilClf).ClfRec.iDnfCode '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                            tgBookByLineAssigned(llUpper).iReRateDnfCode = 0
                            tgBookByLineAssigned(llUpper).iLineNo = tmClfReRate(ilClf).ClfRec.iLine
                            tgBookByLineAssigned(llUpper).iPkLineNo = tmClfReRate(ilClf).ClfRec.iPkLineNo
                            tgBookByLineAssigned(llUpper).iVefCode = tmClfReRate(ilClf).ClfRec.iVefCode
                            tgBookByLineAssigned(llUpper).iRdfCode = tmClfReRate(ilClf).ClfRec.iRdfCode
                            tgBookByLineAssigned(llUpper).iLen = tmClfReRate(ilClf).ClfRec.iLen
                            tgBookByLineAssigned(llUpper).lChfCode = tmClfReRate(ilClf).ClfRec.lChfCode
                            gUnpackDateLong tmClfReRate(ilClf).ClfRec.iStartDate(0), tmClfReRate(ilClf).ClfRec.iStartDate(1), tgBookByLineAssigned(llUpper).lStartDate 'cffStartDate
                            gUnpackDateLong tmClfReRate(ilClf).ClfRec.iEndDate(0), tmClfReRate(ilClf).ClfRec.iEndDate(1), tgBookByLineAssigned(llUpper).lEndDate 'cffEndDate
                            tgBookByLineAssigned(llUpper).iMGCount = 0
                            tgBookByLineAssigned(llUpper).iOutsideCount = 0
                            tgBookByLineAssigned(llUpper).iBonusCount = 0
                            tgBookByLineAssigned(llUpper).sAssignMethod = ""
                            tgBookByLineAssigned(llUpper).iNext = -1
                            If llPrevNextIndex = -1 Then
                                tgBookByLineCntr(ilChf).iFirst = llUpper
                            Else
                                tgBookByLineAssigned(llPrevNextIndex).iNext = llUpper
                            End If
                            llPrevNextIndex = llUpper
                            ReDim Preserve tgBookByLineAssigned(0 To llUpper + 1) As BOOKBYLINEASSIGNED
                            'See if MG/Outsides/Fills exist
                            mMGBonusExist tgBookByLineCntr(ilChf).lChfCode, tmClfReRate(ilClf).ClfRec.iLine, llPrevNextIndex
                            Exit For
                        End If
                    Next ilLen
                Next ilClf
            End If
        End If
    Next ilChf
    'For ilLoop = 0 To UBound(tmGnf) - 1 Step 1
    '    If llRow >= grdBookGrid.Rows Then
    '        grdBookGrid.AddItem ""
    '    End If
    '    grdBookGrid.RowHeight(llRow) = fgBoxGridH + 15
    '
    '    llRow = llRow + 1
    'Next ilLoop
    llRow = llRow - 1
    For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
        If tgBookByLineCntr(ilChf).sSelected = "1" Then
            llRow = llRow + 1
            If llRow >= grdBookGrid.Rows Then
                grdBookGrid.AddItem ""
            End If
            grdBookGrid.RowHeight(llRow) = fgBoxGridH + 15
            grdBookGrid.TextMatrix(llRow, CNTRNOINDEX) = tgBookByLineCntr(ilChf).lCntrNo
            grdBookGrid.TextMatrix(llRow, CHFINDEXINDEX) = ilChf
            blFirstLine = True
            ilNext = tgBookByLineCntr(ilChf).iFirst
            Do While ilNext <> -1
                If (tgBookByLineAssigned(ilNext).sType <> "H") Then
                    If Not blFirstLine Then
                        llRow = llRow + 1
                        If llRow >= grdBookGrid.Rows Then
                            grdBookGrid.AddItem ""
                        End If
                        grdBookGrid.RowHeight(llRow) = fgBoxGridH + 15
                    End If
                    grdBookGrid.Row = llRow
                    blFirstLine = False
                    grdBookGrid.TextMatrix(llRow, CHFINDEXINDEX) = ilChf
                    grdBookGrid.TextMatrix(llRow, LINENOINDEX) = tgBookByLineAssigned(ilNext).iLineNo
                    'grdBookGrid.Col = LINENOINDEX
                    'grdBookGrid.CellAlignment = flexAlignLeftCenter
                    ilVef = gBinarySearchVef(tgBookByLineAssigned(ilNext).iVefCode)
                    If ilVef <> -1 Then
                        grdBookGrid.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                    Else
                        grdBookGrid.TextMatrix(llRow, VEHICLEINDEX) = "Missing:" & tgBookByLineAssigned(ilNext).iVefCode
                    End If
                    If (tgBookByLineAssigned(ilNext).iMGCount <= 0) And (tgBookByLineAssigned(ilNext).iOutsideCount <= 0) And (tgBookByLineAssigned(ilNext).iBonusCount <= 0) Then
                        ilRdf = gBinarySearchRdf(tgBookByLineAssigned(ilNext).iRdfCode)
                        If ilRdf <> -1 Then
                            grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = Trim$(tgMRdf(ilRdf).sName)
                        Else
                            grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "Missing:" & tgBookByLineAssigned(ilNext).iRdfCode
                        End If
                    Else
                        If ((tgBookByLineAssigned(ilNext).iMGCount > 0) Or (tgBookByLineAssigned(ilNext).iOutsideCount > 0)) And (tgBookByLineAssigned(ilNext).iBonusCount > 0) Then
                            grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "MG's/Bonus"
                        ElseIf ((tgBookByLineAssigned(ilNext).iMGCount > 0) Or (tgBookByLineAssigned(ilNext).iOutsideCount > 0)) And (tgBookByLineAssigned(ilNext).iBonusCount <= 0) Then
                            grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "MG's"
                        ElseIf ((tgBookByLineAssigned(ilNext).iMGCount <= 0) And (tgBookByLineAssigned(ilNext).iOutsideCount <= 0)) And (tgBookByLineAssigned(ilNext).iBonusCount > 0) Then
                            grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "Bonus"
                        End If
                    End If
                    grdBookGrid.TextMatrix(llRow, LENGTHINDEX) = tgBookByLineAssigned(ilNext).iLen
                    
                    If tgBookByLineAssigned(ilNext).lStartDate > tgBookByLineAssigned(ilNext).lEndDate Then
'                        'CBS
                        grdBookGrid.TextMatrix(llRow, FLIGHTDATE) = "CBS"
                    Else
                        grdBookGrid.TextMatrix(llRow, FLIGHTDATE) = Format(tgBookByLineAssigned(ilNext).lStartDate, "ddddd") + " - " & Format(tgBookByLineAssigned(ilNext).lEndDate, "ddddd")
                        grdBookGrid.Row = llRow
                        grdBookGrid.Col = RERATEBOOKNAMEINDEX
                        'Dont write ReRate bookname if a Yellow cell
                        If grdBookGrid.CellBackColor <> -2147483643 Then
                            mSetBookNameInGrid llRow, RERATEBOOKNAMEINDEX, tgBookByLineAssigned(ilClf).iReRateDnfCode
                        End If
                    End If

                    mSetBookNameInGrid llRow, PURCHASEBOOKNAMEINDEX, tgBookByLineAssigned(ilNext).iPurchaseDnfCode
                    grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX) = ilNext
                    If (tgBookByLineAssigned(ilNext).sType = "O") Or (tgBookByLineAssigned(ilNext).sType = "A") Or (tgBookByLineAssigned(ilNext).sType = "E") Then
                        'Get hidden lines
                        For ilClf = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                            If (tgBookByLineAssigned(ilClf).sType = "H") And (tgBookByLineAssigned(ilClf).lChfCode = tgBookByLineAssigned(ilNext).lChfCode) And (tgBookByLineAssigned(ilClf).iPkLineNo = tgBookByLineAssigned(ilNext).iLineNo) Then
                                llRow = llRow + 1
                                If llRow >= grdBookGrid.Rows Then
                                    grdBookGrid.AddItem ""
                                End If
                                grdBookGrid.Row = llRow
                                grdBookGrid.TextMatrix(llRow, CHFINDEXINDEX) = ilChf
                                grdBookGrid.RowHeight(llRow) = fgBoxGridH + 15
                                grdBookGrid.TextMatrix(llRow, LINENOINDEX) = tgBookByLineAssigned(ilClf).iLineNo
                                'grdBookGrid.Col = LINENOINDEX
                                'grdBookGrid.CellAlignment = flexAlignRightCenter
                                ilVef = gBinarySearchVef(tgBookByLineAssigned(ilClf).iVefCode)
                                If ilVef <> -1 Then
                                    grdBookGrid.TextMatrix(llRow, VEHICLEINDEX) = "    " & Trim$(tgMVef(ilVef).sName)
                                Else
                                    grdBookGrid.TextMatrix(llRow, VEHICLEINDEX) = "Missing:" & tgBookByLineAssigned(ilClf).iVefCode
                                End If
                                If (tgBookByLineAssigned(ilClf).iMGCount <= 0) And (tgBookByLineAssigned(ilClf).iOutsideCount <= 0) And (tgBookByLineAssigned(ilClf).iBonusCount <= 0) Then
                                    'ilRdf = gBinarySearchRdf(tgBookByLineAssigned(ilNext).iRdfCode)
                                    ilRdf = gBinarySearchRdf(tgBookByLineAssigned(ilClf).iRdfCode) '3/4/21 - TTP 10087: Research Books screen, For hidden lines, it shows the package daypart as the line daypart.
                                    If ilRdf <> -1 Then
                                        grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = Trim$(tgMRdf(ilRdf).sName)
                                    Else
                                        grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "Missing:" & tgBookByLineAssigned(ilNext).iRdfCode
                                    End If
                                Else
                                    If ((tgBookByLineAssigned(ilClf).iMGCount > 0) Or (tgBookByLineAssigned(ilClf).iOutsideCount > 0)) And (tgBookByLineAssigned(ilClf).iBonusCount > 0) Then
                                        grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "MG's/Bonus"
                                    ElseIf ((tgBookByLineAssigned(ilClf).iMGCount > 0) Or (tgBookByLineAssigned(ilClf).iOutsideCount > 0)) And (tgBookByLineAssigned(ilClf).iBonusCount <= 0) Then
                                        grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "MG's"
                                    ElseIf ((tgBookByLineAssigned(ilClf).iMGCount <= 0) And (tgBookByLineAssigned(ilClf).iOutsideCount <= 0)) And (tgBookByLineAssigned(ilClf).iBonusCount > 0) Then
                                        grdBookGrid.TextMatrix(llRow, DAYPARTINDEX) = "Bonus"
                                    End If
                                End If
                                
                                grdBookGrid.TextMatrix(llRow, LENGTHINDEX) = tgBookByLineAssigned(ilClf).iLen
                                
                                If tgBookByLineAssigned(ilClf).lStartDate > tgBookByLineAssigned(ilClf).lEndDate Then
                                    'CBS
                                    grdBookGrid.TextMatrix(llRow, FLIGHTDATE) = "CBS"
                                Else
                                    grdBookGrid.TextMatrix(llRow, FLIGHTDATE) = Format(tgBookByLineAssigned(ilClf).lStartDate, "ddddd") + " - " & Format(tgBookByLineAssigned(ilClf).lEndDate, "ddddd")
                                End If
                                
                                'If tgBookByLineAssigned(ilClf).iReRateDnfCode > 0 Then
                                '    ilDnf = -1
                                '    For ilLoop = 0 To UBound(tgBookInfo) - 1 Step 1
                                '        If tgBookByLineAssigned(ilClf).iReRateDnfCode = tgBookInfo(ilLoop).iDnfCode Then
                                '            ilDnf = ilLoop
                                '            Exit For
                                '        End If
                                '    Next ilLoop
                                '    If ilDnf <> -1 Then
                                '        grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = Trim$(tgBookInfo(ilDnf).sName)
                                '    Else
                                '        grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = ""
                                '    End If
                                'End If
                                mSetBookNameInGrid llRow, PURCHASEBOOKNAMEINDEX, tgBookByLineAssigned(ilClf).iPurchaseDnfCode
                                
                                If tgBookByLineAssigned(ilClf).lStartDate > tgBookByLineAssigned(ilClf).lEndDate Then
                                    'CBS
                                Else
                                    mSetBookNameInGrid llRow, RERATEBOOKNAMEINDEX, tgBookByLineAssigned(ilClf).iReRateDnfCode
                                End If
                                
                                grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX) = ilClf
                                grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX) = tgBookByLineAssigned(ilClf).sAssignMethod
                            End If
                        Next ilClf
                    End If
                    'llRow = llRow + 1
                End If
                ilNext = tgBookByLineAssigned(ilNext).iNext
            Loop
        End If
    Next ilChf
    
    'Color the grid
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        slStr = Trim$(grdBookGrid.TextMatrix(llRow, LINENOINDEX))
        If slStr <> "" Then
            For llCol = CNTRNOINDEX To RERATEBOOKNAMEINDEX Step 1
                grdBookGrid.Row = llRow
                grdBookGrid.Col = llCol
                ilIndex = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
                If llCol = RERATEBOOKNAMEINDEX Then
                    If (tgBookByLineAssigned(ilIndex).sType = "O") Or (tgBookByLineAssigned(ilIndex).sType = "A") Or (tgBookByLineAssigned(ilIndex).sType = "E") Then
                        grdBookGrid.CellBackColor = LIGHTYELLOW
                    End If
                    If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) = "CBS" Then
                        grdBookGrid.CellBackColor = LIGHTYELLOW
                    End If
                Else
                    grdBookGrid.CellBackColor = LIGHTYELLOW
                End If
                If (tgBookByLineAssigned(ilIndex).sType = "O") Or (tgBookByLineAssigned(ilIndex).sType = "A") Or (tgBookByLineAssigned(ilIndex).sType = "E") Then
                    grdBookGrid.CellForeColor = BLUE
                ElseIf tgBookByLineAssigned(ilIndex).sType = "H" Then
                    grdBookGrid.CellForeColor = MIDDLEBLUE
                Else
                    grdBookGrid.CellForeColor = BLACK
                End If
                If ((tgBookByLineAssigned(ilIndex).iMGCount > 0) Or (tgBookByLineAssigned(ilIndex).iOutsideCount > 0)) Or (tgBookByLineAssigned(ilIndex).iBonusCount > 0) Then
                    grdBookGrid.CellFontItalic = True
                End If
                If tgBookByLineAssigned(ilIndex).lStartDate > tgBookByLineAssigned(ilIndex).lEndDate And llCol = FLIGHTDATE Then
                    grdBookGrid.CellForeColor = vbRed
                End If
            Next llCol
        End If
    Next llRow
    'gGrid_AlignAllColsLeft grdBookGrid

    grdBookGrid.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    Resume Next
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim illoop As Integer

    grdBookGrid.ColWidth(DNFCODEINDEX) = 0
    grdBookGrid.ColWidth(CHFINDEXINDEX) = 0
    grdBookGrid.ColWidth(CLFINDEXINDEX) = 0
    grdBookGrid.ColWidth(ASSIGNMETHODINDEX) = 0
    grdBookGrid.ColWidth(CNTRNOINDEX) = grdBookGrid.Width * 0.06
    grdBookGrid.ColWidth(LINENOINDEX) = grdBookGrid.Width * 0.034
    grdBookGrid.ColWidth(VEHICLEINDEX) = grdBookGrid.Width * 0.189
    grdBookGrid.ColWidth(DAYPARTINDEX) = grdBookGrid.Width * 0.1
    grdBookGrid.ColWidth(LENGTHINDEX) = grdBookGrid.Width * 0.032
    grdBookGrid.ColWidth(FLIGHTDATE) = grdBookGrid.Width * 0.11
    grdBookGrid.ColWidth(PURCHASEBOOKNAMEINDEX) = grdBookGrid.Width * 0.2
    grdBookGrid.ColWidth(RERATEBOOKNAMEINDEX) = grdBookGrid.Width * 0.22
    llMinWidth = grdBookGrid.Width
    For ilCol = 0 To grdBookGrid.Cols - 1 Step 1
        llWidth = llWidth + grdBookGrid.ColWidth(ilCol)
        If (grdBookGrid.ColWidth(ilCol) > 15) And (grdBookGrid.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdBookGrid.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdBookGrid.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdBookGrid.Width
            For ilCol = 0 To grdBookGrid.Cols - 1 Step 1
                If (grdBookGrid.ColWidth(ilCol) > 15) And (grdBookGrid.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdBookGrid.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdBookGrid.FixedCols To grdBookGrid.Cols - 1 Step 1
                If grdBookGrid.ColWidth(ilCol) > 15 Then
                    ilColInc = grdBookGrid.ColWidth(ilCol) / llMinWidth
                    For illoop = 1 To ilColInc Step 1
                        grdBookGrid.ColWidth(ilCol) = grdBookGrid.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next illoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdBookGrid.TextMatrix(0, CNTRNOINDEX) = "Contract"
    grdBookGrid.TextMatrix(0, LINENOINDEX) = "Line"
    grdBookGrid.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdBookGrid.TextMatrix(0, DAYPARTINDEX) = "Daypart"
    grdBookGrid.TextMatrix(0, LENGTHINDEX) = "Len"
    grdBookGrid.TextMatrix(0, FLIGHTDATE) = "Flight Date"
    grdBookGrid.TextMatrix(0, PURCHASEBOOKNAMEINDEX) = "Purchase Book Name"
    grdBookGrid.TextMatrix(0, RERATEBOOKNAMEINDEX) = "ReRate Book Name"
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'

    Screen.MousePointer = vbDefault
    gSetMousePointer grdBookGrid, grdBookGrid, vbDefault
    igManUnload = YES
    Unload ReRateLineBook
    igManUnload = NO
End Sub

Private Sub mSetControls()
    'frcByBook.Top = 225
    'frcAssignBook.Top = -15
    cmcDone.Top = Me.Height - cmcDone.Height - 180
    cmcCancel.Top = cmcDone.Top
    cmcClear.Top = cmcDone.Top
    'cmcCancel.Left = Me.Width / 2 - cmcCancel.Width / 2
    'cmcDone.Left = cmcCancel.Left - (3 * cmcDone.Width / 2)
    'cmcClear.Left = cmcCancel.Left + cmcCancel.Width + cmcClear.Width / 2
    'frcAssignBook.Move 130, 30
    'lacGBStart.Top = 270
    'lacGBStart.Left = 120
    'edcGBStart.Top = frcAssignBook.Top + lacGBStart.Top - 15
    'edcGBStart.Left = frcAssignBook.Left + lacGBStart.Left + lacGBStart.Width + 60
    'lacGBEnd.Top = lacGBStart.Top
    'lacGBEnd.Left = lacGBStart.Left + lacGBStart.Width + edcGBStart.Width + 120
    'edcGBEnd.Top = edcGBStart.Top
    'edcGBEnd.Left = frcAssignBook.Left + lacGBEnd.Left + lacGBEnd.Width + 60
    'lacGBVehicle.Top = lacGBStart.Top
    'lacGBVehicle.Left = lacGBEnd.Left + lacGBEnd.Width + edcGBEnd.Width + 120
    'cbcGBVehicle.Top = edcGBStart.Top
    'cbcGBVehicle.Left = frcAssignBook.Left + lacGBVehicle.Left + lacGBVehicle.Width + 60
    
    'cbcAssignBookName.Top = edcGBStart.Top + edcGBStart.Height + 105
    'cmcGetBook.Left = 120
    'cbcAssignBookName.Left = frcAssignBook.Left + cmcGetBook.Left + cmcGetBook.Width + 60
    'cmcGetBook.Top = cbcAssignBookName.Top - frcAssignBook.Top + 15
    'frcBookByGroup.Top = cmcGetBook.Top - 75
    'frcBookByGroup.Left = frcAssignBook.Width - frcBookByGroup.Width - 60
    
    'cbcContract.Top = cbcAssignBookName.Top + cbcAssignBookName.Height + 105
    'lacCntrNo.Left = 120
    'cbcContract.Left = frcAssignBook.Left + lacCntrNo.Left + lacCntrNo.Width + 60
    'lacCntrNo.Top = cbcContract.Top - frcAssignBook.Top + 30
    
    'cbcVehicle.Top = cbcContract.Top
    'cbcVehicle.Left = cbcContract.Left
    
    'lacLines.Top = lacCntrNo.Top
    'lacLines.Left = lacCntrNo.Left + lacCntrNo.Width + cbcContract.Width + 120
    
    'edcLines.Top = cbcContract.Top - frcAssignBook.Top
    'edcLines.Left = lacLines.Left + lacLines.Width + 60
    
    'cmcApply.Top = edcLines.Top + 30
    'cmcApply.Left = frcAssignBook.Width - cmcApply.Width - 120
    
    grdBookGrid.Height = cmcDone.Top - (grdBookGrid.Top + 120) - 3 * cmcDone.Height
    
    lbcApplyLog.Top = grdBookGrid.Top + grdBookGrid.Height + 30
    lbcApplyLog.Height = (Me.Height - cmcDone.Height - 120) - (grdBookGrid.Top + grdBookGrid.Height + 60) - 120 'Span between Grid and Buttons
    
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdBookGrid, fgBoxGridH + 15
    
    imcKey.Top = cmcCancel.Top
    imcKey.Left = imcKey.Width
    'lbcKey.Move imcKey.Left + imcKey.Width - lbcKey.Width, imcKey.Top + imcKey.Height
    lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
    
    mSetOptions imView
    
End Sub

Private Sub mSetOptions(ilview As imViews)
    '3/9/21 New UI
    If frcFilter.Visible = True Then
        cmcFilterBooks.Caption = "Filter Books.."
        cmcFilterBooks.Enabled = True
    End If
    
    frcByBook.Visible = False
    cbcPurchaseBook.Visible = False
    cbcBook.Visible = False
    frcFilter.Visible = False
    edcGBStart.Visible = False
    edcGBEnd.Visible = False
    cbcGBVehicle.Visible = False
    cbcAssignBookName.Visible = False
    'TTP 10143
    If frcFilterReRateBook.Visible = True Then
        cmcFilterReRateBooks.Caption = "Filter Books.."
        cmcFilterReRateBooks.Enabled = True
    End If
    frcFilterReRateBook.Visible = False
    edcGRBStart.Visible = False
    edcGRBEnd.Visible = False
    frcFilterReRateBook.Enabled = True
    
    lacLines.Visible = False
    edcLines.Visible = False
    cbcVehicle.Visible = False
    cbcContract.Visible = False
    lbcApplyOption.Visible = False
    edcLines.Visible = False
    rbcLine(0).Visible = False
    rbcLine(2).Visible = False
    rbcLine(3).Visible = False
    cmcFilterBooks.Visible = False
    lbcMap.Visible = False
    cbcContract.Left = -100
    
    Select Case ilview
        Case VIEWBYBOOK 'By Book
            frcByBook.Visible = True
            cbcPurchaseBook.Visible = True
            cbcBook.Visible = True
            lbcMap.Visible = True
            'lbcMap.Refresh
        Case VIEWBYLINE 'By Line (All Rows)
            lbcApplyOption.Caption = "Lines"
            lbcApplyOption.Visible = True
            cbcAssignBookName.Visible = True
            rbcLine(0).Visible = True
            rbcLine(2).Visible = True
            rbcLine(3).Visible = True
            cmcFilterBooks.Visible = True
        Case VIEWBYVEHICLE 'By Vehicle
            lbcApplyOption.Caption = "Vehicle"
            lbcApplyOption.Visible = True
            cbcVehicle.Visible = True
            cbcVehicle.Visible = True
            cbcAssignBookName.Visible = True
            cmcFilterBooks.Visible = True
        Case VIEWBYCNTR 'By Contract/Line
            lbcApplyOption.Caption = "Contract"
            lbcApplyOption.Visible = True
            lacLines.Visible = True
            edcLines.Visible = True
            edcLines.Visible = True
            cbcContract.Left = 9000
            cbcContract.Visible = True
            lacLines.Visible = True
            cbcAssignBookName.Visible = True
            cmcFilterBooks.Visible = True
    End Select
    
    pbcView.Visible = True
    pbcView_Paint
End Sub

'Private Sub frcAssignBook_Click()
'    mSetShow
'End Sub

Private Sub frcBookByGroup_Click()
    mSetShow
End Sub

Private Sub grdBookGrid_GotFocus()
    mSetShow
End Sub

Private Sub grdBookGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShowToolTip Button, Shift, X, Y
End Sub

Private Sub grdBookGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip Button, Shift, X, Y
End Sub

Private Sub grdBookGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String
    Dim slType As String
    Dim ilClf As Integer

    If Y < grdBookGrid.RowHeight(0) Then
        'grdBookGrid.Col = grdBookGrid.MouseCol
        'mVehSortCol grdBookGrid.Col
        'grdBookGrid.Row = 0
        'grdBookGrid.Col = PRODUCTINDEX
        Exit Sub
    End If
    llCurrentRow = grdBookGrid.MouseRow
    llCol = grdBookGrid.MouseCol
    If llCurrentRow < grdBookGrid.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdBookGrid.FixedRows Then
        If grdBookGrid.TextMatrix(llCurrentRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llCurrentRow, FLIGHTDATE) <> "CBS" Then
            ilClf = grdBookGrid.TextMatrix(llCurrentRow, CLFINDEXINDEX)
            slType = tgBookByLineAssigned(ilClf).sType
            If slType <> "O" And slType <> "A" And slType <> "E" Then
                llTopRow = grdBookGrid.TopRow
                If lmLastClickedRow <> -1 Then
                    grdBookGrid.Row = lmLastClickedRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    grdBookGrid.CellBackColor = WHITE
                End If
                'If (Shift And CTRLMASK) > 0 Then
                '    If grdBookGrid.TextMatrix(llCurrentRow, RERATEBOOKNAMEINDEX) = "" Then
                '        'grdBookGrid.TextMatrix(grdBookGrid.Row, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                '        mAssign llCurrentRow
                '    Else
                '        grdBookGrid.TextMatrix(llCurrentRow, RERATEBOOKNAMEINDEX) = ""
                '    End If
                'Else
                '    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
                '        If grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" Then
                '            ilClf = grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX)
                '            slType = tgBookByLineAssigned(ilClf).sType
                '            If slType <> "O" And slType <> "A" And slType <> "E" Then
                '                'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = ""
                '                If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                '                    If llRow = llCurrentRow Then
                '                        ''grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                '                        'grdBookGrid.Row = llRow
                '                        'grdBookGrid.Col = RERATEBOOKNAMEINDEX
                '                        'grdBookGrid.CellBackColor = GRAY
                '                        mAssign llRow
                '                        Exit For
                '                    End If
                '                ElseIf lmLastClickedRow < llCurrentRow Then
                '                    If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                '                        'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                '                        mAssign llRow
                '                    End If
                '                Else
                '                    If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                '                        'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                '                        mAssign llRow
                '                    End If
                '                End If
                '            End If
                '        End If
                '    Next llRow
                'End If
                grdBookGrid.Row = llCurrentRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                grdBookGrid.CellBackColor = GRAY
                grdBookGrid.TopRow = llTopRow
                grdBookGrid.Row = llCurrentRow
                lmLastClickedRow = llCurrentRow
                If llCol = RERATEBOOKNAMEINDEX Then
                    mEnableBox
                End If
            End If
        End If
    End If

    bmInGrid = False
End Sub

Private Sub mPopCntr()
    Dim ilChf As Integer
    Dim slFirstContract As String
    
    cbcContract.Clear
    cbcContract.SetDropDownWidth (cbcContract.Width)
    cbcContract.SetDropDownNumRows (10)
    For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
        If tgBookByLineCntr(ilChf).sSelected = "1" Then
            If slFirstContract = "" Then slFirstContract = tgBookByLineCntr(ilChf).lCntrNo
            cbcContract.AddItem (tgBookByLineCntr(ilChf).lCntrNo)
            cbcContract.SetItemData = ilChf
        End If
    Next ilChf
    If cbcContract.ListCount > 0 Then cbcContract.Text = slFirstContract
End Sub

'Private Sub mPopAssignBookNames()
'    Dim ilDnf As Integer
'    Dim ilNumberDays As Integer
'
'    'If rbcBookYears(2).Value Then
'        ilNumberDays = -1
'    'ElseIf rbcBookYears(1).Value Then
'    '    ilNumberDays = 740
'    'Else
'    '    ilNumberDays = 370
'    'End If
'    'cbcAssignBookName.Clear
'    'cbcAssignBookName.SetDropDownWidth (cbcAssignBookName.Width)
'    'cbcAssignBookName.SetDropDownNumRows (20)
'    'cbcAssignBookName.AddItem "[Vehicle Default]"
'    'cbcAssignBookName.SetItemData = -1
'    'cbcAssignBookName.AddItem "[Closest to Air Date]"
'    'cbcAssignBookName.SetItemData = -2
'    'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
'    '    cbcAssignBookName.AddItem "[Purchase Book]"
'    '    cbcAssignBookName.SetItemData = -3
'    'End If
'    'For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
'    '    If (tgBookInfo(ilDnf).lBookDate >= lgReRateStartDate - ilNumberDays) Or (ilNumberDays = -1) Then
'    '        cbcAssignBookName.AddItem (Trim$(tgBookInfo(ilDnf).sName))
'    '        cbcAssignBookName.SetItemData = ilDnf
'    '    End If
'    'Next ilDnf
'    mApplyFilter
'End Sub

Private Sub mPopLnBookNames(ilVefCode As Integer)
    Dim ilDnf As Integer
    Dim llNext As Long  '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
    
    cbcLnBookName.Clear
    cbcLnBookName.AddItem "[Vehicle Default]"
    cbcLnBookName.SetItemData = -1
    cbcLnBookName.AddItem "[Closest to Air Date]"
    cbcLnBookName.SetItemData = -2
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        cbcLnBookName.AddItem "[Purchase Book]"
        cbcLnBookName.SetItemData = -3
    End If
    
    For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
        llNext = tgBookInfo(ilDnf).lFirst
        Do While llNext <> -1
            If ilVefCode = tgBookVehicle(llNext).iVefCode Then
                cbcLnBookName.AddItem (Trim$(tgBookInfo(ilDnf).sName)) '& ":" & Format(tgBookInfo(ilDnf).lBookDate, "ddddd")
                cbcLnBookName.SetItemData = tgBookInfo(ilDnf).iDnfCode
                Exit Do
            End If
            llNext = tgBookVehicle(llNext).lNext
        Loop
    Next ilDnf
End Sub
Private Sub mPopBook()
    Dim ilDnf As Integer
    Dim llNext1 As Long 'TTP 10385 - overflow error when pressing Set Book by Line button
    Dim ilNext As Integer
    Dim ilChf As Integer
    Dim ilVef As Integer
    
    lbcBook.Clear
    cbcBook.Clear
    For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
        For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
            If tgBookByLineCntr(ilChf).sSelected = "1" Then
                ilNext = tgBookByLineCntr(ilChf).iFirst
                Do While ilNext <> -1
                    If (tgBookByLineAssigned(ilNext).sType <> "O") And (tgBookByLineAssigned(ilNext).sType <> "A") And (tgBookByLineAssigned(ilNext).sType <> "E") Then
                        ilVef = gBinarySearchVef(tgBookByLineAssigned(ilNext).iVefCode)
                        If ilVef <> -1 Then
                            llNext1 = tgBookInfo(ilDnf).lFirst
                            Do While llNext1 <> -1
                                If tgBookByLineAssigned(ilNext).iVefCode = tgBookVehicle(llNext1).iVefCode Then
                                    gFindMatch Trim$(tgBookInfo(ilDnf).sName), 0, lbcBook
                                    If gLastFound(lbcBook) < 0 Then
                                        lbcBook.AddItem (Trim$(tgBookInfo(ilDnf).sName)) '& ":" & Format(tgBookInfo(ilDnf).lBookDate, "ddddd")
                                        lbcBook.ItemData(lbcBook.NewIndex) = tgBookInfo(ilDnf).iDnfCode
                                    End If
                                    Exit Do
                                End If
                                llNext1 = tgBookVehicle(llNext1).lNext
                            Loop
                        End If
                    End If
                    ilNext = tgBookByLineAssigned(ilNext).iNext
                Loop
            End If
        Next ilChf
    Next ilDnf
    For ilDnf = 0 To lbcBook.ListCount - 1 Step 1
        cbcBook.AddItem lbcBook.List(ilDnf)
        cbcBook.SetItemData = lbcBook.ItemData(ilDnf)
    Next ilDnf
    cbcBook.SetDropDownNumRows 20
End Sub

Private Sub mPopPurchaseBook()
    Dim ilDnf As Integer
    Dim illoop As Integer
    Dim ilNext As Integer
    Dim ilVef As Integer
    
    lbcPurchaseBook.Clear
    cbcPurchaseBook.Clear
    For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
        For illoop = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
            If (tgBookByLineAssigned(illoop).sType <> "O") And (tgBookByLineAssigned(illoop).sType <> "A") And (tgBookByLineAssigned(illoop).sType <> "E") Then
                '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                If tgBookInfo(ilDnf).iDnfCode = tgBookByLineAssigned(illoop).iPurchaseDnfCode Then
                    gFindMatch Trim$(tgBookInfo(ilDnf).sName), 0, lbcPurchaseBook
                    If gLastFound(lbcPurchaseBook) < 0 Then
                        lbcPurchaseBook.AddItem (Trim$(tgBookInfo(ilDnf).sName)) '& ":" & Format(tgBookInfo(ilDnf).lBookDate, "ddddd")
                        lbcPurchaseBook.ItemData(lbcPurchaseBook.NewIndex) = tgBookInfo(ilDnf).iDnfCode
                    End If
                    Exit For
                End If
            End If
        Next illoop
    Next ilDnf
    For ilDnf = 0 To lbcPurchaseBook.ListCount - 1 Step 1
        cbcPurchaseBook.AddItem lbcPurchaseBook.List(ilDnf)
        cbcPurchaseBook.SetItemData = lbcPurchaseBook.ItemData(ilDnf)
    Next ilDnf
    cbcPurchaseBook.SetDropDownNumRows 20
End Sub


Private Sub mSetCommands()
    cmcApply.Enabled = False

'    If cbcAssignBookName.ListIndex < 0 Or cbcAssignBookName.Text = "" Then
'        Exit Sub
'    End If
'
'    If rbcLine(0).Value Then    'All
'        cmcApply.Enabled = True
'    ElseIf rbcLine(1).Value Then    'By Contract/Line
'        If cbcContract.ListIndex >= 0 And cbcContract.Text <> "" And edcLines.Text <> "" Then
'            cmcApply.Enabled = True
'        End If
'    ElseIf rbcLine(2).Value Then    'All except MG/Bonus
'        cmcApply.Enabled = True
'    ElseIf rbcLine(3).Value Then    'MG/Bonus
'        cmcApply.Enabled = True
'    ElseIf rbcLine(4).Value Then    'By Vehicle
'        If cbcVehicle.ListIndex >= 0 And cbcVehicle.Text <> "" Then
'            cmcApply.Enabled = True
'        End If
'    End If

    Select Case imView
        Case VIEWBYBOOK
            If lbcMap.ListCount > 0 Then cmcApply.Enabled = True
        
        Case VIEWBYCNTR
            If edcLines.Text <> "" And cbcContract.Text <> "" Then cmcApply.Enabled = True
            
        Case VIEWBYLINE
            If rbcLine(0).Value = True Then cmcApply.Enabled = True
            If rbcLine(3).Value = True Then cmcApply.Enabled = True
            If rbcLine(2).Value = True Then cmcApply.Enabled = True
            
        Case VIEWBYVEHICLE
            If cbcVehicle.Text <> "" Then cmcApply.Enabled = True
            
    End Select
    mSetOptions imView
    
End Sub

Private Sub grdBookGrid_RowColChange()
    mShowArrow
End Sub

Private Sub grdBookGrid_Scroll()
    mSetShow
    mShowArrow
End Sub

Private Sub imcKey_Click()
    mSetShow
    lbcKey.Visible = Not lbcKey.Visible
    If lbcKey.Visible Then
        lbcKey.ZOrder
    End If
End Sub

Private Sub lbcApplyLog_Click()
    If lbcApplyLog.ItemData(lbcApplyLog.ListIndex) <> 0 Then
        grdBookGrid.TopRow = lbcApplyLog.ItemData(lbcApplyLog.ListIndex)
        grdBookGrid.Row = lbcApplyLog.ItemData(lbcApplyLog.ListIndex)
        mShowArrow
    End If
End Sub

Private Sub lbcBook_Click()
    mSetByBookCommands
End Sub

Private Sub lbcMap_Click()
    mSetByBookCommands
End Sub

Private Sub lbcPurchaseBook_Click()
    mSetByBookCommands
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    mSetShow
End Sub

Private Sub pbcTab_GotFocus()
    mSetShow
End Sub

 Private Sub pbcView_Paint()
    
    pbcView.BackColor = &HFF0000
    pbcView.Cls
    'pbcView.CurrentX = fgBoxInsetX
    pbcView.CurrentX = 60
    pbcView.CurrentY = 10 'fgBoxInsetY
    
    Select Case imView
        Case VIEWBYBOOK 'By Book
            pbcView.Print "By Book"
        Case VIEWBYLINE 'By Line
            pbcView.Print "By Line"
        Case VIEWBYVEHICLE 'By Vehicle
            pbcView.Print "By Vehicle"
        Case VIEWBYCNTR 'By Contract/Line
            pbcView.Print "By Contract/Line"
        Case Else
            imView = VIEWBYBOOK
            pbcView.Print "By Book"
    End Select
    
End Sub

Private Sub pbcView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imView = imView + 1
    If imView > 3 Then imView = 0
    mSetOptions imView
    mSetCommands
End Sub

'Private Sub rbcByBook_Click()
'    If rbcByBook.Value Then
'        frcByBook.Visible = True
'        cbcPurchaseBook.Visible = True
'        cbcBook.Visible = True
'    End If
'End Sub

'Private Sub rbcByLine_Click()
'    If rbcByLine.Value Then
'        frcByBook.Visible = False
'        cbcPurchaseBook.Visible = False
'        cbcBook.Visible = False
'    End If
'End Sub

Private Sub rbcLine_Click(Index As Integer)
'    If rbcLine(Index).Value Then
'        edcLines.Visible = False
'        cbcContract.Visible = False
'        lacLines.Visible = False
        'lacCntrNo.Visible = False
'        cbcVehicle.Visible = False
'        Select Case Index
'            Case 0  'All
'            Case 1  'By contract/line number specified
'                edcLines.Visible = True
'                cbcContract.Visible = True
'                lacLines.Visible = True
'                'lacCntrNo.Visible = True
'                'lacCntrNo.Caption = "Contract"
'            Case 2  'All except MG/Bonus
'            Case 3  'MG/Bonus only
'            Case 4  'By Vehicle
'                cbcVehicle.Visible = True
'                'lacCntrNo.Visible = True
'                'lacCntrNo.Caption = "Vehicle"
'        End Select
'    End If
'    mSetCommands
End Sub

Public Sub mSetBookNameInGrid(llRow As Long, llCol As Long, ilDnfCode As Integer)
    Dim ilDnf As Integer
    Dim illoop As Integer
    If ilDnfCode > 0 Then
        ilDnf = -1
        For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
            If ilDnfCode = tgBookInfo(illoop).iDnfCode Then
                ilDnf = illoop
                Exit For
            End If
        Next illoop
        If ilDnf <> -1 Then
            grdBookGrid.TextMatrix(llRow, llCol) = Trim$(tgBookInfo(ilDnf).sName)
        Else
            grdBookGrid.TextMatrix(llRow, llCol) = ""
        End If
        grdBookGrid.Row = llRow
        grdBookGrid.Col = llCol
        grdBookGrid.CellAlignment = flexAlignLeftCenter

    End If

End Sub

Private Sub mMGBonusExist(llChfCode As Long, ilLineNo As Integer, llPrevNextIndex As Long)
    Dim slSQLQuery As String
    Dim rst_Sdf As ADODB.Recordset
    Dim llSvPrevNextIndex As Long
    Dim llUpper As Long
    
    llSvPrevNextIndex = llPrevNextIndex
    slSQLQuery = "SELECT Distinct sdfVefCode, Count(If(sdfSchStatus='G' and sdfSpotType <>'X', 1, Null)) as GSpots, Count(If(sdfSchStatus='O' and sdfSpotType <>'X', 1, Null)) as OSpots, Count(If(sdfSpotType='X', 1, Null)) as XSpots FROM sdf_Spot_Detail "
    slSQLQuery = slSQLQuery & "Where sdfChfCode = " & llChfCode & " And sdfLineNo = " & ilLineNo
    slSQLQuery = slSQLQuery & " And sdfDate >= '" & Format(lgReRateStartDate, sgSQLDateForm) & "'" & " And sdfDate <= '" & Format(lgReRateEndDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And (sdfSpotType = '" & "X" & "'"
    slSQLQuery = slSQLQuery & " Or sdfSchStatus = '" & "G" & "'"
    slSQLQuery = slSQLQuery & " Or sdfSchStatus = '" & "O" & "'" & ")"
    slSQLQuery = slSQLQuery & " Group By sdfVefCode"
    slSQLQuery = slSQLQuery & " Having GSpots+OSpots+XSpots > 0"
    Set rst_Sdf = gSQLSelectCall(slSQLQuery)
    Do While (Not rst_Sdf.EOF) And (Not rst_Sdf.BOF)
        If (rst_Sdf!GSpots > 0) Or (rst_Sdf!OSpots > 0) Or (rst_Sdf!XSpots > 0) Then
            llUpper = UBound(tgBookByLineAssigned)
            tgBookByLineAssigned(llUpper) = tgBookByLineAssigned(llSvPrevNextIndex)
            tgBookByLineAssigned(llUpper).iVefCode = rst_Sdf!sdfVefCode
            tgBookByLineAssigned(llUpper).iMGCount = rst_Sdf!GSpots
            tgBookByLineAssigned(llUpper).iOutsideCount = rst_Sdf!OSpots
            tgBookByLineAssigned(llUpper).iBonusCount = rst_Sdf!XSpots
            tgBookByLineAssigned(llUpper).iNext = -1
            tgBookByLineAssigned(llPrevNextIndex).iNext = llUpper
            llPrevNextIndex = llUpper
            ReDim Preserve tgBookByLineAssigned(0 To llUpper + 1) As BOOKBYLINEASSIGNED
        End If
        rst_Sdf.MoveNext
    Loop

End Sub
Private Sub mPopListKey()
    Dim llMaxWidth As Long
    lbcKey.Clear
'    lbcKey.AddItem "Steps to assigning Books to Lines"
'    lbcKey.AddItem "  1.  Select Research Book to be assigned."
'    lbcKey.AddItem "      If [Vehicle Default] selected, when assigned to lines the default "
'    lbcKey.AddItem "         vehicle book will be shown in Green."
'    lbcKey.AddItem "      If [Closest to Air Date] selected, when assigned to lines the closest"
'    lbcKey.AddItem "         book to the line first spot date will be shown in Blue.  This will not"
'    lbcKey.AddItem "         necessarily be the book that will be assigned to each spot."
'    lbcKey.AddItem "  2.  Specify method to assign."
'    lbcKey.AddItem "  3a. If 'All' selected, press Assign Books to complete assignment."
'    lbcKey.AddItem "  3b. If 'All except MG/Bonus', press Assign Books to complete assignment."
'    lbcKey.AddItem "  3c. If 'MG/Bonus only', press Assign Books to complete assignment."
'    lbcKey.AddItem "  3d. If 'By Vehicle', proceed to Steps 4d and 6."
'    lbcKey.AddItem "  3e. If 'By Contract/Line', proceed to Steps 4e, 5e and 6."
'    lbcKey.AddItem "  4d. Pick Vehicle to be assigned."
'    lbcKey.AddItem "  4e. Pick Contract to be assigned."
'    lbcKey.AddItem "  5e. Define lines to be assigned the Research books."
'    lbcKey.AddItem "        Separate each line and line range with a comma"
'    lbcKey.AddItem "        To assign to lines 3, 5, 6, 7, 8, 10 and 14. Enter as follows;"
'    lbcKey.AddItem "        3,5-8,10,14"
'    lbcKey.AddItem "  6.  Press Assign Books to complete assignment."
'    lbcKey.AddItem "  "
'    lbcKey.AddItem "Additional way to assign research book to Lines."
'    lbcKey.AddItem "  Select Research book."
'    lbcKey.AddItem "  To assign to a Line: mouse click on line book name field."
'    lbcKey.AddItem "  To assign to range of lines: mouse click on start line, then press shift key and mouse click on end line."
'    lbcKey.AddItem "  "
'    lbcKey.AddItem "To Remove assignment, press Ctrl key and mouse click on line book to be removed."
'    lbcKey.AddItem "  "
'    lbcKey.AddItem "Note(s):"
'    lbcKey.AddItem "  The book to be assigned will be checked that it contains research data for the assigned Book"
'    lbcKey.AddItem "  If the vehicle is not found, the book field will not be altered"
    
    lbcKey.AddItem "Steps to assigning Books to Lines:"
    lbcKey.AddItem "  Before you begin, Select which method to apply Research books."
    lbcKey.AddItem "    Click 'Option' to toggle through the 4 methods:"
    lbcKey.AddItem "    [By Line], [By Vehicle], [By Contract/Line], and [By Book]"
    lbcKey.AddItem ""
    lbcKey.AddItem "    --------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    lbcKey.AddItem "    [By Line], [By Vehicle], [By Contract/Line]"
    lbcKey.AddItem "    --------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    lbcKey.AddItem "    1. Select the 'Book Name' to be assigned."
    lbcKey.AddItem "        1a. You can Filter the list of the Books by clicking [Filter Books..]"
    lbcKey.AddItem "            Optionally specify Starting and or Ending Book Date(s), and/or Vehicle to Filter list."
    lbcKey.AddItem "            Click [Apply Filter]"
    lbcKey.AddItem "    2. Specify which Line(s) will be Assigned ReRate Books."
    lbcKey.AddItem "        2a. [By Line] "
    lbcKey.AddItem "            'All Lines' – Selected book will be applied to all applicable lines."
    lbcKey.AddItem "            'MG/Bonus only' – Selected book will only be applied to MG/Bonus lines."
    lbcKey.AddItem "            'All except MG/Bonus' – Selected book will not be applied to MG/Bonus lines."
    lbcKey.AddItem "        2b. [By Vehicle]"
    lbcKey.AddItem "            Pick Vehicle to be assigned."
    lbcKey.AddItem "        2c. [By Contract/Line]"
    lbcKey.AddItem "            Pick Contract to be assigned."
    lbcKey.AddItem "            Define lines to be assigned the Research books."
    lbcKey.AddItem "            Separate each line and line range with a comma."
    lbcKey.AddItem "            Example: To assign to lines 3, 5, 6, 7, 8, 10 and 14. Enter: 3,5-8,10,14"
    lbcKey.AddItem "    3. 'Assign' the ReRate books to the applicable Line(s)."
    lbcKey.AddItem "        3a. 'Retain previous assignment'"
    lbcKey.AddItem "            When checked, Assigning books will not overwrite books you've already assigned."
    lbcKey.AddItem "            Click Assign."
    lbcKey.AddItem "    4. Once completed with all Assignments, click [Ok]"
    lbcKey.AddItem ""
    lbcKey.AddItem "    --------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    lbcKey.AddItem "    [By Book]"
    lbcKey.AddItem "    --------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    lbcKey.AddItem "    1. Select a 'Purchase Book' - Any lines with Purchased with Book will apply."
    lbcKey.AddItem "    2. Select a 'ReRate Book' - This is the Book that will be assinged to applicable line(s)."
    'TTP 10155 add description for book date filter for By Book method to key
    lbcKey.AddItem "        2a. You can Filter the list of the ReRate Books by typing parts of the book names. "
    lbcKey.AddItem "            Multiple search terms can be used (separated by spaces).  "
    lbcKey.AddItem "            Example: ""SP 2021 FA"" will find all book names containing ""SP"" and ""2021"" and ""FA""."
    lbcKey.AddItem "            Note: Search is not case sensitive."
    lbcKey.AddItem "        2b. To filter the list of book names, click [Filter Books..]"
    lbcKey.AddItem "            Specify a Starting and/or a Ending Book Date."
    lbcKey.AddItem "            Click [Apply Filter]"
    lbcKey.AddItem "    3. Click [Move ->] to Add the Mapped pair of 'Purchase Book' and 'ReRate Book'."
    lbcKey.AddItem "    3. Repeat Steps 1 and 2 until you have all the desired mappings."
    lbcKey.AddItem "    4. Assign the ReRate books to the Lines."
    lbcKey.AddItem "        4a. 'Retain previous assignment'"
    lbcKey.AddItem "            When checked, Assigning books will not overwrite books you've already assigned."
    lbcKey.AddItem "            Click Assign."
    lbcKey.AddItem "    5. If you've added a incorrect Mapping, select the Mapping and Click [<- Move]. "
    lbcKey.AddItem "    6. Once completed with all Assignments, click [Ok]"
    lbcKey.AddItem ""
    lbcKey.AddItem "Additional way to assign research book to Lines:"
    lbcKey.AddItem "    To manually assign (or reassign) a contract Line's book:"
    lbcKey.AddItem "    1. Click on the contract line 'ReRate Book Name' field."
    lbcKey.AddItem "    2. Select the research book from the dropdown."
    lbcKey.AddItem ""
    lbcKey.AddItem "Notes:"
    lbcKey.AddItem "    1. The book to be assigned will be checked for valid research data."
    lbcKey.AddItem "    2. If the vehicle is not found, the book field will not be altered, a message will be logged."
    lbcKey.AddItem "    3. If [Closest to Air Date] is selected, the book name assigned will not necessarily"
    lbcKey.AddItem "       be the book that will be assigned to each spot."
    lbcKey.AddItem "    4. A summary will show if any lines requiring mapping remain."
    lbcKey.AddItem "       4b. select 'Retain previous assignment' and continue 'Applying' assignments as needed."
    lbcKey.AddItem "       4b. You can optionally Assign the remaining items."
    lbcKey.AddItem "    5. Click [Clear Books] to clear all assignments."
    lbcKey.AddItem ""
    lbcKey.AddItem "Key:"
    lbcKey.AddItem "    1. 'ReRate Book Name' field:"
    lbcKey.AddItem "       If [Vehicle Default] is selected, the book name will be shown in Green."
    lbcKey.AddItem "       If [Closest to Air Date] is selected, the book name be shown in Blue."
    lbcKey.AddItem "    2. 'Flight Dates' field:"
    lbcKey.AddItem "       'CBS' is displayed in RED.  CBS indicates Canceled Before Start."
    lbcKey.AddItem "    3. MGs/Bonus Items shown in italics"
    lbcKey.AddItem "    4. Hold Right mouse button on Line to display tooltip."
    
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontName = "Arial"
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontSize = 8
    llMaxWidth = (Traffic.pbcArial.TextWidth("  To assign to range of lines: mouse click on start line, then press shift key and mouse click on end line MMMM"))
    lbcKey.Width = llMaxWidth
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount) * 225
    lbcKey.Height = gListBoxHeight(lbcKey.ListCount, 20)
'    imcKey.Top = cmcCancel.Top
'    imcKey.Left = imcKey.Width
'    'lbcKey.Move imcKey.Left + imcKey.Width - lbcKey.Width, imcKey.Top + imcKey.Height
'    lbcKey.Move imcKey.Left, imcKey.Top - imcKey.Height
End Sub

Private Sub mAssign(llRow As Long)
    Dim ilClf As Integer
    Dim slType As String
    Dim ilChf As Integer
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim blMatchRules As Boolean
    Dim slBookName As String
    lbcApplyLog.Clear
    
    ilClf = Val(grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX))
    slType = tgBookByLineAssigned(ilClf).sType
    If slType = "O" Or slType = "A" Or slType = "E" Then
        Exit Sub
    End If
    blMatchRules = False
    
    If imView = VIEWBYLINE And rbcLine(0).Value = True Then  'All
        'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
        blMatchRules = True
    ElseIf imView = VIEWBYCNTR Then     'By Contract/Line
        'If contract/line not defined, then function as if (*) All selected
        If (cbcContract.ListIndex < 0 Or cbcContract.Text = "") And edcLines.Text = "" Then
            'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
            blMatchRules = True
        Else
            'If contract specified but not lines, assign only to that contract
            ilClf = Val(grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX))
            If (cbcContract.ListIndex >= 0 Or cbcContract.Text <> "") Then
                ilChf = cbcContract.GetItemData(cbcContract.ListIndex)
                llChfCode = tgBookByLineCntr(ilChf).lChfCode
            End If
            If (cbcContract.ListIndex >= 0 Or cbcContract.Text <> "") And edcLines.Text = "" Then
                If tgBookByLineAssigned(ilClf).lChfCode = llChfCode Then
                    'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                    blMatchRules = True
                End If
            Else
                If (cbcContract.ListIndex >= 0 Or cbcContract.Text <> "") And edcLines.Text <> "" Then
                    If tgBookByLineAssigned(ilClf).lChfCode = llChfCode Then
                        'Check line
                        ilLineNo = grdBookGrid.TextMatrix(llRow, LINENOINDEX)
                        If InStr(1, smLineNo, "," & ilLineNo & ",") > 0 Then
                            'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                            blMatchRules = True
                        End If
                    End If
                End If
            End If
        End If
    ElseIf imView = VIEWBYLINE And rbcLine(2).Value = True Then    'All except MG/Bonus
        If (tgBookByLineAssigned(ilClf).iMGCount <= 0) And (tgBookByLineAssigned(ilClf).iOutsideCount <= 0) And (tgBookByLineAssigned(ilClf).iBonusCount <= 0) Then
            'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
            blMatchRules = True
        End If
    ElseIf imView = VIEWBYLINE And rbcLine(3).Value = True Then    'MG/Bonus
        If (tgBookByLineAssigned(ilClf).iMGCount > 0) Or (tgBookByLineAssigned(ilClf).iOutsideCount > 0) Or (tgBookByLineAssigned(ilClf).iBonusCount > 0) Then
            'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
            blMatchRules = True
        End If
    ElseIf imView = VIEWBYVEHICLE Then      'Vehicle
        'If vehicle not defined, then function as if (*) All selected
        If (cbcVehicle.ListIndex < 0 Or cbcVehicle.Text = "") And edcLines.Text = "" Then
            'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
            blMatchRules = True
        Else
            'If contract specified but not lines, assign only to that contract
            ilClf = Val(grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX))
            If (cbcVehicle.ListIndex >= 0 Or cbcVehicle.Text <> "") Then
                If tgBookByLineAssigned(ilClf).iVefCode = cbcVehicle.GetItemData(cbcVehicle.ListIndex) Then
                    'grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = cbcAssignBookName.Text
                    blMatchRules = True
                End If
            End If
        End If
    End If
    If blMatchRules Then
        ilClf = Val(grdBookGrid.TextMatrix(llRow, CLFINDEXINDEX))
        slBookName = mVehicleInBook(tgBookByLineAssigned(ilClf).iPurchaseDnfCode, tgBookByLineAssigned(ilClf).iVefCode, ilClf)  'tgBookByLineAssigned(ilClf).lStartDate)
        If slBookName <> "" Then
            If (ckcDontOverwriteByLine.Value = 1 And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Or ckcDontOverwriteByLine.Value = 0 Then '3/3/21 - Bonus improvements: "don't overwrite previously assigned lines"
                grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = slBookName
                If cbcAssignBookName.Text = "[Vehicle Default]" Then
                    grdBookGrid.Row = llRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    grdBookGrid.CellForeColor = DARKGREEN
                ElseIf cbcAssignBookName.Text = "[Closest to Air Date]" Then
                    grdBookGrid.Row = llRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    grdBookGrid.CellForeColor = BLUE
                ElseIf cbcAssignBookName.Text = "[Purchase Book]" Then
                    grdBookGrid.Row = llRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    grdBookGrid.CellForeColor = ORANGE
                Else
                    grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = slBookName
                    grdBookGrid.Row = llRow
                    grdBookGrid.Col = RERATEBOOKNAMEINDEX
                    grdBookGrid.CellForeColor = BLACK
                End If
                If grdBookGrid.TextMatrix(llRow, FLIGHTDATE) <> "CBS" Then imApplied = imApplied + 1
            End If
        Else
            If (ckcDontOverwriteByLine.Value = 1 And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Or ckcDontOverwriteByLine.Value = 0 Then '3/3/21 - Bonus improvements: "don't overwrite previously assigned lines"
                grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = slBookName
                grdBookGrid.Row = llRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                grdBookGrid.CellForeColor = BLACK
                lbcApplyLog.AddItem "No matching Book found for Vehicle:'" & Trim(grdBookGrid.TextMatrix(llRow, VEHICLEINDEX)) & "'.  Contract #" & smCurrentCntrNo & ", Line:" & Trim(grdBookGrid.TextMatrix(llRow, LINENOINDEX))
                lbcApplyLog.ItemData(lbcApplyLog.NewIndex) = llRow
                grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX) = "L"
            End If
        End If
        grdBookGrid.CellAlignment = flexAlignLeftCenter
    End If
End Sub

Private Function mGetLineNo() As String
    Dim slValues() As String
    Dim ilValue As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilNumStart As Integer
    Dim ilNumEnd As Integer
    Dim illoop As Integer
    Dim slvalue As String
    Dim sllines As String
    On Error GoTo mGetError
    
    slvalue = edcLines.Text
    slValues = Split(slvalue, ",")
    If Not IsArray(slValues) Then
        mGetLineNo = ""
        Exit Function
    End If
    sllines = ""
    For ilValue = LBound(slValues) To UBound(slValues) Step 1
        slStr = slValues(ilValue)
        ilPos = InStr(1, slValues(ilValue), "-")
        If ilPos <= 0 Then
            sllines = sllines & "," & slStr
        Else
            ilNumStart = Left(slStr, ilPos - 1)
            ilNumEnd = Mid(slStr, ilPos + 1)
            For illoop = ilNumStart To ilNumEnd Step 1
                sllines = sllines & "," & illoop
            Next illoop
        End If
    Next ilValue
    sllines = sllines & ","
    mGetLineNo = sllines
    Exit Function
mGetError:
    mGetLineNo = ""

End Function

Private Sub mPopVehicle()
    Dim ilClf As Integer
    Dim ilChf As Integer
    Dim ilNext As Integer
    Dim ilVef As Integer
    Dim slName As String
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim ilOk  As Integer
    Dim slFirstVehicle As String
    
    lbcVehicle.Clear
    For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
        If tgBookByLineCntr(ilChf).sSelected = "1" Then
            ilNext = tgBookByLineCntr(ilChf).iFirst
            Do While ilNext <> -1
                If (tgBookByLineAssigned(ilNext).sType <> "O") And (tgBookByLineAssigned(ilNext).sType <> "A") And (tgBookByLineAssigned(ilNext).sType <> "E") Then
                    ilVef = gBinarySearchVef(tgBookByLineAssigned(ilNext).iVefCode)
                    If ilVef <> -1 Then
                        ilOk = True
                        '2/10/21 - Test if PodCast vehicle has Programming defined
                        'If tgBookByLineAssigned(ilNext).sType = "P" Or tgBookByLineAssigned(ilNext).sType = "C" Then  'Package or Conventional
                        ilVpf = gBinarySearchVpf(tgBookByLineAssigned(ilNext).iVefCode)
                        If tgVpf(ilVpf).sGMedium = "P" Then
                            If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER) Then
                                ilOk = gExistLtf(tgBookByLineAssigned(ilNext).iVefCode)
                            End If
                        End If
                        'End If
                        If ilOk Then
                            slName = Trim$(tgMVef(ilVef).sName)
                            ilVefCode = tgBookByLineAssigned(ilNext).iVefCode
                            gFindMatch slName, 0, lbcVehicle
                            If gLastFound(lbcVehicle) < 0 Then
                                lbcVehicle.AddItem slName
                                lbcVehicle.ItemData(lbcVehicle.NewIndex) = ilVefCode
                            End If
                        End If
                    End If
                End If
                ilNext = tgBookByLineAssigned(ilNext).iNext
            Loop
        End If
    Next ilChf
    
    cbcVehicle.Clear
    cbcGBVehicle.Clear
    
    cbcVehicle.SetDropDownWidth (cbcVehicle.Width)
    cbcVehicle.SetDropDownNumRows (10)
    
    For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
        If ilVef = 0 Then slFirstVehicle = lbcVehicle.List(ilVef)
        cbcVehicle.AddItem (lbcVehicle.List(ilVef))
        cbcVehicle.SetItemData = lbcVehicle.ItemData(ilVef)
        cbcGBVehicle.AddItem (lbcVehicle.List(ilVef))
        cbcGBVehicle.ItemData(cbcGBVehicle.NewIndex) = lbcVehicle.ItemData(ilVef)
    Next ilVef
    cbcGBVehicle.AddItem "[All Vehicles]", 0
    cbcGBVehicle.ItemData(cbcGBVehicle.NewIndex) = 0
    cbcGBVehicle.Text = "[All Vehicles]"
    If cbcVehicle.ListCount > 0 Then cbcVehicle.Text = slFirstVehicle
End Sub

Private Function mVehicleInBook(ilPurchaseDnfCode As Integer, ilVefCode As Integer, ilClf As Integer) As String ' As Long) As String
    Dim llNext As Long
    Dim ilDnf As Integer
    Dim illoop As Integer
    Dim ilVef As Integer
    Dim slBookName As String
    Dim ilDnfCode As Integer
    
    mVehicleInBook = ""
    ilDnf = -1
    If cbcAssignBookName.Text = "[Vehicle Default]" Then
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
                If tgBookInfo(illoop).iDnfCode = tgMVef(ilVef).iDnfCode Then
                    slBookName = Trim$(tgBookInfo(illoop).sName)
                    ilDnf = illoop
                    Exit For
                End If
            Next illoop
        End If
    ElseIf cbcAssignBookName.Text = "[Closest to Air Date]" Then
        'If lgReRateStartDate < llDate Then
            ilDnf = mFindClosestBook(ilVefCode, ilClf)  'llDate)
        'Else
        '    ilDnf = mFindClosestBook(ilVefCode, lgReRateStartDate)
        'End If
        
        If ilDnf <> -1 Then
            slBookName = Trim$(tgBookInfo(ilDnf).sName)
        End If
    ElseIf cbcAssignBookName.Text = "[Purchase Book]" Then
        For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
            If tgBookInfo(illoop).iDnfCode = ilPurchaseDnfCode Then
                slBookName = Trim$(tgBookInfo(illoop).sName)
                ilDnf = illoop
                Exit For
            End If
        Next illoop
    Else
        If cbcAssignBookName.ListIndex = -1 Then Exit Function
        ilDnfCode = cbcAssignBookName.GetItemData(cbcAssignBookName.ListIndex)
        slBookName = Trim$(cbcAssignBookName.Text)
        For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
            If tgBookInfo(illoop).iDnfCode = ilDnfCode Then
                slBookName = Trim$(tgBookInfo(illoop).sName)
                ilDnf = illoop
                Exit For
            End If
        Next illoop
    End If
    If ilDnf = -1 Then
        Exit Function
    End If
    llNext = tgBookInfo(ilDnf).lFirst
    Do While llNext <> -1
        If ilVefCode = tgBookVehicle(llNext).iVefCode Then
            mVehicleInBook = slBookName
            Exit Do
        End If
        llNext = tgBookVehicle(llNext).lNext
    Loop
End Function

Function mFindClosestBook(ilVefCode As Integer, ilClf As Integer) As Integer    ' As Long) As Integer
    Dim ilDnf As Integer
    Dim llNext As Long
    Dim llDate As Long
    Dim slSQLQuery As String
    Dim sdf_rst As ADODB.Recordset
    
    mFindClosestBook = -1
    slSQLQuery = "Select sdfDate from SDF_Spot_Detail Where sdfVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " And sdfChfCode = " & tgBookByLineAssigned(ilClf).lChfCode
    slSQLQuery = slSQLQuery & " And sdfLineNo = " & tgBookByLineAssigned(ilClf).iLineNo
    slSQLQuery = slSQLQuery & " And sdfFsfCode = " & 0
    slSQLQuery = slSQLQuery & " And sdfDate >= '" & Format(smRangeStartDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And sdfDate <= '" & Format(smRangeEndDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " Order By sdfDate"
    Set sdf_rst = gSQLSelectCall(slSQLQuery)
    If Not sdf_rst.EOF Then
        llDate = gDateValue(sdf_rst!sdfDate)
        'Books are in descending date order
        For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
            If tgBookInfo(ilDnf).lBookDate <= llDate Then
                If gBinarySearchExcludeBook(tgBookInfo(ilDnf).iDnfCode) = -1 Then
                    llNext = tgBookInfo(ilDnf).lFirst
                    Do While llNext <> -1
                        If tgBookVehicle(llNext).iVefCode = ilVefCode Then
                            mFindClosestBook = ilDnf
                            Exit Function
                        End If
                        llNext = tgBookVehicle(llNext).lNext
                    Loop
                End If
            End If
        Next ilDnf
    End If
End Function

Private Sub mApplyFilter(Optional ilwhichDropdown As Integer = 0)
    'TTP 10143
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim llDate As Long
    Dim blOk As Boolean
    Dim slSQLQuery As String
    Dim dnf_rst As ADODB.Recordset
    Dim drf_rst As ADODB.Recordset
    
'    If (lmFilterStartDate = 0) And (lmFilterEndDate = 0) And (imFilterVefCode <= 0) Then
'        For ilLoop = 2 To cbcAssignBookName.ListCount - 1 Step 1
'            slNameCode = tmBNCode(ilLoop - 2).sKey    'lbcBNCode.List(ilIndex - 1)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            cbcAssignBookName.ItemData(ilLoop) = Val(clCode)
'        Next ilLoop
'        cbcAssignBookName.ItemData(0) = -1
'        cbcAssignBookName.ItemData(1) = -1
'        Exit Sub
'    End If

    If ilwhichDropdown = 0 Then
        cbcAssignBookName.Clear
        
        cbcAssignBookName.SetDropDownWidth (cbcAssignBookName.Width)
        cbcAssignBookName.SetDropDownNumRows (20)
        cbcAssignBookName.AddItem "[Vehicle Default]"
        cbcAssignBookName.SetItemData = -1
        cbcAssignBookName.AddItem "[Closest to Air Date]"
        cbcAssignBookName.SetItemData = -2
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            cbcAssignBookName.AddItem "[Purchase Book]"
            cbcAssignBookName.SetItemData = -3
        End If
    End If
    If ilwhichDropdown = 1 Then
        cbcBook.Clear
        cbcBook.SetDropDownWidth (cbcAssignBookName.Width)
        cbcBook.SetDropDownNumRows (20)
    End If
    
    slSQLQuery = "Select dnfCode, dnfBookName, dnfBookDate from DNF_Demo_Rsrch_Names "
    If lmFilterStartDate > 0 Then
        slSQLQuery = slSQLQuery & " Where dnfBookDate >= '" & Format(smFilterStartDate, sgSQLDateForm) & "'"
        If lmFilterEndDate > 0 Then
            slSQLQuery = slSQLQuery & " And dnfBookDate <= '" & Format(smFilterEndDate, sgSQLDateForm) & "'"
        End If
    ElseIf lmFilterEndDate > 0 Then
        slSQLQuery = slSQLQuery & " Where dnfBookDate <= '" & Format(smFilterEndDate, sgSQLDateForm) & "'"
    End If
    slSQLQuery = slSQLQuery & " Order By dnfBookDate Desc, dnfBookName"
    Set dnf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not dnf_rst.EOF
        blOk = True
        If imFilterVefCode > 0 Then
            slSQLQuery = "Select drfVefCode from DRF_Demo_Rsrch_Data where drfDnfCode = " & dnf_rst!dnfCode
            slSQLQuery = slSQLQuery & " And drfVefCode = " & imFilterVefCode
            slSQLQuery = slSQLQuery & " And drfDemoDataType = '" & "D" & "'"
            Set drf_rst = gSQLSelectCall(slSQLQuery)
            If drf_rst.EOF Then
                blOk = False
            End If
        End If
        If blOk Then
            If ilwhichDropdown = 0 Then
                cbcAssignBookName.AddItem Trim$(dnf_rst!dnfBookName) & ":" & dnf_rst!dnfBookDate
                cbcAssignBookName.SetItemData = dnf_rst!dnfCode
            End If
            If ilwhichDropdown = 1 Then
                cbcBook.AddItem Trim$(dnf_rst!dnfBookName)
                cbcBook.SetItemData = dnf_rst!dnfCode
            End If
            
        End If
        dnf_rst.MoveNext
    Loop
    'cbcAssignBookName.ListIndex = 0
    
    If ilwhichDropdown = 0 Then
        If cbcAssignBookName.ListCount < 20 Then
            'gSetComboboxDropdownHeight ReRateLineBook, cbcAssignBookName, cbcAssignBookName.ListCount
            cbcAssignBookName.SetDropDownNumRows cbcAssignBookName.ListCount
        Else
            'gSetComboboxDropdownHeight ReRateLineBook, cbcAssignBookName, 20
            cbcAssignBookName.SetDropDownNumRows 20
        End If
    End If
        
    If ilwhichDropdown = 1 Then
        If cbcBook.ListCount < 20 Then
            'gSetComboboxDropdownHeight ReRateLineBook, cbcAssignBookName, cbcAssignBookName.ListCount
            cbcBook.SetDropDownNumRows cbcBook.ListCount
        Else
            'gSetComboboxDropdownHeight ReRateLineBook, cbcAssignBookName, 20
            cbcBook.SetDropDownNumRows 20
        End If
    End If
    
End Sub

Private Sub mDetermineDateRange()
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    
    If ExptReRate.rbcDatesBy(0).Value And ExptReRate.edcDate(0).Text <> "" Then 'By Week
        slStartDate = Format$(ExptReRate.edcDate(0).Text, "m/d/yy")               'reformat date to insure year is there
        slEndDate = DateAdd("d", 6, slStartDate)
    ElseIf ExptReRate.rbcDatesBy(1).Value And (ExptReRate.edcStart.Text <> "") And (ExptReRate.edcYear.Text <> "") Then  'By Month
        slStr = ExptReRate.edcStart.Text & "/15/" & ExptReRate.edcYear.Text
        slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
        slEndDate = gObtainEndStd(slStartDate)
    ElseIf ExptReRate.rbcDatesBy(2).Value And (ExptReRate.edcStart.Text <> "") And (ExptReRate.edcYear.Text <> "") Then    'By Quarter
        If ExptReRate.edcStart.Text = 2 Then
            slStr = "4/15/" & ExptReRate.edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "6/15/" & ExptReRate.edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        ElseIf ExptReRate.edcStart.Text = 3 Then
            slStr = "7/15/" & ExptReRate.edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "9/15/" & ExptReRate.edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        ElseIf ExptReRate.edcStart.Text = 4 Then
            slStr = "10/15/" & ExptReRate.edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "12/15/" & ExptReRate.edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        Else
            slStr = "1/15/" & ExptReRate.edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "3/15/" & ExptReRate.edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        End If
    ElseIf ExptReRate.rbcDatesBy(3).Value And ExptReRate.edcDate(0).Text <> "" Then   'By Contract
        slStartDate = Format$(ExptReRate.edcDate(0).Text, "m/d/yy")
        slEndDate = "12/31/2069"
    ElseIf ExptReRate.rbcDatesBy(4).Value And (ExptReRate.edcDate(1).Text <> "") And ExptReRate.edcDate(2).Text <> "" Then   'By Range
        slStartDate = Format$(ExptReRate.edcDate(1).Text, "m/d/yy")
        slEndDate = Format$(ExptReRate.edcDate(2).Text, "m/d/yy")
    Else
        slStartDate = "1/1/1970"
        slEndDate = "12/31/2069"
    End If
    smRangeStartDate = slStartDate
    smRangeEndDate = slEndDate
End Sub

Private Sub mEnableBox()
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilClf As Integer
    Dim slType As String
    
    ilClf = grdBookGrid.TextMatrix(grdBookGrid.Row, CLFINDEXINDEX)
    slType = tgBookByLineAssigned(ilClf).sType
    If slType = "O" Or slType = "A" Or slType = "E" Then
        Exit Sub
    End If
    
    If (grdBookGrid.Row >= grdBookGrid.FixedRows) And (grdBookGrid.Row < grdBookGrid.Rows) And (grdBookGrid.Col >= RERATEBOOKNAMEINDEX) And (grdBookGrid.Col < grdBookGrid.Cols - 1) Then
        lmEnableRow = grdBookGrid.Row
        lmEnableCol = grdBookGrid.Col

        Select Case grdBookGrid.Col
            Case RERATEBOOKNAMEINDEX
                mPopLnBookNames tgBookByLineAssigned(ilClf).iVefCode
                cbcLnBookName.Move grdBookGrid.Left + grdBookGrid.ColPos(grdBookGrid.Col) - grdBookGrid.ColWidth(grdBookGrid.Col) / 2 + 30 - GRIDSCROLLWIDTH, grdBookGrid.Top + grdBookGrid.RowPos(grdBookGrid.Row) + 15, (3 * grdBookGrid.ColWidth(grdBookGrid.Col) / 2) - 30, grdBookGrid.RowHeight(grdBookGrid.Row) - 15
                cbcLnBookName.SetDropDownWidth cbcLnBookName.Width
                slStr = grdBookGrid.TextMatrix(lmEnableRow, lmEnableCol)
                If slStr = "" Then
                    slStr = Trim$(cbcLnBookName.GetName(0))
                End If
                cbcLnBookName.PopUpListDirection "B"
                cbcLnBookName.ZOrder vbBringToFront
                cbcLnBookName.Visible = True  'Set visibility
                cbcLnBookName.SelText (Trim(slStr))
                cbcLnBookName.SetFocus
        End Select
    End If
End Sub

Private Sub mSetShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llColor As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    If (lmEnableRow >= grdBookGrid.FixedRows) And (lmEnableRow < grdBookGrid.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        llSvRow = grdBookGrid.Row
        llSvCol = grdBookGrid.Col
        Select Case lmEnableCol
            Case RERATEBOOKNAMEINDEX
                cbcLnBookName.Visible = False  'Set visibility
                slStr = cbcLnBookName.Text
                If slStr = "" Then
                    If cbcLnBookName.ListCount > 0 Then
                        cbcLnBookName.Text = cbcLnBookName.GetName(0)
                        slStr = cbcLnBookName.Text
                    End If
                End If
                If Trim$(slStr) = "[Vehicle Default]" Then
                    llColor = DARKGREEN
                ElseIf Trim$(slStr) = "[Closest to Air Date]" Then
                    llColor = BLUE
                ElseIf Trim$(slStr) = "[Purchase Book]" Then
                    llColor = ORANGE
                Else
                    llColor = BLACK
                End If
                grdBookGrid.Row = lmEnableRow
                grdBookGrid.Col = RERATEBOOKNAMEINDEX
                grdBookGrid.CellForeColor = llColor
                grdBookGrid.CellAlignment = flexAlignLeftCenter
                grdBookGrid.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(slStr)
                If cbcLnBookName.ListIndex >= 0 Then
                    grdBookGrid.TextMatrix(lmEnableRow, DNFCODEINDEX) = cbcLnBookName.GetItemData(cbcLnBookName.ListIndex)
                Else
                    grdBookGrid.TextMatrix(lmEnableRow, DNFCODEINDEX) = 0
                End If
                grdBookGrid.TextMatrix(lmEnableRow, ASSIGNMETHODINDEX) = "M"
                
        End Select
        If lmLastClickedRow <> -1 Then
            grdBookGrid.Row = lmLastClickedRow
            grdBookGrid.Col = RERATEBOOKNAMEINDEX
            grdBookGrid.CellBackColor = WHITE
            lmLastClickedRow = -1
        End If
        grdBookGrid.Row = llSvRow
        grdBookGrid.Col = llSvCol
        lmEnableCol = -1
        lmEnableRow = -1
        bmInGrid = False
    End If
End Sub

Private Sub rbcLine_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub mReassignBooksToLines()
    Dim llRow As Long
    Dim slType As String
    Dim ilClf As Integer
    Dim ilDnf As Integer
    Dim llChfCode As Long
    smCurrentCntrNo = ""
    llChfCode = -1
    For llRow = grdBookGrid.FixedRows To grdBookGrid.Rows - 1 Step 1
        If grdBookGrid.TextMatrix(llRow, LINENOINDEX) <> "" And grdBookGrid.TextMatrix(llRow, PURCHASEBOOKNAMEINDEX) <> "" Then
            If (grdBookGrid.TextMatrix(llRow, CNTRNOINDEX) <> "") Then smCurrentCntrNo = grdBookGrid.TextMatrix(llRow, CNTRNOINDEX)
            For ilClf = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                slType = tgBookByLineAssigned(ilClf).sType
                If slType <> "O" And slType <> "A" And slType <> "E" Then
                    llChfCode = tgBookByLineCntr(Val(grdBookGrid.TextMatrix(llRow, CHFINDEXINDEX))).lChfCode
                    If llChfCode = tgBookByLineAssigned(ilClf).lChfCode Then
                        If (Val(grdBookGrid.TextMatrix(llRow, LINENOINDEX)) = tgBookByLineAssigned(ilClf).iLineNo) And (Trim$(tgBookByLineAssigned(ilClf).sDaypartName) = Trim(grdBookGrid.TextMatrix(llRow, DAYPARTINDEX))) Then
                            grdBookGrid.TextMatrix(llRow, ASSIGNMETHODINDEX) = tgBookByLineAssigned(ilClf).sAssignMethod
                            If tgBookByLineAssigned(ilClf).iReRateDnfCode = -2 And tgBookByLineAssigned(ilClf).sAssignMethod <> "B" Then
                                mAssign llRow
                            Else
                                For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
                                    '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
                                    If tgBookInfo(ilDnf).iDnfCode = tgBookByLineAssigned(ilClf).iReRateDnfCode Then
                                        If (ckcDontOverwriteByLine.Value = 1 And grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = "") Or ckcDontOverwriteByLine.Value = 0 Then '3/3/21 - Bonus improvements: "don't overwrite previously assigned lines"
                                            grdBookGrid.TextMatrix(llRow, RERATEBOOKNAMEINDEX) = Trim$(tgBookInfo(ilDnf).sName)
                                        End If
                                        Exit For
                                    End If
                                Next ilDnf
                            End If
                            grdBookGrid.Row = llRow
                            grdBookGrid.Col = RERATEBOOKNAMEINDEX
                            grdBookGrid.CellForeColor = tgBookByLineAssigned(ilClf).lColor
                            Exit For
                        End If
                    End If
                End If
            Next ilClf
        End If
    Next llRow
End Sub

Private Sub mMoveName(slDirection As String)
    Dim slName As String
    Dim slItemData As String
    Dim ilPos As Integer
    
    If slDirection = "ToMap" Then
        'If lbcPurchaseBook.ListIndex >= 0 And lbcBook.ListIndex >= 0 Then
        If cbcPurchaseBook.ListIndex >= 0 And cbcBook.ListIndex >= 0 Then
            'slName = lbcPurchaseBook.List(lbcPurchaseBook.ListIndex) & "->" & lbcBook.List(lbcBook.ListIndex)
            'slItemData = lbcPurchaseBook.ItemData(lbcPurchaseBook.ListIndex)
            slName = cbcPurchaseBook.GetName(cbcPurchaseBook.ListIndex) & "->" & cbcBook.GetName(cbcBook.ListIndex)
            slItemData = cbcPurchaseBook.GetItemData(cbcPurchaseBook.ListIndex)
            lbcMap.AddItem slName
            lbcMap.ItemData(lbcMap.NewIndex) = slItemData
            'lbcPurchaseBook.RemoveItem lbcPurchaseBook.ListIndex
            'lbcBook.ListIndex = -1
            cbcPurchaseBook.RemoveListIndex = cbcPurchaseBook.ListIndex
            'cbcPurchaseBook.Text = ""
            cbcBook.Text = ""
        End If
    ElseIf slDirection = "FromMap" Then
        If lbcMap.ListIndex >= 0 Then
            slName = lbcMap.List(lbcMap.ListIndex)
            slItemData = lbcMap.ItemData(lbcMap.ListIndex)
            ilPos = InStr(1, slName, "->")
            If ilPos > 0 Then
                slName = Left(slName, ilPos - 1)
                'lbcPurchaseBook.AddItem slName
                'lbcPurchaseBook.ItemData(lbcPurchaseBook.NewIndex) = slItemData
                cbcPurchaseBook.AddItem slName
                cbcPurchaseBook.SetItemData = slItemData
                lbcMap.RemoveItem lbcMap.ListIndex
            End If
        End If
    End If
    mSetByBookCommands
End Sub

Private Sub mSetByBookCommands()
    cmcToMap.Enabled = True
    'If lbcPurchaseBook.ListIndex >= 0 And lbcBook.ListIndex >= 0 Then
    If cbcPurchaseBook.ListIndex < 0 Or cbcBook.ListIndex < 0 Then
        cmcToMap.Enabled = False
    End If
    If lbcMap.ListIndex < 0 Then
        cmcFromMap.Enabled = False
    End If
    If lbcMap.ListCount > 0 Then
        cmcAssign.Enabled = True
    Else
        cmcAssign.Enabled = False
    End If
End Sub

Private Sub mShowArrow()
    pbcArrow.Left = grdBookGrid.Left - pbcArrow.Width
    pbcArrow.Top = grdBookGrid.Top + grdBookGrid.RowPos(grdBookGrid.Row)
    If grdBookGrid.TopRow <= grdBookGrid.Row And (pbcArrow.Top + grdBookGrid.RowHeight(0)) < (grdBookGrid.Top + grdBookGrid.Height) Then
        pbcArrow.Visible = True
        grdBookGrid_GotFocus
    Else
        pbcArrow.Visible = False
    End If
End Sub

Sub ShowToolTip(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llCurrentRow As Long
    Dim illoop As Integer
    Dim slStr As String
    
    If Button <> 2 Then
        pbcTooltip.Visible = False
        pbcDropShadow.Visible = False
        Exit Sub
    End If
    If Y < grdBookGrid.RowHeight(0) Then
        Exit Sub
    End If
    llCurrentRow = grdBookGrid.MouseRow
    If llCurrentRow < grdBookGrid.FixedRows Then
        Exit Sub
    End If
    If grdBookGrid.TextMatrix(llCurrentRow, VEHICLEINDEX) = "" Then
        pbcTooltip.Visible = False
        pbcDropShadow.Visible = False
        Exit Sub
    End If
    cbcLnBookName.Visible = False
    
    For illoop = grdBookGrid.FixedRows To grdBookGrid.Rows - 1
        If Trim(grdBookGrid.TextMatrix(illoop, CNTRNOINDEX)) <> "" Then slStr = grdBookGrid.TextMatrix(illoop, CNTRNOINDEX)
        If illoop >= llCurrentRow Then Exit For
    Next illoop
    
    If llCurrentRow >= grdBookGrid.FixedRows Then
        lbcTooltip(0).Caption = "Contract:" & slStr
        lbcTooltip(1).Caption = "Line:" & grdBookGrid.TextMatrix(llCurrentRow, LINENOINDEX)
        lbcTooltip(2).Caption = "Daypart:" & grdBookGrid.TextMatrix(llCurrentRow, DAYPARTINDEX)
        lbcTooltip(3).Caption = "Vehicle:" & grdBookGrid.TextMatrix(llCurrentRow, VEHICLEINDEX)
        
        If grdBookGrid.Left + X + pbcTooltip.Width < grdBookGrid.Width + grdBookGrid.Left Then
            pbcTooltip.Left = grdBookGrid.Left + X
            pbcTooltip.Top = grdBookGrid.Top + Y - pbcTooltip.Height - 10
        Else
            pbcTooltip.Left = grdBookGrid.Width + grdBookGrid.Left - pbcTooltip.Width
            pbcTooltip.Top = grdBookGrid.Top + Y - pbcTooltip.Height - 120
        End If
        
        pbcDropShadow.Top = pbcTooltip.Top + 30
        pbcDropShadow.Left = pbcTooltip.Left + 30
        pbcTooltip.Visible = True
        pbcDropShadow.Visible = True
    End If
End Sub

