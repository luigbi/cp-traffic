VERSION 5.00
Begin VB.Form Research 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6105
   ClientLeft      =   3900
   ClientTop       =   3555
   ClientWidth     =   9360
   ControlBox      =   0   'False
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
   Icon            =   "Research.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   9360
   Begin VB.PictureBox plcACT1Settings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   5760
      Picture         =   "Research.frx":08CA
      ScaleHeight     =   930
      ScaleWidth      =   1950
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1900
      Visible         =   0   'False
      Width           =   2010
      Begin VB.TextBox edcACT1SettingF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "No"
         Top             =   620
         Width           =   705
      End
      Begin VB.TextBox edcACT1SettingC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "No"
         Top             =   620
         Width           =   705
      End
      Begin VB.TextBox edcACT1SettingS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "No"
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox edcACT1SettingT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "No"
         Top             =   210
         Width           =   735
      End
   End
   Begin V81Research.CSI_ComboBoxMS CSI_ComboBoxMS1 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmcBaseDuplicate 
      Appearance      =   0  'Flat
      Caption         =   "Base Duplicate"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3015
      TabIndex        =   87
      Top             =   4920
      Width           =   1470
   End
   Begin VB.CommandButton cmcDuplicate 
      Appearance      =   0  'Flat
      Caption         =   "Row Duplicate"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   86
      Top             =   4905
      Width           =   1470
   End
   Begin VB.CommandButton cmcGetBook 
      Appearance      =   0  'Flat
      Caption         =   "Get Books"
      Height          =   285
      Left            =   6510
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox cbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4875
      TabIndex        =   3
      Top             =   15
      Width           =   1530
   End
   Begin V81Research.CSI_Calendar edcEnd 
      Height          =   285
      Left            =   3495
      TabIndex        =   2
      Top             =   15
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      Text            =   "12/20/2023"
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   0
   End
   Begin V81Research.CSI_Calendar edcStart 
      Height          =   285
      Left            =   2025
      TabIndex        =   1
      Top             =   15
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      Text            =   "12/20/2023"
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.CommandButton cmcAdjust 
      Appearance      =   0  'Flat
      Caption         =   "&Adjust"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   83
      Top             =   5310
      Width           =   1050
   End
   Begin VB.ComboBox cbcDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6030
      TabIndex        =   82
      Top             =   4905
      Width           =   3165
   End
   Begin VB.ListBox lbcPopSrce 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1125
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   3345
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   2805
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   30
         Picture         =   "Research.frx":67FC
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   30
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Research.frx":74BA
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
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
      Height          =   1770
      Left            =   6180
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   2715
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Research.frx":77C4
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   68
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   315
         TabIndex        =   71
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcPlusDropDown 
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
      Left            =   4980
      Picture         =   "Research.frx":A5DE
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcPlusDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   180
      Left            =   3885
      MaxLength       =   20
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4545
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   945
      Picture         =   "Research.frx":A6D8
      ScaleHeight     =   1335
      ScaleWidth      =   5700
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   5730
   End
   Begin VB.PictureBox pbcPArrow 
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
      Height          =   180
      Left            =   60
      Picture         =   "Research.frx":2441E
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5115
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ListBox lbcPlusDemos 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5445
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pbcPlusSTab 
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
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   37
      Top             =   4920
      Width           =   60
   End
   Begin VB.PictureBox pbcPlusTab 
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
      Height          =   105
      Left            =   -15
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   45
      Top             =   5910
      Width           =   90
   End
   Begin VB.CommandButton cmcSocEco 
      Appearance      =   0  'Flat
      Caption         =   "&Qualitative"
      Height          =   285
      Left            =   6240
      TabIndex        =   52
      Top             =   5715
      Width           =   1050
   End
   Begin VB.CheckBox ckcSocEco 
      Caption         =   "Include Qualitative Data"
      Height          =   210
      Left            =   240
      TabIndex        =   73
      Top             =   1395
      Width           =   2535
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4035
      TabIndex        =   50
      Top             =   5715
      Width           =   1050
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7695
      TabIndex        =   88
      Top             =   45
      Width           =   1410
   End
   Begin VB.PictureBox pbcNewTab 
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
      Left            =   9165
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   285
      Width           =   60
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1065
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lbcDays 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4455
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmcSpecDropDown 
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
      Left            =   2955
      Picture         =   "Research.frx":24728
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcSpecDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox pbcSpecSTab 
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
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   780
      Width           =   60
   End
   Begin VB.PictureBox pbcSpecTab 
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
      Height          =   105
      Left            =   15
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   1560
      Width           =   90
   End
   Begin VB.PictureBox pbcArrow 
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
      Height          =   180
      Left            =   45
      Picture         =   "Research.frx":24822
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2445
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ListBox lbcDaypart 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4440
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3165
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.ListBox lbcSocEco 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4245
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3285
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8850
      Top             =   5625
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8460
      Top             =   5640
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   5880
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   885
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
      Left            =   6855
      Picture         =   "Research.frx":24B2C
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3675
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3315
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox pbcClickFocus 
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
      Height          =   60
      Left            =   3720
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   105
      Left            =   -45
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   32
      Top             =   4845
      Width           =   90
   End
   Begin VB.PictureBox pbcSTab 
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
      ScaleWidth      =   60
      TabIndex        =   19
      Top             =   1695
      Width           =   60
   End
   Begin VB.VScrollBar vbcDemo 
      Height          =   3090
      LargeChange     =   6
      Left            =   8865
      Min             =   1
      TabIndex        =   36
      Top             =   1665
      Value           =   1
      Width           =   240
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      ScaleHeight     =   195
      ScaleWidth      =   930
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   930
   End
   Begin VB.ListBox lbcSocEcoCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3180
      Sorted          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lbcDPCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   915
      Sorted          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   735
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3645
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3900
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4125
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   255
      Picture         =   "Research.frx":24C26
      ScaleHeight     =   765
      ScaleWidth      =   8610
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   555
      Width           =   8610
      Begin VB.PictureBox pbcEstByLorU 
         Appearance      =   0  'Flat
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
         Height          =   180
         Left            =   2580
         ScaleHeight     =   180
         ScaleWidth      =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox plcSpec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   195
      ScaleHeight     =   825
      ScaleWidth      =   8655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   8715
   End
   Begin VB.PictureBox plcDataType 
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
      Height          =   195
      Left            =   4380
      ScaleHeight     =   195
      ScaleWidth      =   4875
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1395
      Width           =   4875
      Begin VB.OptionButton rbcDataType 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3795
         TabIndex        =   18
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton rbcDataType 
         Caption         =   "Time"
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
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   750
      End
      Begin VB.OptionButton rbcDataType 
         Caption         =   "Extra Daypart"
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
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
      Begin VB.OptionButton rbcDataType 
         Caption         =   "Sold Daypart"
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
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.PictureBox pbcDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3105
      Index           =   2
      Left            =   2085
      Picture         =   "Research.frx":3AA98
      ScaleHeight     =   3105
      ScaleWidth      =   8625
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1665
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   0
         TabIndex        =   72
         Top             =   1320
         Visible         =   0   'False
         Width           =   8610
      End
   End
   Begin VB.PictureBox pbcDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3105
      Index           =   1
      Left            =   885
      Picture         =   "Research.frx":9239A
      ScaleHeight     =   3105
      ScaleWidth      =   8625
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1650
      Width           =   8625
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   60
         Top             =   1320
         Visible         =   0   'False
         Width           =   8610
      End
   End
   Begin VB.PictureBox pbcDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3105
      Index           =   0
      Left            =   270
      Picture         =   "Research.frx":E9C9C
      ScaleHeight     =   3105
      ScaleWidth      =   8625
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   8610
      End
   End
   Begin VB.PictureBox plcDemo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3240
      Left            =   210
      ScaleHeight     =   3180
      ScaleWidth      =   8880
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1635
      Width           =   8940
   End
   Begin VB.CommandButton cmcSetDefault 
      Appearance      =   0  'Flat
      Caption         =   "Set Default"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   53
      Top             =   5715
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo"
      Height          =   285
      Left            =   5130
      TabIndex        =   51
      Top             =   5715
      Width           =   1050
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   6240
      TabIndex        =   49
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5145
      TabIndex        =   48
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   4050
      TabIndex        =   47
      Top             =   5310
      Width           =   1050
   End
   Begin VB.PictureBox pbcDPorEst 
      Appearance      =   0  'Flat
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
      Height          =   180
      Left            =   225
      ScaleHeight     =   180
      ScaleWidth      =   3975
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox pbcUSA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   600
      Picture         =   "Research.frx":14159E
      ScaleHeight     =   1005
      ScaleWidth      =   2970
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5460
      Visible         =   0   'False
      Width           =   2970
      Begin VB.Label lacUSA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   75
         TabIndex        =   81
         Top             =   750
         Visible         =   0   'False
         Width           =   2970
      End
   End
   Begin VB.PictureBox pbcEst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   330
      Picture         =   "Research.frx":14B540
      ScaleHeight     =   1005
      ScaleWidth      =   2970
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5325
      Visible         =   0   'False
      Width           =   2970
      Begin VB.Label lacEst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   78
         Top             =   735
         Visible         =   0   'False
         Width           =   2970
      End
   End
   Begin VB.VScrollBar vbcPlus 
      Height          =   990
      LargeChange     =   3
      Left            =   3210
      Max             =   1
      Min             =   1
      TabIndex        =   46
      Top             =   5055
      Value           =   1
      Width           =   240
   End
   Begin VB.PictureBox pbcPlus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   240
      Picture         =   "Research.frx":1554E2
      ScaleHeight     =   1005
      ScaleWidth      =   2970
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2970
      Begin VB.Label lacPlus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   74
         Top             =   735
         Visible         =   0   'False
         Width           =   2970
      End
   End
   Begin VB.PictureBox plcPlus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   210
      ScaleHeight     =   1140
      ScaleWidth      =   3225
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5025
      Width           =   3285
   End
   Begin VB.Label lacEnd 
      Caption         =   "End"
      Height          =   210
      Left            =   3165
      TabIndex        =   85
      Top             =   30
      Width           =   450
   End
   Begin VB.Label lacStart 
      Caption         =   "Dates: Start"
      Height          =   210
      Left            =   1230
      TabIndex        =   84
      Top             =   30
      Width           =   1110
   End
   Begin VB.Label lacPlusTitle 
      Caption         =   "Pre-defined Dayparts"
      Height          =   195
      Left            =   210
      TabIndex        =   77
      Top             =   4845
      Width           =   2805
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   960
      Picture         =   "Research.frx":15F484
      Top             =   105
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3630
      Top             =   5145
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8730
      Picture         =   "Research.frx":15F78E
      Top             =   5295
      Width           =   480
   End
End
Attribute VB_Name = "Research"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Research.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Research.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Research input screen code
Option Explicit
Option Compare Text
Dim tmSCtrls(0 To 22)  As FIELDAREA     'Spec
Dim imLBSCtrls As Integer
Dim tmDCtrls(0 To 25)  As FIELDAREA    'Sold Daypart
Dim imLBDCtrls As Integer
Dim tmXCtrls(0 To 27)  As FIELDAREA    'Extra Daypart
Dim imLBXCtrls As Integer
Dim tmTCtrls(0 To 24)  As FIELDAREA    'Time
Dim imLBTCtrls As Integer
Dim tmVCtrls(0 To 22)  As FIELDAREA     'Vehicle
Dim imLBVCtrls As Integer
Dim tmPCtrls(0 To 3) As FIELDAREA   'Plus data
Dim imLBPCtrls As Integer
Dim imSBoxNo As Integer   'Current event name Box
Dim imLBDrf As Integer
Dim imLBDpf As Integer
Dim imLBDef As Integer
Dim imLBMnf As Integer
Dim imLBSaveShow As Integer
Dim imBoxNo As Integer
Dim imRowNo As Integer
Dim lmRowNo As Long
Dim imRowOffset As Integer '1=Vehicle on each row; 2=vehicle on each other row (skip row).
Dim smDataForm As String 'Blank or 6 = 16 demos; 8 = 18 demos
Dim smSSave(0 To 25) As String '1=NAMEINDEX, 2=DATEINDEX, 3=POPSRCEINDEX, 4=QUALPOPSRCEINDEX, 5=POPINDEX (5 - 23 = POP), 24=QUALSRCDESCINDEX, 25=POPSRCDESCINDEX
Dim tmSaveShow() As SAVESHOW
Dim imTestAddStdDemo As Integer
Dim smStdDemo() As String
Dim smCustomDemo(0 To 17) As String
Dim tmCustInfo() As CUSTINFO
Dim smUniqueGroupDataTypes() As String
Dim tmBNCode() As SORTCODE
Dim smBNCodeTag As String
Dim imDnfChg As Integer  'True=book name value changed; False=No changes
Dim imDrfChg As Integer
Dim imDpfChg As Integer
Dim imDefChg As Integer
Dim imPopChg As Integer
Dim smTotalPop As String
Dim tmDnfSrchKey As INTKEY0    'Dnf key record image
Dim hmDnf As Integer    'Demo Book Name file handle
Dim tmDnf As DNF
Dim imDnfRecLen As Integer        'DNF record length
Dim imPlusBoxNo As Integer
Dim lmPlusRowNo As Long      'lmRowNo last used to obtain DPF data
Dim lmEstRowNo As Long
Dim imEstBoxNo As Integer
Dim lmSDrfPopRecPos As Long 'Standard Pop
Dim lmCDrfPopRecPos As Long 'Custom Pop
Dim tmDrfSrchKey As DRFKEY0    'Drf key record image
Dim tmDrfSrchKey2 As LONGKEY0
Dim hmDrf As Integer    'Demo Research Data file handle
Dim imDrfRecLen As Integer        'DRF record length
Dim hmDpf As Integer 'Demo plus data file handle
Dim tmDpf As DPF        'DPF record image
Dim tmDpfSrchKey As LONGKEY0    'DPF key record image
Dim tmDpfSrchKey1 As DPFKEY1    'DPF key record image
Dim tmDpfSrchKey2 As DPFKEY2    'DPF key record image
Dim imDpfRecLen As Integer        'DPF record length
Dim tmDrfMap() As DRFMAP
Dim hmDef As Integer 'Demo plus data file handle
Dim tmDef As DEF        'DEF record image
Dim tmDefSrchKey As LONGKEY0    'DEF key record image
Dim tmDefSrchKey1 As DEFKEY1    'DEF key record image
Dim imDefRecLen As Integer        'DPF record length
Dim tmMnf As MNF        'Mnf record image
Dim tmMnfSrchKey As INTKEY0    'Mnf key record image
Dim hmMnf As Integer    'Multi-Name file handle
Dim imMnfRecLen As Integer        'MNF record length
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0     'VEF key 0 image
Dim imVefRecLen As Integer      'VEF record length
Dim hmVef As Integer            'Vehicle file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imSettingValue As Integer
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imIgnoreRightMove As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imCustomIndex As Integer
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim lmNowDate As Long
Dim imCurYear As Integer
Dim imCurMonth As Integer
Dim smMonDate As String
Dim smSyncDate As String
Dim smSyncTime As String
Dim imInNewTab As Integer
'Dim imListField(1 To 4) As Integer 'Set but not used
Dim tmPlusDemoCode() As SORTCODE
Dim smPlusDemoCodeTag As String
Dim imDPorEst As Integer '0=DP; 1=Est

Dim imEstByLOrU As Integer '0=Lister; 1=USA

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String
'11/15/11: Retain Model flag
Dim bmModelUsed As Boolean
'1/3/18
Dim bmResearchSaved As Boolean
'2/23/1:Filter information
Dim smFilterStartDate As String
Dim lmFilterStartDate As Long
Dim smFilterEndDate As String
Dim lmFilterEndDate As Long
Dim imFilterVefCode As Integer
'6/18/19: add source test
Dim smSource As String
Dim imP12PlusMnfCode As Integer
Dim tmSortDrfRec() As DRFREC
Dim tmSortDpfRec() As DPFREC
Dim tmSortDefRec() As DEFREC

Dim tmDP As FIELDAREA
Dim tmGroup As FIELDAREA
Dim bmIgnoreChg As Boolean
Dim lmVbcDemoLargeChg As Long

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Dim dnf_rst As ADODB.Recordset
Dim drf_rst As ADODB.Recordset

Const NAMEINDEX = 1
Const DATEINDEX = 2
Const POPSRCEINDEX = 3
Const QUALPOPSRCEINDEX = 4
Const POPINDEX = 5
Const QUALSRCDESCINDEX = 24
Const POPSRCDESCINDEX = 25

'by sold daypart
Const DVEHICLEINDEX = 1
Const DACT1CODEINDEX = 2
Const DACT1SETTINGINDEX = 3
Const DDAYPARTINDEX = 4
Const DGROUPINDEX = 5
Const DDEMOINDEX = 6
Const DAIRTIMEGRPNOINDEX = 24
Const DIMPRESSIONSINDEX = 25
'by extra daypart
Const XVEHICLEINDEX = 1
Const XACT1CODEINDEX = 2
Const XACT1SETTINGINDEX = 3
Const XTIMEINDEX = 4
Const XDAYSINDEX = 6
Const XGROUPINDEX = 7
Const XDEMOINDEX = 8
Const XGROUPNINDEX = 26
Const XGROUPIINDEX = 27
'by time
Const TVEHICLEINDEX = 1
Const TTIMEINDEX = 2
Const TDAYSINDEX = 4
Const TDEMOINDEX = 5
'by Vehicle
Const VVEHICLEINDEX = 1
Const VACT1CODEINDEX = 2
Const VACT1SETTINGINDEX = 3
Const VDAYSINDEX = 4
Const VDEMOINDEX = 5
'Plus
Const PDEMOINDEX = 1
Const PAUDINDEX = 2
Const PPOPINDEX = 3
'Est
Const EDATEINDEX = 1
Const EPOPINDEX = 2
Const EESTPCTINDEX = 3
'Hide cursor in inputs for ACT1Settings boxes
Private Declare Function ShowCaret Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal HWnd As Long) As Long

Dim imDataType As Integer '0=Daypart,1=ExtraDaypart,2=Time,3=Vehicle
Dim mAct1ColsWidth As Integer 'When in Podcast impression mode, the Act1code and setting column have to be removed - keeps track of width of these 2 columns

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Book name list box    *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mPopSrce()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String

    Exit Sub
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAddStdDemo                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Standard Demos              *
'*                                                     *
'*******************************************************
Private Function mAddStdDemo() As Integer
'
'   ilRet = mAddStdDemo ()
'   Where:
'       ilRet (O)- True = populated; False = error
'
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilAddMissingOnly As Integer

    If Not imTestAddStdDemo Then
        mAddStdDemo = True
        Exit Function
    End If
    imTestAddStdDemo = False
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    ilFilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilFilter(1) = INTEGERFILTER
    slFilter(1) = "0"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
    lbcDemo.Clear
    ilRet = gIMoveListBox(Research, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    smNameCodeTag = ""
    If lbcDemo.ListCount > 0 Then
        'Test if 20 exist
        For illoop = 1 To lbcDemo.ListCount - 1 Step 1
            If InStr(1, lbcDemo.List(illoop), "20", vbTextCompare) > 0 Then
                mAddStdDemo = True
                Exit Function
            End If
        Next illoop
        'Add in missing demos
        ilAddMissingOnly = True
    Else
        ilAddMissingOnly = False
    End If
    lbcDemo.Clear
    gDemoPop lbcDemo   'Get demo names
    gGetSyncDateTime slSyncDate, slSyncTime
    For illoop = 1 To lbcDemo.ListCount - 1 Step 1
        ilFound = False
        If ilAddMissingOnly Then
            For ilIndex = LBound(tmNameCode) To UBound(tmNameCode) - 1 Step 1
                If InStr(1, Trim$(tmNameCode(ilIndex).sKey), Trim$(lbcDemo.List(illoop)), vbTextCompare) > 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilIndex
        End If
        If Not ilFound Then
            tmMnf.iCode = 0
            tmMnf.sType = "D"
            tmMnf.sName = lbcDemo.List(illoop)
            tmMnf.sRPU = ""
            tmMnf.sUnitType = ""
            tmMnf.iMerge = 0
            tmMnf.iGroupNo = 0
            tmMnf.sCodeStn = ""
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            Do
                tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmMnf.iAutoCode = tmMnf.iCode
                gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    Next illoop
    mAddStdDemo = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mEstEnableBox                   *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEstEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBPCtrls Or ilBoxNo > UBound(tmPCtrls) Then
        Exit Sub
    End If
    If (lmEstRowNo < vbcPlus.Value) Or (lmEstRowNo > (vbcPlus.Value + vbcPlus.LargeChange)) Then
        pbcPArrow.Visible = False
        lacEst.Visible = False
        Exit Sub
    End If
    lacEst.Move 0, tmPCtrls(PDEMOINDEX).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15) - 30
    lacEst.Visible = True
    pbcPArrow.Move pbcPArrow.Left, plcPlus.Top + tmPCtrls(PDEMOINDEX).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15) + 45
    pbcPArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case EDATEINDEX
            edcPlusDropDown.Width = tmPCtrls(DATEINDEX).fBoxW - cmcPlusDropDown.Width
            edcPlusDropDown.MaxLength = 10
            gMoveTableCtrl pbcEst, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            cmcPlusDropDown.Move edcPlusDropDown.Left + edcPlusDropDown.Width, edcPlusDropDown.Top
            plcCalendar.Move edcPlusDropDown.Left, edcPlusDropDown.Top - plcCalendar.Height
            slStr = Trim$(tgDefRec(lmEstRowNo).sStartDate)
            If (slStr = "") And (smSSave(EDATEINDEX) <> "") Then
                slStr = gObtainNextMonday(smSSave(EDATEINDEX))
            End If
            If slStr = "" Then
                slStr = Format$(gNow(), "m/d/yy")   'Get year
                slStr = gObtainEndStd(slStr)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcPlusDropDown.Text = Trim$(tgDefRec(lmEstRowNo).sStartDate)
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True
            cmcPlusDropDown.Visible = True
            edcPlusDropDown.SetFocus
        Case EPOPINDEX
            edcPlusDropDown.Width = tmPCtrls(ilBoxNo).fBoxW
            edcPlusDropDown.MaxLength = 8
            gMoveTableCtrl pbcEst, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            edcPlusDropDown.Text = Trim$(tgDefRec(lmEstRowNo).sPop)
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
        Case EESTPCTINDEX
            If imEstByLOrU = 1 Then
                edcPlusDropDown.Width = tmPCtrls(EPOPINDEX).fBoxW + tmPCtrls(ilBoxNo).fBoxW + 15
            Else
                edcPlusDropDown.Width = tmPCtrls(ilBoxNo).fBoxW
            End If
            edcPlusDropDown.MaxLength = 8
            If imEstByLOrU = 1 Then
                gMoveTableCtrl pbcEst, edcPlusDropDown, tmPCtrls(EPOPINDEX).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            Else
                gMoveTableCtrl pbcEst, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            End If
            If Trim$(tgDefRec(lmEstRowNo).sEstPct) <> "" Then
                edcPlusDropDown.Text = gSubStr(Trim$(tgDefRec(lmEstRowNo).sEstPct), "100.00")
            Else
                edcPlusDropDown.Text = ""
            End If
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPlusEnableBox                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPlusEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBPCtrls Or ilBoxNo > UBound(tmPCtrls) Then
        Exit Sub
    End If
    If (lmPlusRowNo < vbcPlus.Value) Or (lmPlusRowNo > (vbcPlus.Value + vbcPlus.LargeChange)) Then
        'mPlusSetShow ilBoxNo
        pbcPArrow.Visible = False
        lacPlus.Visible = False
        Exit Sub
    End If
    lacPlus.Move 0, tmPCtrls(PDEMOINDEX).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15) - 30
    lacPlus.Visible = True
    pbcPArrow.Move pbcPArrow.Left, plcPlus.Top + tmPCtrls(PDEMOINDEX).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15) + 45
    pbcPArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case PDEMOINDEX
            lbcPlusDemos.Height = gListBoxHeight(lbcPlusDemos.ListCount, 6)
            edcPlusDropDown.Width = tmPCtrls(ilBoxNo).fBoxW
            edcPlusDropDown.MaxLength = 10
            gMoveTableCtrl pbcPlus, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            cmcPlusDropDown.Move edcPlusDropDown.Left + edcPlusDropDown.Width, edcPlusDropDown.Top
            lbcPlusDemos.Move edcPlusDropDown.Left, edcPlusDropDown.Top - lbcPlusDemos.Height
            slStr = Trim$(tgDpfRec(lmPlusRowNo).sKey)
            imChgMode = True
            gFindMatch slStr, 0, lbcPlusDemos
            If gLastFound(lbcPlusDemos) >= 0 Then
                lbcPlusDemos.ListIndex = gLastFound(lbcPlusDemos)
            Else
                lbcPlusDemos.ListIndex = -1
            End If
            If lbcPlusDemos.ListIndex < 0 Then
                edcPlusDropDown.Text = ""
            Else
                edcPlusDropDown.Text = lbcPlusDemos.List(lbcPlusDemos.ListIndex)
            End If
            imChgMode = False
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True  'Set visibility
            cmcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
        Case PAUDINDEX
            edcPlusDropDown.Width = tmPCtrls(ilBoxNo).fBoxW
            edcPlusDropDown.MaxLength = 8
            gMoveTableCtrl pbcPlus, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            edcPlusDropDown.Text = Trim$(tgDpfRec(lmPlusRowNo).sDemo)
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
        Case PPOPINDEX
            edcPlusDropDown.Width = tmPCtrls(ilBoxNo).fBoxW
            edcPlusDropDown.MaxLength = 8
            gMoveTableCtrl pbcPlus, edcPlusDropDown, tmPCtrls(ilBoxNo).fBoxX, tmPCtrls(ilBoxNo).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15)
            edcPlusDropDown.Text = Trim$(tgDpfRec(lmPlusRowNo).sPop)
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPlusSetFocus                   *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mEstSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPCtrls Or ilBoxNo > UBound(tmPCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case EDATEINDEX
            edcPlusDropDown.Visible = True
            cmcPlusDropDown.Visible = True
            edcPlusDropDown.SetFocus
        Case EPOPINDEX
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
        Case EESTPCTINDEX
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPlusSetFocus                   *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPlusSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPCtrls Or ilBoxNo > UBound(tmPCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PDEMOINDEX
            edcPlusDropDown.Visible = True
            cmcPlusDropDown.Visible = True
            edcPlusDropDown.SetFocus
        Case PAUDINDEX
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
        Case PPOPINDEX
            edcPlusDropDown.Visible = True  'Set visibility
            edcPlusDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEstSetShow                     *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mEstSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcPArrow.Visible = False
    lacEst.Visible = False
    If (ilBoxNo < imLBPCtrls) Or (ilBoxNo > UBound(tmPCtrls)) Then
        Exit Sub
    End If
    If (lmEstRowNo < imLBDef) Or (lmEstRowNo > UBound(tgDefRec)) Then
        Exit Sub
    End If
    Select Case ilBoxNo
        Case EDATEINDEX
            plcCalendar.Visible = False
            cmcPlusDropDown.Visible = False
            edcPlusDropDown.Visible = False  'Set visibility
            slStr = edcPlusDropDown.Text
            If gValidDate(slStr) Then
                If tgDefRec(lmEstRowNo).sStartDate <> slStr Then
                    'If imSelectedIndex > 0 Then
                        imDefChg = True
                    'End If
                End If
                tgDefRec(lmEstRowNo).sStartDate = slStr
            ElseIf slStr <> "" Then
                Beep
                edcPlusDropDown.Text = tgDefRec(lmEstRowNo).sStartDate
            End If
        Case EPOPINDEX
            edcPlusDropDown.Visible = False  'Set visibility
            slStr = edcPlusDropDown.Text
            If tgSpf.sSAudData = "H" Then
                gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
            End If
            If tgSpf.sSAudData = "N" Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            End If
            If tgSpf.sSAudData = "U" Then
                gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
            End If
            If gCompNumberStr(Trim$(tgDefRec(lmEstRowNo).sPop), slStr) <> 0 Then
                tgDefRec(lmEstRowNo).sPop = slStr
                slStr = gMulStr(tgDefRec(lmEstRowNo).sPop, "100.00")
                slStr = gDivStr(slStr, smTotalPop)
                slStr = gRoundStr(slStr, ".01", 2)
                tgDefRec(lmEstRowNo).sEstPct = slStr
                If lmEstRowNo < UBound(tgDefRec) Then
                    imDefChg = True
                End If
            End If
        Case EESTPCTINDEX
            edcPlusDropDown.Visible = False  'Set visibility
            slStr = edcPlusDropDown.Text
            slStr = gAddStr("100.00", slStr)
            If gCompNumberStr(Trim$(tgDefRec(lmEstRowNo).sEstPct), slStr) <> 0 Then
                tgDefRec(lmEstRowNo).sEstPct = slStr
                If lmEstRowNo < UBound(tgDefRec) Then
                    imDefChg = True
                End If
                'Compute Population
                slStr = gMulStr(smTotalPop, slStr)
                slStr = gDivStr(slStr, "100.00")
                slStr = gRoundStr(slStr, "1", 0)
                tgDefRec(lmEstRowNo).sPop = slStr
            End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPlusSetShow                    *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPlusSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcPArrow.Visible = False
    lacPlus.Visible = False
    '3/24/20: added here because many calls to this routine with ilBoxNo set to -1 prior to the call.  Most of these are in the GotFocus
    lbcPlusDemos.Visible = False
    cmcPlusDropDown.Visible = False
    edcPlusDropDown.Visible = False
    If (ilBoxNo < imLBPCtrls) Or (ilBoxNo > UBound(tmPCtrls)) Then
        Exit Sub
    End If
    If (lmPlusRowNo < imLBDpf) Or (lmPlusRowNo > UBound(tgDpfRec)) Then
        Exit Sub
    End If
    Select Case ilBoxNo
        Case PDEMOINDEX
            lbcPlusDemos.Visible = False
            cmcPlusDropDown.Visible = False
            edcPlusDropDown.Visible = False
            If lbcPlusDemos.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcPlusDemos.List(lbcPlusDemos.ListIndex)
            End If
            If Trim$(tgDpfRec(lmPlusRowNo).sKey) <> slStr Then
                tgDpfRec(lmPlusRowNo).sKey = slStr
                If lmPlusRowNo < UBound(tgDpfRec) Then
                    imDpfChg = True
                End If
            End If
        Case PAUDINDEX
            edcPlusDropDown.Visible = False  'Set visibility
            slStr = edcPlusDropDown.Text
            If tgSpf.sSAudData = "H" Then
                gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
            End If
            If tgSpf.sSAudData = "N" Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            End If
            If tgSpf.sSAudData = "U" Then
                gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
            End If
            If gCompNumberStr(Trim$(tgDpfRec(lmPlusRowNo).sDemo), slStr) <> 0 Then
                tgDpfRec(lmPlusRowNo).sDemo = slStr
                If lmPlusRowNo < UBound(tgDpfRec) Then
                    imDpfChg = True
                End If
            End If
        Case PPOPINDEX
            edcPlusDropDown.Visible = False  'Set visibility
            slStr = edcPlusDropDown.Text
            If tgSpf.sSAudData = "H" Then
                gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
            End If
            If tgSpf.sSAudData = "N" Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            End If
            If tgSpf.sSAudData = "U" Then
                gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
            End If
            If (gCompNumberStr(Trim$(tgDpfRec(lmPlusRowNo).sPop), slStr) <> 0) Or ((Trim$(tgDpfRec(lmPlusRowNo).sPop) = "") And (Trim$(slStr) <> "")) Then
                tgDpfRec(lmPlusRowNo).sPop = slStr
                If lmPlusRowNo < UBound(tgDpfRec) Then
                    imDpfChg = True
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub cbcDemo_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilLoopCount As Integer
    
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcDemo.ListIndex >= 0 Then
                    cbcDemo.Text = cbcDemo.List(cbcDemo.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ilRet = gOptionLookAhead(cbcDemo, imBSMode, slStr)
            mMoveCtrlToRec
            If ilRet = 0 Then
                imCustomIndex = cbcDemo.ListIndex
            Else
                imCustomIndex = -1
            End If
            mDetermineUniqueGroups
            mAddPopIfReq
            
            mMoveRecToCtrl
            mInitSShow
            mInitShow
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                ilIndex = 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                ilIndex = 2
            Else 'Vehicle
                ilIndex = 1
            End If
            pbcSpec.Cls
            pbcSpec_Paint
            pbcDemo(ilIndex).Cls
            pbcDemo_Paint ilIndex
        Loop While (imCustomIndex <> cbcDemo.ListIndex) And ((imCustomIndex <> 0) Or (cbcDemo.ListIndex >= 0))
        mSetCommands
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault    'Default
    Exit Sub
cbcDemoErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    imChgMode = False
    mTerminate
    Exit Sub
End Sub

Private Sub cbcDemo_Click()
    cbcDemo_Change
End Sub

Private Sub cbcDemo_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    imComboBoxIndex = imCustomIndex
    gCtrlGotFocus cbcDemo
    mSetCommands
End Sub

Private Sub cbcDemo_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcDemo_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilLoopCount As Integer
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcSelect.ListIndex >= 0 Then
                    cbcSelect.Text = cbcSelect.List(cbcSelect.ListIndex)
                End If
            End If
            '11/15/11: Clear Model flag
            bmModelUsed = False
            pbcEst.Cls
            pbcUSA.Cls
            pbcPlus.Cls
            ilLoopCount = ilLoopCount + 1
            ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
            If (ilRet = 0) And (cbcSelect.ListIndex > 1) Then
                slCode = cbcSelect.ItemData(cbcSelect.ListIndex)
                If Not mReadRec(Val(slCode), False) Then
                    GoTo cbcSelectErr
                End If
            ElseIf (ilRet = 0) And ((cbcSelect.ListIndex = 0) Or (cbcSelect.ListIndex = 1)) Then
                ilRet = 2
            Else
                If ilRet = 1 Then
                    cbcSelect.ListIndex = 0
                End If
                ilRet = 1   'Clear fields as no match name found
            End If
            pbcSpec.Cls
            If rbcDataType(0).Value Or rbcDataType(2).Value Or smSource = "I" Then 'Daypart or Time
                pbcDemo(0).Cls
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo(2).Cls
            Else 'Vehicle
                pbcDemo(1).Cls
            End If
            mDetermineUniqueGroups
            mAddPopIfReq
            If ilRet = 0 Then
                imSelectedIndex = cbcSelect.ListIndex
                mMoveRecToCtrl
            ElseIf ilRet = 2 Then
                imSelectedIndex = cbcSelect.ListIndex
                If imSelectedIndex = 0 Then
                    smDataForm = 8
                Else
                    smDataForm = 6
                End If
                mClearCtrlFields
                If InStr(1, slStr, "[New ", vbTextCompare) <= 0 Then
                    smSSave(NAMEINDEX) = slStr
                End If
            Else
                imSelectedIndex = 0
                smDataForm = 8
                mClearCtrlFields
                 If InStr(1, slStr, "[New ", vbTextCompare) <= 0 Then
                    smSSave(NAMEINDEX) = slStr
                End If
            End If
            mSetControls False
            If smSource = "I" Then 'Podcast Impression mode
                bmIgnoreChg = True
                rbcDataType(0).Value = True 'Daypart
                bmIgnoreChg = False
            End If
            mInitSShow
            mInitShow
            pbcSpec_Paint
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
        Loop While (imSelectedIndex <> cbcSelect.ListIndex) And ((imSelectedIndex <> 0) Or (cbcSelect.ListIndex >= 0))
        mSetControls False
        igDnfModel = 0
        mSetCommands
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault    'Default
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    imChgMode = False
    mTerminate
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    If cbcSelect.ListCount > 0 Then
        If imFirstFocus Then
            imFirstFocus = False
            cbcSelect.ListIndex = 0
            imSelectedIndex = 0
        End If
        If imSelectedIndex = -1 Then
            cbcSelect.ListIndex = 0
            imSelectedIndex = 0
        End If
        imComboBoxIndex = imSelectedIndex
        gCtrlGotFocus cbcSelect
    Else
        cmcGetBook.SetFocus
    End If
    mSetCommands
    Exit Sub
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcVehicle_Change()
    If cbcVehicle.ListIndex >= 0 Then
        imFilterVefCode = cbcVehicle.ItemData(cbcVehicle.ListIndex)
    Else
        imFilterVefCode = -1
    End If
    If imFirstActivate = False Then
        cmcGetBook.Visible = True
    End If
End Sub

Private Sub cbcVehicle_Click()
    cbcVehicle_Change
End Sub

Private Sub cbcVehicle_GotFocus()
    mSetCommands
End Sub

Private Sub ckcSocEco_Click()
    cbcSelect_Click
End Sub

Private Sub cmcAdjust_Click()
    Dim llRowNo As Long
    Dim blFd As Boolean
    Dim llVef As Long
    Dim ilVef As Integer
    Dim llUpper As Long
    Dim blModelUsed As Boolean
    ReDim tgResearchAdjustVehicle(0 To 0) As RESEARCHADJUSTVEHICLE
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    mMoveCtrlToRec
    If UBound(tgAllDpf) <= imLBDpf Then
        If imSelectedIndex > 1 Then
            mGetDpf tgDnf.iCode, False
        End If
    End If
    mBuildRearchAdjustVehicles False
    
    If UBound(tgResearchAdjustVehicle) > 0 Then
        igResearchModelMethod = 1 'Vehicle adjust
        RSModel.Show vbModal
        DoEvents
        If (igReturn = 1) Then
            mAdjustVehicleFields
            If smSource <> "I" Then 'Standard Airtime mode
                mMoveRecToCtrl
                lmSDrfPopRecPos = 0
                lmCDrfPopRecPos = 0
                mInitSShow
                mInitShow
                pbcSpec_Paint
            Else 'Podcast Impression mode
                mResetStatus
            End If
            pbcDemo(0).Cls
            pbcDemo(1).Cls
            pbcDemo(2).Cls
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
        Else
            mMoveRecToCtrl
            lmSDrfPopRecPos = 0
            lmCDrfPopRecPos = 0
            mInitSShow
            mInitShow
            pbcSpec_Paint
            pbcDemo(0).Cls
            pbcDemo(1).Cls
            pbcDemo(2).Cls
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
        End If
    End If
End Sub

Private Sub cmcAdjust_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf True
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcBaseDuplicate_Click()
    Dim llDpf As Long
    Dim llDrf As Long
    
    mMoveCtrlToRec  'Move values from tgSaveShow to tgDrfRec, then tgDrfRec to tgAllDrf
    llDpf = UBound(tgAllDpf)
    llDrf = UBound(tgAllDrf)
    igDuplDnfCode = tgDnf.iCode
    ResearchBaseDupl.Show vbModal

    mMoveRecToCtrl  'Move tgAllDrf to tgDrfRec, then tgDrfRec to tgSaveShow
    
    pbcDemo(0).Cls
    mInitShow
    pbcDemo_Paint 0

    If UBound(tgAllDrf) <> llDrf Then
        imDrfChg = True
    End If
    If UBound(tgAllDpf) <> llDpf Then
        imDpfChg = True
    End If
    mSetCommands
End Sub

Private Sub cmcBaseDuplicate_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imSBoxNo = DATEINDEX Then
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
        edcSpecDropDown.SetFocus
    ElseIf (imDPorEst = 1) And (imEstBoxNo = EDATEINDEX) Then
        edcPlusDropDown.SelStart = 0
        edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
        edcPlusDropDown.SetFocus
    End If
End Sub

Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imSBoxNo = DATEINDEX Then
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
        edcSpecDropDown.SetFocus
    ElseIf (imDPorEst = 1) And (imEstBoxNo = EDATEINDEX) Then
        edcPlusDropDown.SelStart = 0
        edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
        edcPlusDropDown.SetFocus
    End If
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDropDown_Click()
    If rbcDataType(0).Value Then 'Daypart
        Select Case imBoxNo
            Case DVEHICLEINDEX
                lbcVehicle.Visible = Not lbcVehicle.Visible
            Case DDAYPARTINDEX
                lbcDaypart.Visible = Not lbcDaypart.Visible
            Case DGROUPINDEX
                lbcSocEco.Visible = Not lbcSocEco.Visible
            Case DDEMOINDEX To DDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        Select Case imBoxNo
            Case XVEHICLEINDEX
                lbcVehicle.Visible = Not lbcVehicle.Visible
            Case XTIMEINDEX To XTIMEINDEX + 1
                plcTme.Visible = Not plcTme.Visible
            Case XDAYSINDEX
                lbcDays.Visible = Not lbcDays.Visible
            Case XGROUPINDEX
                lbcSocEco.Visible = Not lbcSocEco.Visible
            Case XDEMOINDEX To XDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(2).Value Then 'Time
        Select Case imBoxNo
            Case TVEHICLEINDEX
                lbcVehicle.Visible = Not lbcVehicle.Visible
            Case TTIMEINDEX To TTIMEINDEX + 1
                plcTme.Visible = Not plcTme.Visible
            Case TDAYSINDEX
                lbcDays.Visible = Not lbcDays.Visible
            Case TDEMOINDEX To TDEMOINDEX + 17
        End Select
    Else 'Vehicle
        Select Case imBoxNo
            Case VVEHICLEINDEX
                lbcVehicle.Visible = Not lbcVehicle.Visible
            Case VDAYSINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    lbcDays.Visible = Not lbcDays.Visible
                End If
            Case VDEMOINDEX To VDEMOINDEX + 17
        End Select
    End If
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDuplicate_Click()
    Dim ilNumberToCreate As Integer
    Dim illoop As Integer
    Dim llUpper As Long
    Dim slStr As String
    Dim ilDemo As Integer
    Dim ilAudSave As Integer
    Dim ilAudShow As Integer
    Dim ilAudStart As Integer
    Dim ilAudEnd As Integer
    
    If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
        Exit Sub
    End If
    sgGenMsg = "Number of Copies to Paste into New Research Rows"
    'Same daypart disallowed on same vehicle
    If rbcDataType(0).Value Then 'Daypart
        sgCMCTitle(0) = "Duplicate with Vehicle"
        sgCMCTitle(1) = "Duplicate W/O Vehicle"
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        sgCMCTitle(0) = "Duplicate with Vehicle"
        sgCMCTitle(1) = "Duplicate W/O Vehicle"
    ElseIf rbcDataType(2).Value Then 'Time
        sgCMCTitle(0) = "Duplicate with Vehicle"
        sgCMCTitle(1) = "Duplicate W/O Vehicle"
    Else 'Vehicle
        sgCMCTitle(0) = "Duplicate with Vehicle"
        sgCMCTitle(1) = "Duplicate W/O Vehicle"
    End If
    sgCMCTitle(2) = "Cancel"
    sgCMCTitle(3) = ""
    igDefCMC = 0
    igEditBox = 1
    sgEditValue = 1
    igEditBoxMaxCharacters = 2
    lgEditBoxMaxValue = 10
    edcStart.TabStop = False 'prevent psycho tabbing between date pickers
    edcEnd.TabStop = False
    GenMsg.Show vbModal
    If igAnsCMC = 2 Then
        Exit Sub
    End If
    If Val(sgEditValue) <= 0 Then
        Exit Sub
    End If
    ilNumberToCreate = Val(sgEditValue)
    imDrfChg = True
    For illoop = 1 To ilNumberToCreate Step 1
        'Vehicle
        llUpper = UBound(tmSaveShow)
        If igAnsCMC = 0 Then
            tmSaveShow(llUpper).sSave(1) = tmSaveShow(lmRowNo).sSave(1)
            slStr = tmSaveShow(lmRowNo).sSave(1)
        Else
            tmSaveShow(llUpper).sSave(1) = ""
            slStr = tmSaveShow(llUpper).sSave(1)
        End If
        
        If rbcDataType(0).Value Then 'Daypart
            'Vehicle name
            gSetShow pbcDemo(0), slStr, tmDCtrls(DVEHICLEINDEX)
            tmSaveShow(llUpper).sShow(DVEHICLEINDEX) = tmDCtrls(DVEHICLEINDEX).sShow
            'Act1 Code
            slStr = ""
            tmSaveShow(llUpper).sSave(DACT1CODEINDEX) = tmSaveShow(lmRowNo).sSave(DACT1CODEINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(DACT1CODEINDEX)
            gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1CODEINDEX)
            tmSaveShow(llUpper).sShow(DACT1CODEINDEX) = tmDCtrls(DACT1CODEINDEX).sShow
            'Act1 Setting
            slStr = ""
            tmSaveShow(llUpper).sSave(DACT1SETTINGINDEX) = tmSaveShow(lmRowNo).sSave(DACT1SETTINGINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(DACT1SETTINGINDEX)
            gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1SETTINGINDEX)
            tmSaveShow(llUpper).sShow(DACT1SETTINGINDEX) = tmDCtrls(DACT1SETTINGINDEX).sShow
            'Daypart
            slStr = ""
            tmSaveShow(llUpper).sSave(DDAYPARTINDEX) = tmSaveShow(lmRowNo).sSave(DDAYPARTINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(DDAYPARTINDEX)
            gSetShow pbcDemo(0), slStr, tmDCtrls(DDAYPARTINDEX)
            tmSaveShow(llUpper).sShow(DDAYPARTINDEX) = tmDCtrls(DDAYPARTINDEX).sShow
            'Group
            slStr = ""
            tmSaveShow(llUpper).sSave(DAIRTIMEGRPNOINDEX) = tmSaveShow(lmRowNo).sSave(DAIRTIMEGRPNOINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(DAIRTIMEGRPNOINDEX)
            gSetShow pbcDemo(0), slStr, tmDCtrls(DAIRTIMEGRPNOINDEX)
            tmSaveShow(llUpper).sShow(DAIRTIMEGRPNOINDEX) = tmDCtrls(DAIRTIMEGRPNOINDEX).sShow
            
            'Audience setup
            ilAudShow = DDEMOINDEX
            ilAudSave = DDEMOINDEX
            ilAudStart = DDEMOINDEX
            ilAudEnd = DDEMOINDEX + 17
            
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            'Vehicle name
            gSetShow pbcDemo(2), slStr, tmXCtrls(XVEHICLEINDEX)
            tmSaveShow(llUpper).sShow(XVEHICLEINDEX) = tmXCtrls(XVEHICLEINDEX).sShow
            'Act1 Code
            slStr = ""
            tmSaveShow(llUpper).sSave(XACT1CODEINDEX) = tmSaveShow(lmRowNo).sSave(XACT1CODEINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(XACT1CODEINDEX)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XACT1CODEINDEX)
            tmSaveShow(llUpper).sShow(XACT1CODEINDEX) = tmXCtrls(XACT1CODEINDEX).sShow
            'Act1 Setting
            slStr = ""
            tmSaveShow(llUpper).sSave(XACT1SETTINGINDEX) = tmSaveShow(lmRowNo).sSave(XACT1SETTINGINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(XACT1SETTINGINDEX)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XACT1SETTINGINDEX)
            tmSaveShow(llUpper).sShow(XACT1SETTINGINDEX) = tmXCtrls(XACT1SETTINGINDEX).sShow
            'Time1
            slStr = ""
            tmSaveShow(llUpper).sSave(XTIMEINDEX) = tmSaveShow(lmRowNo).sSave(XTIMEINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(XTIMEINDEX)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XTIMEINDEX)
            tmSaveShow(llUpper).sShow(XTIMEINDEX) = tmXCtrls(XTIMEINDEX).sShow
            'Time2
            slStr = ""
            tmSaveShow(llUpper).sSave(XTIMEINDEX + 1) = tmSaveShow(lmRowNo).sSave(XTIMEINDEX + 1)
            slStr = tmSaveShow(lmRowNo).sSave(XTIMEINDEX + 1)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XTIMEINDEX + 1)
            tmSaveShow(llUpper).sShow(XTIMEINDEX + 1) = tmXCtrls(XTIMEINDEX + 1).sShow
            'Days
            slStr = ""
            tmSaveShow(llUpper).sSave(XDAYSINDEX) = tmSaveShow(lmRowNo).sSave(XDAYSINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(XDAYSINDEX)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XDAYSINDEX)
            tmSaveShow(llUpper).sShow(XDAYSINDEX) = tmXCtrls(XDAYSINDEX).sShow
            'Group
            slStr = ""
            tmSaveShow(llUpper).sSave(XGROUPNINDEX) = tmSaveShow(lmRowNo).sSave(XGROUPNINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(XGROUPNINDEX)
            gSetShow pbcDemo(2), slStr, tmXCtrls(XGROUPNINDEX)
            tmSaveShow(llUpper).sShow(XGROUPNINDEX) = tmXCtrls(XGROUPNINDEX).sShow
            
            'Audience setup
            ilAudShow = XDEMOINDEX
            ilAudSave = XDEMOINDEX
            ilAudStart = XDEMOINDEX
            ilAudEnd = XDEMOINDEX + 17
            
        ElseIf rbcDataType(2).Value Then 'Time
            'Vehicle Name
            gSetShow pbcDemo(0), slStr, tmTCtrls(TVEHICLEINDEX)
            tmSaveShow(llUpper).sShow(TVEHICLEINDEX) = tmTCtrls(TVEHICLEINDEX).sShow
            'Time1
            slStr = ""
            tmSaveShow(llUpper).sSave(TTIMEINDEX) = tmSaveShow(lmRowNo).sSave(TTIMEINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(TTIMEINDEX)
            gSetShow pbcDemo(2), slStr, tmTCtrls(TTIMEINDEX)
            tmSaveShow(llUpper).sShow(TTIMEINDEX) = tmTCtrls(TTIMEINDEX).sShow
            'Time2
            slStr = ""
            tmSaveShow(llUpper).sSave(TTIMEINDEX + 1) = tmSaveShow(lmRowNo).sSave(TTIMEINDEX + 1)
            slStr = tmSaveShow(lmRowNo).sSave(TTIMEINDEX + 1)
            gSetShow pbcDemo(2), slStr, tmTCtrls(TTIMEINDEX + 1)
            tmSaveShow(llUpper).sShow(TTIMEINDEX + 1) = tmTCtrls(TTIMEINDEX + 1).sShow
            'Days
            slStr = ""
            tmSaveShow(llUpper).sSave(TDAYSINDEX) = tmSaveShow(lmRowNo).sSave(TDAYSINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(TDAYSINDEX)
            gSetShow pbcDemo(2), slStr, tmTCtrls(TDAYSINDEX)
            tmSaveShow(llUpper).sShow(TDAYSINDEX) = tmTCtrls(TDAYSINDEX).sShow
            
            'Audience setup
            ilAudShow = TDEMOINDEX
            ilAudSave = TDEMOINDEX
            ilAudStart = TDEMOINDEX
            ilAudEnd = TDEMOINDEX + 17
            
        Else 'Vehicle
            'Vehicle Name
            gSetShow pbcDemo(1), slStr, tmVCtrls(VVEHICLEINDEX)
            tmSaveShow(llUpper).sShow(VVEHICLEINDEX) = tmVCtrls(VVEHICLEINDEX).sShow
            'Act1 Code
            slStr = ""
            tmSaveShow(llUpper).sSave(VACT1CODEINDEX) = tmSaveShow(lmRowNo).sSave(VACT1CODEINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(VACT1CODEINDEX)
            gSetShow pbcDemo(2), slStr, tmVCtrls(VACT1CODEINDEX)
            tmSaveShow(llUpper).sShow(VACT1CODEINDEX) = tmVCtrls(VACT1CODEINDEX).sShow
            'Act1 Setting
            slStr = ""
            tmSaveShow(llUpper).sSave(VACT1CODEINDEX) = tmSaveShow(lmRowNo).sSave(VACT1SETTINGINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(VACT1SETTINGINDEX)
            gSetShow pbcDemo(2), slStr, tmVCtrls(VACT1SETTINGINDEX)
            tmSaveShow(llUpper).sShow(VACT1SETTINGINDEX) = tmVCtrls(VACT1SETTINGINDEX).sShow
            'Days
            slStr = ""
            tmSaveShow(llUpper).sSave(VDAYSINDEX) = tmSaveShow(lmRowNo).sSave(VDAYSINDEX)
            slStr = tmSaveShow(lmRowNo).sSave(VDAYSINDEX)
            gSetShow pbcDemo(2), slStr, tmVCtrls(VDAYSINDEX)
            tmSaveShow(llUpper).sShow(VDAYSINDEX) = tmVCtrls(VDAYSINDEX).sShow
            
            'Audience setup
            ilAudShow = VDEMOINDEX
            ilAudSave = VDEMOINDEX
            ilAudStart = VDEMOINDEX
            ilAudEnd = VDEMOINDEX + 17
        End If
        
        '-----------------------------------------
        For ilDemo = ilAudStart To ilAudEnd Step 1
            tmSaveShow(llUpper).sSave(ilAudSave + ilDemo - ilAudStart) = tmSaveShow(lmRowNo).sSave(ilAudSave + ilDemo - ilAudStart)
            slStr = tmSaveShow(lmRowNo).sSave(ilAudSave + ilDemo - ilAudStart)
            If rbcDataType(0).Value Then 'Daypart
                gSetShow pbcDemo(0), slStr, tmDCtrls(ilDemo)
                tmSaveShow(llUpper).sShow(ilAudShow + ilDemo - ilAudStart) = tmDCtrls(ilDemo).sShow
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                gSetShow pbcDemo(2), slStr, tmXCtrls(ilDemo)
                tmSaveShow(llUpper).sShow(ilAudShow + ilDemo - ilAudStart) = tmXCtrls(ilDemo).sShow
            ElseIf rbcDataType(2).Value Then 'Time
                gSetShow pbcDemo(0), slStr, tmTCtrls(ilDemo)
                tmSaveShow(llUpper).sShow(ilAudShow + ilDemo - ilAudStart) = tmTCtrls(ilDemo).sShow
            Else 'Vehicle
                gSetShow pbcDemo(1), slStr, tmVCtrls(ilDemo)
                tmSaveShow(llUpper).sShow(ilAudShow + ilDemo - ilAudStart) = tmVCtrls(ilDemo).sShow
            End If
        Next ilDemo
        ReDim Preserve tmSaveShow(0 To llUpper + 1) As SAVESHOW
        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
        mInitNewDrf True, UBound(tgDrfRec)
    Next illoop
    
    'Refresh data Grid
    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
        pbcDemo_Paint 0
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        pbcDemo_Paint 2
    Else 'Vehicle
        pbcDemo_Paint 1
    End If
    mSetCommands
    
    ' TTP 10765 - JJB 2023-0630  -- Scrollbar was not readjusting itself when new added records extended past the visible window
    If vbcDemo.LargeChange < UBound(tmSaveShow) Then vbcDemo.Max = UBound(tmSaveShow) - vbcDemo.LargeChange
    
    edcStart.TabStop = True
    edcEnd.TabStop = True
End Sub

Private Sub cmcDuplicate_GotFocus()
    If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
        cmcCancel.SetFocus
    End If
End Sub

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim illoop As Integer
    If imSelectedIndex > 1 Then
        ilRet = mEraseBook()
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If imTerminate Then
            Screen.MousePointer = vbDefault
            cmcCancel_Click
            Exit Sub
        End If
        
        edcStart.TabStop = False
        edcEnd.TabStop = False
        
        imBoxNo = -1 'Initialize current Box to N/A
        lmRowNo = -1
        lmPlusRowNo = -1
        imPlusBoxNo = -1
        imDnfChg = False
        imDrfChg = False
        imPopChg = False
        imDpfChg = False
        imDefChg = False
        lmPlusRowNo = -1
        ReDim tgDrfRec(0 To 1) As DRFREC
        ReDim tgDrfDel(0 To 1) As DRFREC
        ReDim tgAllDrf(0 To 1) As DRFREC
        ReDim tgDpfRec(0 To 1) As DPFREC
        ReDim tgDpfDel(0 To 1) As DPFREC
        ReDim tgAllDpf(0 To 1) As DPFREC
        ReDim tgDefRec(0 To 1) As DEFREC
        ReDim tgDefDel(0 To 1) As DEFREC
        For illoop = LBound(smSSave) To UBound(smSSave) Step 1
            smSSave(illoop) = ""
        Next illoop
        For illoop = LBound(tmSCtrls) To UBound(tmSCtrls) Step 1
            tmSCtrls(illoop).sShow = ""
            tmSCtrls(illoop).iChg = False
        Next illoop
        ReDim tmSaveShow(0 To 1) As SAVESHOW
        mInitNewDrf True, UBound(tgDrfRec)
        mInitNewDpf
        mInitNewDef
        
        pbcSpec.Cls
        pbcDemo(0).Cls
        pbcDemo(1).Cls
        pbcDemo(2).Cls
        pbcEst.Cls
        pbcUSA.Cls
        rbcDataType(3).Value = True
        smBNCodeTag = ""
        smTotalPop = ""
        imSelectedIndex = -1
        imCustomIndex = -1
        imFirstFocus = True
        cbcSelect.Clear
        mPopulate
        pbcSpec_Paint
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            pbcDemo_Paint 0
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            pbcDemo_Paint 2
        Else 'Vehicle
            pbcDemo_Paint 1
        End If
        mSetCommands
        
    ''' TTP 10560 BEGIN JJB
        If cbcSelect.Visible = True Then cbcSelect.SetFocus
        'cbcSelect.SetFocus
    ''' TTP 10560 END
        Screen.MousePointer = vbDefault
    End If
    
        
    edcStart.TabStop = True
    edcEnd.TabStop = True
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub

Private Sub cmcErase_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcGetBook_Click()
    If edcStart.Text <> "" Then
        If gIsDate(edcStart.Text) Then
            smFilterStartDate = edcStart.Text
            lmFilterStartDate = gDateValue(edcStart.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterStartDate = ""
        lmFilterStartDate = 0
    End If
    If edcEnd.Text <> "" Then
        If gIsDate(edcEnd.Text) Then
            smFilterEndDate = edcEnd.Text
            lmFilterEndDate = gDateValue(edcEnd.Text)
        Else
            Exit Sub
        End If
    Else
        smFilterEndDate = ""
        lmFilterEndDate = 0
    End If
    If cbcVehicle.ListIndex > 0 Then
        imFilterVefCode = cbcVehicle.ItemData(cbcVehicle.ListIndex)
    Else
        imFilterVefCode = -1
    End If
    Screen.MousePointer = vbHourglass  'Wait
    smBNCodeTag = ""
    mPopulate
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcGetBook_GotFocus()
    mSetCommands
End Sub

Private Sub cmcPlusDropDown_Click()
    If imDPorEst = 0 Then
        Select Case imPlusBoxNo
            Case PDEMOINDEX
                lbcPlusDemos.Visible = Not lbcPlusDemos.Visible
            Case PAUDINDEX
            Case PPOPINDEX
        End Select
    Else
        Select Case imEstBoxNo
            Case EDATEINDEX
                plcCalendar.Visible = Not plcCalendar.Visible
                edcPlusDropDown.SelStart = 0
                edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
                edcPlusDropDown.SetFocus
            Case EPOPINDEX
        End Select
    End If

End Sub

Private Sub cmcPlusDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSetDefault_GotFocus()
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcSetDefault_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcSave_Click()
    Dim slName As String
    Dim slDate As String
    Dim llDemoValue As Long
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    llDemoValue = vbcDemo.Value
    slName = smSSave(NAMEINDEX)
    slDate = smSSave(DATEINDEX)
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imSBoxNo > 0 Then
            mSEnableBox imSBoxNo
        ElseIf imBoxNo > 0 Then
            mEnableBox imBoxNo
        Else
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    ReDim tgDrfDel(0 To 1) As DRFREC
    smBNCodeTag = ""
    mPopulate
    slName = Trim$(slName) & ": " & slDate
    gFindMatch slName, 0, cbcSelect
    If gLastFound(cbcSelect) >= 0 Then
        If cbcSelect.ListIndex = gLastFound(cbcSelect) Then
            cbcSelect_Change
        Else
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        End If
        imDnfChg = False
        imDrfChg = False
        imPopChg = False
        imDpfChg = False
        imDefChg = False
        mSetCommands
        '2/20/19: Reset top
        If llDemoValue < vbcDemo.Max Then
            vbcDemo.Value = llDemoValue
        End If
        On Error Resume Next
        pbcClickFocus.SetFocus
        On Error GoTo 0
    Else
        If cbcSelect.ListIndex = 0 Then
            cbcSelect_Change
        Else
            cbcSelect.ListIndex = 0
        End If
        imDnfChg = False
        imDrfChg = False
        imPopChg = False
        imDpfChg = False
        imDefChg = False
        cbcSelect.SetFocus
    End If
End Sub

Private Sub cmcSave_GotFocus()
    lmPlusRowNo = -1
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcSetDefault_Click()
    igVehSelType = 1
    igVehSelCode = tgDnf.iCode
    VehSel.Show vbModal
End Sub

Private Sub cmcSocEco_Click()
    Dim slStr As String
    Dim ilRet As Integer

    sgMnfCallType = "F"
    igMNmCallSource = CALLNONE
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource))
    End If
    sgCommandStr = slStr
    On Error Resume Next
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilRet = mObtainSocEco()
End Sub

Private Sub cmcSocEco_GotFocus()
    mSetCommands
End Sub

Private Sub cmcSpecDropDown_Click()
    Select Case imSBoxNo
        Case NAMEINDEX
        Case DATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.SetFocus
        Case POPSRCEINDEX
            lbcPopSrce.Visible = Not lbcPopSrce.Visible
        Case QUALPOPSRCEINDEX
            lbcPopSrce.Visible = Not lbcPopSrce.Visible
        Case POPINDEX To POPINDEX + 17
    End Select
End Sub

Private Sub cmcUndo_Click()
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    '11/15/11: Clear model flag
    bmModelUsed = False
    
    pbcSpec.Cls
    pbcEst.Cls
    pbcUSA.Cls
    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
        pbcDemo(0).Cls
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        pbcDemo(2).Cls
    Else 'Vehicle
        pbcDemo(1).Cls
    End If
    If imSelectedIndex > 1 Then
        slCode = cbcSelect.ItemData(imSelectedIndex)
        ilRet = mReadRec(Val(slCode), False)
        mMoveRecToCtrl
    Else
        mClearCtrlFields
    End If
    mInitSShow
    mInitShow
    pbcSpec_Paint
    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
        pbcDemo_Paint 0
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        pbcDemo_Paint 2
    Else 'Vehicle
        pbcDemo_Paint 1
    End If
    imDnfChg = False
    imDrfChg = False
    imPopChg = False
    imDpfChg = False
    imDefChg = False
    mSetCommands
End Sub

Private Sub cmcUndo_GotFocus()
    lmPlusRowNo = -1
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    mSetCommands
End Sub

Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub CSI_ComboBoxMS1_GotFocus()
    pbcSpec.Enabled = False
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    If cbcSelect.ListCount > 0 Then
        If imFirstFocus Then
            imFirstFocus = False
            cbcSelect.ListIndex = 0
            imSelectedIndex = 0
        End If
        If imSelectedIndex = -1 Then
            cbcSelect.ListIndex = 0
            imSelectedIndex = 0
        End If
        imComboBoxIndex = imSelectedIndex
        gCtrlGotFocus cbcSelect
    Else
        cmcGetBook.SetFocus
    End If
    mSetCommands
    Exit Sub
End Sub

Private Sub CSI_ComboBoxMS1_GotInputFocus()
    pbcSpec.Enabled = False
End Sub

Private Sub CSI_ComboBoxMS1_LostInputFocus()
    pbcSpec.Enabled = True
End Sub

Private Sub CSI_ComboBoxMS1_OnChange()
    Dim illoop As Integer
    For illoop = 0 To cbcSelect.ListCount - 1
        If cbcSelect.List(illoop) = CSI_ComboBoxMS1.Text Then
            cbcSelect.ListIndex = illoop
            Exit For
        End If
    Next illoop
End Sub

Private Sub edcACT1SettingC_Click()
    If edcACT1SettingC.Text = "No" Then
        edcACT1SettingC.Text = "Yes"
    Else
        edcACT1SettingC.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingC_GotFocus()
    HideCaret edcACT1SettingC.HWnd
    edcACT1SettingC.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingC_KeyPress(KeyAscii As Integer)
    If edcACT1SettingC.Text = "No" Then
        edcACT1SettingC.Text = "Yes"
    Else
        edcACT1SettingC.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingC.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingC.Text = "No"
End Sub

Private Sub edcACT1SettingC_LostFocus()
    edcACT1SettingC.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingF_Click()
    If edcACT1SettingF.Text = "No" Then
        edcACT1SettingF.Text = "Yes"
    Else
        edcACT1SettingF.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingF_GotFocus()
    HideCaret edcACT1SettingF.HWnd
    edcACT1SettingF.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingF_KeyPress(KeyAscii As Integer)
    If edcACT1SettingF.Text = "No" Then
        edcACT1SettingF.Text = "Yes"
    Else
        edcACT1SettingF.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingF.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingF.Text = "No"
End Sub

Private Sub edcACT1SettingF_LostFocus()
    edcACT1SettingF.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingS_Click()
    If edcACT1SettingS.Text = "No" Then
        edcACT1SettingS.Text = "Yes"
    Else
        edcACT1SettingS.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingS_GotFocus()
    HideCaret edcACT1SettingS.HWnd
    edcACT1SettingS.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingS_KeyPress(KeyAscii As Integer)
    If edcACT1SettingS.Text = "No" Then
        edcACT1SettingS.Text = "Yes"
    Else
        edcACT1SettingS.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingS.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingS.Text = "No"
End Sub

Private Sub edcACT1SettingS_LostFocus()
    edcACT1SettingS.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingT_Click()
    If edcACT1SettingT.Text = "No" Then
        edcACT1SettingT.Text = "Yes"
    Else
        edcACT1SettingT.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingT_GotFocus()
    HideCaret edcACT1SettingT.HWnd
    edcACT1SettingT.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingT_KeyPress(KeyAscii As Integer)
    If edcACT1SettingT.Text = "No" Then
        edcACT1SettingT.Text = "Yes"
    Else
        edcACT1SettingT.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingT.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingT.Text = "No"
End Sub

Private Sub edcACT1SettingT_LostFocus()
    edcACT1SettingT.BackColor = &HFFFFFF
End Sub

Private Sub edcDropDown_Change()
    If rbcDataType(0).Value Then 'Daypart
        Select Case imBoxNo
            Case DVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            Case DDAYPARTINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcDaypart, imBSMode, imComboBoxIndex
            Case DGROUPINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSocEco, imBSMode, imComboBoxIndex
                End If
            Case DDEMOINDEX To DDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        Select Case imBoxNo
            Case XVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            Case XTIMEINDEX To XTIMEINDEX + 1
            Case XDAYSINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
            Case XGROUPINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcSocEco, imBSMode, imComboBoxIndex
            Case XDEMOINDEX To XDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(2).Value Then 'Time
        Select Case imBoxNo
            Case TVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            Case TTIMEINDEX To TTIMEINDEX + 1
            Case TDAYSINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
            Case TDEMOINDEX To TDEMOINDEX + 17
        End Select
    Else 'Vehicle
        Select Case imBoxNo
            Case VVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            Case VDAYSINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
                End If
            Case VDEMOINDEX To VDEMOINDEX + 17
        End Select
    End If
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub

Private Sub edcDropDown_GotFocus()
    If rbcDataType(0).Value Then 'Daypart
        Select Case imBoxNo
            Case DVEHICLEINDEX
                imComboBoxIndex = lbcVehicle.ListIndex
            Case DDAYPARTINDEX
                imComboBoxIndex = lbcDaypart.ListIndex
            Case DGROUPINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    imComboBoxIndex = lbcSocEco.ListIndex
                End If
            Case DDEMOINDEX To DDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        Select Case imBoxNo
            Case XVEHICLEINDEX
                imComboBoxIndex = lbcVehicle.ListIndex
            Case XTIMEINDEX To XTIMEINDEX + 1
            Case XDAYSINDEX
                imComboBoxIndex = lbcDays.ListIndex
            Case XGROUPINDEX
                imComboBoxIndex = lbcSocEco.ListIndex
            Case XDEMOINDEX To XDEMOINDEX + 17
        End Select
    ElseIf rbcDataType(2).Value Then 'Time
        Select Case imBoxNo
            Case TVEHICLEINDEX
                imComboBoxIndex = lbcVehicle.ListIndex
            Case TTIMEINDEX To TTIMEINDEX + 1
            Case TDAYSINDEX
                imComboBoxIndex = lbcDays.ListIndex
            Case TDEMOINDEX To TDEMOINDEX + 17
        End Select
    Else 'Vehicle
        Select Case imBoxNo
            Case VVEHICLEINDEX
                imComboBoxIndex = lbcVehicle.ListIndex
            Case VDAYSINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    imComboBoxIndex = lbcDays.ListIndex
                End If
            Case VDEMOINDEX To VDEMOINDEX + 17
        End Select
    End If
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilPos As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If rbcDataType(0).Value Then 'Daypart
        Select Case imBoxNo
            Case DVEHICLEINDEX
            Case DACT1CODEINDEX
            Case DACT1SETTINGINDEX
            Case DDAYPARTINDEX
            Case DGROUPINDEX
                If smSource = "I" Then 'Podcast Impression mode
                    If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        ilPos = InStr(edcDropDown.SelText, ".")
                        If ilPos = 0 Then
                            ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                            If ilPos > 0 Then
                                If KeyAscii = KEYDECPOINT Then
                                    Beep
                                    KeyAscii = 0
                                    Exit Sub
                                End If
                            End If
                        End If
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                    slStr = edcDropDown.Text
                    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                    If gCompNumberStr(slStr, "99999999") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            Case DDEMOINDEX To DDEMOINDEX + 17
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcDropDown.Text
                slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        Select Case imBoxNo
            Case XVEHICLEINDEX
            Case XACT1CODEINDEX
            Case XACT1SETTINGINDEX
            Case XTIMEINDEX To XTIMEINDEX + 1
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    ilFound = False
                    For illoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                        If KeyAscii = igLegalTime(illoop) Then
                            ilFound = True
                            Exit For
                        End If
                    Next illoop
                    If Not ilFound Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            Case XDAYSINDEX
            Case XGROUPINDEX
            Case XDEMOINDEX To XDEMOINDEX + 17
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcDropDown.Text
                slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    ElseIf rbcDataType(2).Value Then 'Time
        Select Case imBoxNo
            Case TVEHICLEINDEX
            Case TTIMEINDEX To TTIMEINDEX + 1
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    ilFound = False
                    For illoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                        If KeyAscii = igLegalTime(illoop) Then
                            ilFound = True
                            Exit For
                        End If
                    Next illoop
                    If Not ilFound Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            Case TDAYSINDEX
            Case TDEMOINDEX To TDEMOINDEX + 17
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcDropDown.Text
                slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    Else 'Vehicle
        Select Case imBoxNo
            Case VVEHICLEINDEX
            Case VACT1CODEINDEX
            Case VACT1SETTINGINDEX
            Case VDAYSINDEX
                If smSource = "I" Then 'Podcast Impression mode
                    If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        ilPos = InStr(edcDropDown.SelText, ".")
                        If ilPos = 0 Then
                            ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                            If ilPos > 0 Then
                                If KeyAscii = KEYDECPOINT Then
                                    Beep
                                    KeyAscii = 0
                                    Exit Sub
                                End If
                            End If
                        End If
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                    slStr = edcDropDown.Text
                    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                    If gCompNumberStr(slStr, "99999999") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            Case VDEMOINDEX To VDEMOINDEX + 17
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcDropDown.Text
                slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If rbcDataType(0).Value Then 'Daypart
            Select Case imBoxNo
                Case DVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                Case DACT1CODEINDEX
                Case DACT1SETTINGINDEX
                Case DDAYPARTINDEX
                    gProcessArrowKey Shift, KeyCode, lbcDaypart, imLbcArrowSetting
                Case DGROUPINDEX
                    If smSource <> "I" Then 'Standard Airtime mode
                        gProcessArrowKey Shift, KeyCode, lbcSocEco, imLbcArrowSetting
                    End If
                Case DDEMOINDEX To DDEMOINDEX + 17
            End Select
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            Select Case imBoxNo
                Case XVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                Case XACT1CODEINDEX
                Case XACT1SETTINGINDEX
                Case XTIMEINDEX To XTIMEINDEX + 1
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                Case XDAYSINDEX
                    gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
                Case XGROUPINDEX
                    gProcessArrowKey Shift, KeyCode, lbcSocEco, imLbcArrowSetting
                Case XDEMOINDEX To XDEMOINDEX + 17
            End Select
        ElseIf rbcDataType(2).Value Then 'Time
            Select Case imBoxNo
                Case TVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                Case TTIMEINDEX To TTIMEINDEX + 1
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                Case TDAYSINDEX
                    gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
                Case TDEMOINDEX To TDEMOINDEX + 17
            End Select
        Else 'Vehicle
            Select Case imBoxNo
                Case VVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                Case VACT1CODEINDEX
                Case VACT1SETTINGINDEX
                Case VDAYSINDEX
                    If smSource <> "I" Then 'Standard Airtime mode
                        gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
                    End If
                Case VDEMOINDEX To VDEMOINDEX + 17
            End Select
        End If
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imDoubleClickName = False
End Sub

Private Sub edcEnd_Change()
    cmcGetBook.Visible = True
End Sub

Private Sub edcEnd_GotFocus()
    mSetCommands
End Sub

Private Sub edcLinkDestDoneMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcLinkDestHelpMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcLinkSrceDoneMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcPlusDropDown_Change()
    Dim slStr As String

    If imDPorEst = 0 Then
        Select Case imPlusBoxNo
            Case PDEMOINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcPlusDropDown, lbcPlusDemos, imBSMode, imComboBoxIndex
            Case PAUDINDEX
            Case PPOPINDEX
        End Select
        imLbcArrowSetting = False
    Else
        Select Case imEstBoxNo
            Case EDATEINDEX
                slStr = edcPlusDropDown.Text
                If Not gValidDate(slStr) Then
                    lacDate.Visible = False
                    Exit Sub
                End If
                lacDate.Visible = True
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint   'mBoxCalDate called within paint
            Case EPOPINDEX
            Case EESTPCTINDEX
        End Select
    End If
End Sub

Private Sub edcPlusDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub

Private Sub edcPlusDropDown_GotFocus()
    If imDPorEst = 0 Then
        Select Case imPlusBoxNo
            Case PDEMOINDEX
                imComboBoxIndex = lbcPlusDemos.ListIndex
            Case PAUDINDEX
            Case PPOPINDEX
        End Select
    Else
        Select Case imEstBoxNo
            Case EDATEINDEX
            Case EPOPINDEX
            Case EESTPCTINDEX
        End Select
    End If
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcPlusDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcPlusDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    Dim ilPos As Integer

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcPlusDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If imDPorEst = 0 Then
        Select Case imPlusBoxNo
            Case PDEMOINDEX
            Case PAUDINDEX, PPOPINDEX
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcPlusDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcPlusDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcPlusDropDown.Text
                slStr = Left$(slStr, edcPlusDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPlusDropDown.SelStart - edcPlusDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    Else
        Select Case imEstBoxNo
            Case EDATEINDEX
            Case EPOPINDEX
                If (tgSpf.sSAudData <> "H") And (tgSpf.sSAudData <> "N") And (tgSpf.sSAudData <> "U") Then
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Else
                    ilPos = InStr(edcPlusDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcPlusDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                slStr = edcPlusDropDown.Text
                slStr = Left$(slStr, edcPlusDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPlusDropDown.SelStart - edcPlusDropDown.SelLength)
                If gCompNumberStr(slStr, "99999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case EESTPCTINDEX
                ilPos = InStr(edcPlusDropDown.SelText, ".")
                If ilPos = 0 Then
                    ilPos = InStr(edcPlusDropDown.Text, ".")    'Disallow multi-decimal points
                    If ilPos > 0 Then
                        If KeyAscii = KEYDECPOINT Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If (KeyAscii <> KEYBACKSPACE) And (KeyAscii <> KEYNEG) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
                slStr = edcPlusDropDown.Text
                slStr = Left$(slStr, edcPlusDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPlusDropDown.SelStart - edcPlusDropDown.SelLength)
                If gCompNumberStr(slStr, "999999") > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub edcPlusDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If imDPorEst = 0 Then
            Select Case imPlusBoxNo
                Case PDEMOINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPlusDemos, imLbcArrowSetting
                Case PAUDINDEX
                Case PPOPINDEX
            End Select
        Else
             Select Case imEstBoxNo
                Case EDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plcCalendar.Visible = Not plcCalendar.Visible
                    Else
                        slDate = edcPlusDropDown.Text
                        If gValidDate(slDate) Then
                            If KeyCode = KEYUP Then 'Up arrow
                                slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                            Else
                                slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                            edcPlusDropDown.Text = slDate
                        End If
                    End If
                Case EPOPINDEX
                Case EESTPCTINDEX
            End Select
            edcPlusDropDown.SelStart = 0
            edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
        End If
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If imDPorEst = 0 Then
            Select Case imPlusBoxNo
                Case PDEMOINDEX
                Case PAUDINDEX
                Case PPOPINDEX
            End Select
        Else
            Select Case imEstBoxNo
                Case EDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                    Else
                        slDate = edcDropDown.Text
                        If gValidDate(slDate) Then
                            If KeyCode = KEYLEFT Then 'Up arrow
                                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                            Else
                                slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                            edcPlusDropDown.Text = slDate
                        End If
                    End If
                    edcPlusDropDown.SelStart = 0
                    edcPlusDropDown.SelLength = Len(edcDropDown.Text)
                Case EPOPINDEX
                Case EESTPCTINDEX
            End Select
        End If
    End If
End Sub

Private Sub edcPlusDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imDoubleClickName = False
End Sub

Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Select Case imSBoxNo
        Case NAMEINDEX
        Case DATEINDEX
            slStr = edcSpecDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case POPSRCEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSpecDropDown, lbcPopSrce, imBSMode, imComboBoxIndex
        Case QUALPOPSRCEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSpecDropDown, lbcPopSrce, imBSMode, imComboBoxIndex
        Case POPINDEX To POPINDEX + 17
    End Select
End Sub

Private Sub edcSpecDropDown_GotFocus()
    lmPlusRowNo = -1
    mShowDpf False
    lmPlusRowNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    Select Case imBoxNo
        Case NAMEINDEX
        Case DATEINDEX
        Case POPSRCEINDEX
            imComboBoxIndex = lbcPopSrce.ListIndex
        Case QUALPOPSRCEINDEX
            imComboBoxIndex = lbcPopSrce.ListIndex
        Case POPINDEX To POPINDEX + 17
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case NAMEINDEX
        Case DATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case POPSRCEINDEX
        Case QUALPOPSRCEINDEX
        Case POPINDEX To POPINDEX + 17
    End Select
End Sub

Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imSBoxNo
            Case NAMEINDEX
            Case DATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpecDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropDown.Text = slDate
                    End If
                End If
            Case POPSRCEINDEX
                gProcessArrowKey Shift, KeyCode, lbcPopSrce, imLbcArrowSetting
            Case QUALPOPSRCEINDEX
                gProcessArrowKey Shift, KeyCode, lbcPopSrce, imLbcArrowSetting
            Case POPINDEX To POPINDEX + 17
        End Select
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case NAMEINDEX
            Case DATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropDown.Text = slDate
                    End If
                End If
                edcSpecDropDown.SelStart = 0
                edcSpecDropDown.SelLength = Len(edcDropDown.Text)
            Case POPINDEX To POPINDEX + 17
        End Select
    End If
End Sub

Private Sub edcSpecDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imDoubleClickName = False
End Sub

Private Sub edcStart_Change()
    cmcGetBook.Visible = True
End Sub

Private Sub edcStart_GotFocus()
    mSetCommands
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(RESEARCHLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
        pbcSpec.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcDemo(0).Enabled = False
        pbcDemo(1).Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        pbcPlus.Enabled = False
        pbcPlusSTab.Enabled = False
        pbcPlusTab.Enabled = False
    Else
        If imSelectedIndex < 0 Then
            pbcSpec.Enabled = False
            pbcSpecSTab.Enabled = False
            pbcSpecTab.Enabled = False
            pbcDemo(0).Enabled = False
            pbcDemo(1).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            pbcPlus.Enabled = False
            pbcPlusSTab.Enabled = False
            pbcPlusTab.Enabled = False
            cmcSave.Enabled = False
            cmcUndo.Enabled = False
        Else
            pbcSpec.Enabled = True
            pbcSpecSTab.Enabled = True
            pbcSpecTab.Enabled = True
            pbcDemo(0).Enabled = True
            pbcDemo(1).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            pbcPlus.Enabled = True
            pbcPlusSTab.Enabled = True
            pbcPlusTab.Enabled = True
        End If
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Research.Refresh
    
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imSBoxNo > 0 Then
            mSEnableBox imSBoxNo
        ElseIf imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        'Expand only vehicle column
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    dnf_rst.Close
    drf_rst.Close
    
    Erase tmSortDrfRec
    Erase tmSortDefRec
    Erase tmSortDpfRec
    Erase tmDrfMap
    Erase tgMnfSocEco
    Erase tgDrfRec
    Erase tgDrfDel
    Erase tgAllDrf
    Erase tgDpfRec
    Erase tgDpfDel
    Erase tgAllDpf
    Erase tgGDrfPop
    Erase tgCDrfPop
    Erase tgLinkDrfRec
    Erase tmNameCode
    Erase tmSaveShow
    Erase smStdDemo
    Erase tmPlusDemoCode
    Erase tmCustInfo
    Erase smUniqueGroupDataTypes
    smPlusDemoCodeTag = ""
    Erase tmBNCode
    btrExtClear hmDef   'Clear any previous extend operation
    ilRet = btrClose(hmDef)
    btrDestroy hmDef
    btrExtClear hmDpf   'Clear any previous extend operation
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    btrExtClear hmDrf   'Clear any previous extend operation
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    btrExtClear hmDnf   'Clear any previous extend operation
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    
    Set Research = Nothing
End Sub

Private Sub imcHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = False
End Sub

Private Sub imcTrash_Click()
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim llRowNo As Long
    Dim llUpperBound As Long
    Dim llDpf As Long
    Dim llDel As Long
    Dim ilFound As Integer
    Dim llMove As Long
    Dim llDelDrfCode As Long

    If ((lmRowNo < vbcDemo.Value) Or (lmRowNo > vbcDemo.Value + vbcDemo.LargeChange + 1)) And (imEstBoxNo >= imLBPCtrls) And (imEstBoxNo <= UBound(tmPCtrls)) And (lmEstRowNo > 0) Then
        If imDPorEst <> 1 Then
            Exit Sub
        End If
    
        edcStart.TabStop = False
        edcEnd.TabStop = False
        
        llRowNo = lmEstRowNo
        lacEst.Visible = False
        pbcPArrow.Visible = False
        mEstSetShow imEstBoxNo
        imEstBoxNo = -1
        lmEstRowNo = -1
        gCtrlGotFocus ActiveControl
        llUpperBound = UBound(tgDefRec)
        If tgDefRec(llRowNo).iStatus = 1 Then
            tgDefDel(UBound(tgDefDel)) = tgDefRec(llRowNo)
            ReDim Preserve tgDefDel(0 To UBound(tgDefDel) + 1) As DEFREC
        End If
        For llLoop = llRowNo To llUpperBound - 1 Step 1
            tgDefRec(llLoop) = tgDefRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tgDefRec(0 To UBound(tgDefRec) - 1) As DEFREC
        mSetDefScrollBar
        imDefChg = True
        mSetCommands
        If imEstByLOrU = 1 Then
            pbcUSA.Cls
            pbcUSA_Paint
        Else
            pbcEst.Cls
            pbcEst_Paint
        End If
        Exit Sub
    End If
    If (lmRowNo < vbcDemo.Value) Or (lmRowNo > vbcDemo.Value + vbcDemo.LargeChange + 1) Then
        If imSelectedIndex > 1 Then
            ilRet = mEraseBook()
            If Not ilRet Then
                Exit Sub
            End If
            edcStart.TabStop = False
            edcEnd.TabStop = False
            
            mPopulate
            pbcSpec.Cls
            pbcEst.Cls
            pbcUSA.Cls
            pbcDemo(0).Cls
            pbcDemo(1).Cls
            pbcDemo(2).Cls
            mClearCtrlFields
            pbcSpec_Paint
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
            Screen.MousePointer = vbDefault
        End If
        Exit Sub
    End If
    llRowNo = lmRowNo
    If (imPlusBoxNo >= imLBPCtrls) And (imPlusBoxNo <= UBound(tmPCtrls)) And (lmPlusRowNo > 0) Then
        If (imCustomIndex > 0) Or (lmRowNo = -1) Then
            Exit Sub
        End If
        If tgDrfRec(lmRowNo).iStatus = 0 Then
            Exit Sub
        End If
        edcStart.TabStop = False
        edcEnd.TabStop = False
        
        llRowNo = lmPlusRowNo
        mShowDpf False
        lacPlus.Visible = False
        pbcPArrow.Visible = False
        lacEst.Visible = False
        gCtrlGotFocus ActiveControl
        llUpperBound = UBound(tgDpfRec)
        If tgDpfRec(llRowNo).iStatus = 1 Then
            tgDpfDel(UBound(tgDpfDel)) = tgDpfRec(llRowNo)
            ReDim Preserve tgDpfDel(0 To UBound(tgDpfDel) + 1) As DPFREC
        End If
        For llLoop = llRowNo To ilUpperBound - 1 Step 1
            tgDpfRec(llLoop) = tgDpfRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) - 1) As DPFREC
        imSettingValue = True
        If UBound(tgDpfRec) <= vbcPlus.LargeChange + 1 Then
            vbcPlus.Max = imLBDpf   'LBound(tgDpfRec)
        Else
            vbcPlus.Max = UBound(tgDpfRec) - vbcPlus.LargeChange
        End If
        imDpfChg = True
        mSetCommands
        pbcPlus.Cls
        pbcPlus_Paint
    Else
        'Save values, the one below will cause the array to be cleared
        edcStart.TabStop = False
        edcEnd.TabStop = False
        
        mShowDpf False
        mSetShow imBoxNo, True
        imBoxNo = -1
        lmRowNo = -1
        mShowDpf False
        mEstSetShow imEstBoxNo
        imEstBoxNo = -1
        lmEstRowNo = -1
        pbcArrow.Visible = False
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            lacFrame(0).Visible = False
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            lacFrame(2).Visible = False
        Else 'Vehicle
            lacFrame(1).Visible = False
        End If
        'gCtrlGotFocus ActiveControl
        llUpperBound = UBound(tmSaveShow)
        llDelDrfCode = 0
        If tgDrfRec(llRowNo).iStatus = 1 Then
            llDelDrfCode = tgDrfRec(llRowNo).tDrf.lCode
            tgDrfDel(UBound(tgDrfDel)) = tgDrfRec(llRowNo)
            ReDim Preserve tgDrfDel(0 To UBound(tgDrfDel) + 1) As DRFREC
        End If
        'Remove record from tgRjf1Rec- Leave tgPjf2Rec
        For llLoop = llRowNo To llUpperBound - 1 Step 1
            tgDrfRec(llLoop) = tgDrfRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) - 1) As DRFREC
        For llLoop = llRowNo To llUpperBound - 1 Step 1
            For ilIndex = imLBSaveShow To UBound(tmSaveShow(imLBSaveShow).sSave) Step 1
                tmSaveShow(llLoop).sSave(ilIndex) = tmSaveShow(llLoop + 1).sSave(ilIndex)
            Next ilIndex
            For ilIndex = imLBSaveShow To UBound(tmSaveShow(imLBSaveShow).sShow) Step 1
                tmSaveShow(llLoop).sShow(ilIndex) = tmSaveShow(llLoop + 1).sShow(ilIndex)
            Next ilIndex
        Next llLoop
        llUpperBound = UBound(tmSaveShow)
        ReDim Preserve tmSaveShow(0 To llUpperBound - 1) As SAVESHOW
        imSettingValue = True
        'Move dpf to del
        llDpf = imLBDpf 'LBound(tgAllDpf)
        Do While llDpf < UBound(tgAllDpf)
            If tgAllDpf(llDpf).lDrfCode = llDelDrfCode Then
                ilFound = False
                For llDel = imLBDpf To UBound(tgDpfDel) - 1 Step 1
                    If tgDpfDel(llDel).lDpfCode = tgAllDpf(llDpf).lDpfCode Then
                        ilFound = True
                        Exit For
                    End If
                Next llDel
                If Not ilFound Then
                    imDpfChg = True
                    tgDpfDel(UBound(tgDpfDel)) = tgAllDpf(llDpf)
                    ReDim Preserve tgDpfDel(0 To UBound(tgDpfDel) + 1) As DPFREC
                    For llMove = llDpf To UBound(tgAllDpf) - 1 Step 1
                        tgAllDpf(llMove) = tgAllDpf(llMove + 1)
                    Next llMove
                    ReDim Preserve tgAllDpf(0 To UBound(tgAllDpf) - 1) As DPFREC
                Else
                    llDpf = llDpf + 1
                End If
            Else
                llDpf = llDpf + 1
            End If
        Loop
        If UBound(tmSaveShow) <= vbcDemo.LargeChange + 1 Then ' + 1 Then
            vbcDemo.Max = imLBSaveShow   'LBound(tmSaveShow)
        Else
            vbcDemo.Max = UBound(tmSaveShow) - vbcDemo.LargeChange '- 1
        End If
        imDrfChg = True
        mSetCommands
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            lacFrame(0).DragIcon = IconTraf!imcIconDrag.DragIcon
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            lacFrame(2).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else 'Vehicle
            lacFrame(1).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        pbcDemo(0).Cls
        pbcDemo(1).Cls
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            pbcDemo_Paint 0
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            pbcDemo_Paint 2
        Else 'Vehicle
            pbcDemo_Paint 1
        End If
    End If
    pbcClickFocus.SetFocus
    edcStart.TabStop = True
    edcEnd.TabStop = True
    
End Sub

Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    imcTrash_Click
End Sub

Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            lacFrame(0).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            lacFrame(2).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        Else 'Vehicle
            lacFrame(1).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            lacFrame(0).DragIcon = IconTraf!imcIconDrag.DragIcon
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            lacFrame(2).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else 'Vehicle
            lacFrame(1).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub

Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lacDate_Click()
    mSetCommands
End Sub

Private Sub lacEnd_Click()
    mSetCommands
End Sub

Private Sub lacPlusTitle_Click()
    mSetCommands
End Sub

Private Sub lacStart_Click()
    mSetCommands
End Sub

Private Sub lbcDaypart_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcDaypart, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcDaypart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcDays_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcDays, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPlusDemos_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcPlusDemos, edcPlusDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcPlusDemos_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPopSrce_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcPopSrce, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcPopSrce_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcSocEco_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcSocEco, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcSocEco_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehicle_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imSBoxNo = DATEINDEX Then
        slStr = edcSpecDropDown.Text
    ElseIf (imDPorEst = 1) And (imEstBoxNo = EDATEINDEX) Then
        slStr = edcPlusDropDown.Text
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim illoop As Integer
    imEstByLOrU = 0
    pbcEstByLorU.Cls
    pbcEstByLorU_Paint
    If smDataForm = "6" Then
        smStdDemo(0) = "M12-17"
        smStdDemo(1) = "M18-24"
        smStdDemo(2) = "M25-34"
        smStdDemo(3) = "M35-44"
        smStdDemo(4) = "M45-49"
        smStdDemo(5) = "M50-54"
        smStdDemo(6) = "M55-64"
        smStdDemo(7) = "M65+"
        smStdDemo(8) = ""
        smStdDemo(9) = "W12-17"
        smStdDemo(10) = "W18-24"
        smStdDemo(11) = "W25-34"
        smStdDemo(12) = "W35-44"
        smStdDemo(13) = "W45-49"
        smStdDemo(14) = "W50-54"
        smStdDemo(15) = "W55-64"
        smStdDemo(16) = "W65+"
        smStdDemo(17) = ""
    Else
        smStdDemo(0) = "M12-17"
        smStdDemo(1) = "M18-20"
        smStdDemo(2) = "M21-24"
        smStdDemo(3) = "M25-34"
        smStdDemo(4) = "M35-44"
        smStdDemo(5) = "M45-49"
        smStdDemo(6) = "M50-54"
        smStdDemo(7) = "M55-64"
        smStdDemo(8) = "M65+"
        smStdDemo(9) = "W12-17"
        smStdDemo(10) = "W18-20"
        smStdDemo(11) = "W21-24"
        smStdDemo(12) = "W25-34"
        smStdDemo(13) = "W35-44"
        smStdDemo(14) = "W45-49"
        smStdDemo(15) = "W50-54"
        smStdDemo(16) = "W55-64"
        smStdDemo(17) = "W65+"
    End If
    lmSDrfPopRecPos = 0
    lmCDrfPopRecPos = 0
    imPopChg = False
    imDnfChg = False
    imDrfChg = False
    imDpfChg = False
    imDefChg = False
    smTotalPop = ""
    lbcVehicle.ListIndex = -1
    lbcDaypart.ListIndex = -1
    lbcSocEco.ListIndex = -1
    lbcDays.ListIndex = -1
    smSource = ""
    ReDim tgDrfRec(0 To 1) As DRFREC
    ReDim tgDrfDel(0 To 1) As DRFREC
    ReDim tgAllDrf(0 To 1) As DRFREC
    ReDim tgDpfRec(0 To 1) As DPFREC
    ReDim tgDpfDel(0 To 1) As DPFREC
    ReDim tgAllDpf(0 To 1) As DPFREC
    ReDim tgDefRec(0 To 1) As DEFREC
    ReDim tgDefDel(0 To 1) As DEFREC
    
    ReDim tmSaveShow(0 To 1) As SAVESHOW
    For illoop = LBound(smSSave) To UBound(smSSave) Step 1
        smSSave(illoop) = ""
    Next illoop
    For illoop = imLBSCtrls To UBound(tmSCtrls) Step 1
        tmSCtrls(illoop).sShow = ""
        tmSCtrls(illoop).iChg = False
    Next illoop
    mInitNewDrf False, UBound(tgDrfRec)
    mInitNewDpf
    mInitNewDef
    If tgSpf.sDemoEstAllowed = "Y" Then
        lacPlusTitle.Caption = "Population Estimates"
    End If
    mComputeTotalPop
    imSettingValue = True
    vbcDemo.Min = imLBSaveShow   'LBound(tmSaveShow)
    imSettingValue = True
    
    If UBound(tmSaveShow) <= vbcDemo.LargeChange + 1 Then ' + 1 Then
        vbcDemo.Max = imLBSaveShow  'LBound(tmSaveShow)
    Else
        vbcDemo.Max = UBound(tmSaveShow) - vbcDemo.LargeChange '- 1
    End If
    imSettingValue = True
    vbcDemo.Value = vbcDemo.Min
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDaysPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Day list box with     *
'*                      standard days allowed          *
'*                                                     *
'*******************************************************
Private Sub mDaysPop()
    lbcDays.Clear
    lbcDays.AddItem "Mo-Fr"
    lbcDays.ItemData(lbcDays.NewIndex) = 0
    lbcDays.AddItem "Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 1
    lbcDays.AddItem "Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 2
    lbcDays.AddItem "Mo-Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 3
    lbcDays.AddItem "Mo-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 4
    lbcDays.AddItem "Sa-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 5
    lbcDays.AddItem "Tu-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 6
    lbcDays.AddItem "Tu-Fr"
    lbcDays.ItemData(lbcDays.NewIndex) = 7
    lbcDays.AddItem "We-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 8
    lbcDays.AddItem "Mo"
    lbcDays.ItemData(lbcDays.NewIndex) = 9
    lbcDays.AddItem "Tu"
    lbcDays.ItemData(lbcDays.NewIndex) = 10
    lbcDays.AddItem "We"
    lbcDays.ItemData(lbcDays.NewIndex) = 11
    lbcDays.AddItem "Th"
    lbcDays.ItemData(lbcDays.NewIndex) = 12
    lbcDays.AddItem "Fr"
    lbcDays.ItemData(lbcDays.NewIndex) = 13
    
    lbcDays.AddItem "Mo-Th"
    lbcDays.ItemData(lbcDays.NewIndex) = 14
    lbcDays.AddItem "Mo-We"
    lbcDays.ItemData(lbcDays.NewIndex) = 15
    lbcDays.AddItem "Mo-Tu"
    lbcDays.ItemData(lbcDays.NewIndex) = 16
    lbcDays.AddItem "Tu-Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 17
    lbcDays.AddItem "Tu-Th"
    lbcDays.ItemData(lbcDays.NewIndex) = 18
    lbcDays.AddItem "Tu-We"
    lbcDays.ItemData(lbcDays.NewIndex) = 19
    lbcDays.AddItem "We-Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 20
    lbcDays.AddItem "We-Fr"
    lbcDays.ItemData(lbcDays.NewIndex) = 21
    lbcDays.AddItem "We-Th"
    lbcDays.ItemData(lbcDays.NewIndex) = 22
    lbcDays.AddItem "Th-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 23
    lbcDays.AddItem "Th-Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 24
    lbcDays.AddItem "Th-Fr"
    lbcDays.ItemData(lbcDays.NewIndex) = 25
    lbcDays.AddItem "Fr-Su"
    lbcDays.ItemData(lbcDays.NewIndex) = 26
    lbcDays.AddItem "Fr-Sa"
    lbcDays.ItemData(lbcDays.NewIndex) = 27
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
    Dim ilRet As Integer
    Dim slType As String
    Dim illoop As Integer
    Dim slStr As String
    Dim slFontName As String
    Dim flFontSize As Single
    Dim slNameCode As String
    Dim slCode As String
    ReDim smStdDemo(0 To 17) As String
    smStdDemo(0) = "M12-17"
    smStdDemo(1) = "M18-24"
    smStdDemo(2) = "M25-34"
    smStdDemo(3) = "M35-44"
    smStdDemo(4) = "M45-49"
    smStdDemo(5) = "M50-54"
    smStdDemo(6) = "M55-64"
    smStdDemo(7) = "M65+"
    smStdDemo(8) = ""
    smStdDemo(9) = "W12-17"
    smStdDemo(10) = "W18-24"
    smStdDemo(11) = "W25-34"
    smStdDemo(12) = "W35-44"
    smStdDemo(13) = "W45-49"
    smStdDemo(14) = "W50-54"
    smStdDemo(15) = "W55-64"
    smStdDemo(16) = "W65+"
    smStdDemo(17) = ""
    slType = "DC"
    ilRet = gPopMnfPlusFieldsBox(Research, cbcDemo, tgCDemoCode(), sgCDemoCodeTag, slType)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox: Demo)", Research
        On Error GoTo 0
    End If
    
    ReDim tmCustInfo(0 To cbcDemo.ListCount) As CUSTINFO
    If cbcDemo.ListCount > 0 Then
        slFontName = pbcSpec.FontName
        flFontSize = pbcSpec.FontSize
        pbcSpec.FontBold = False
        pbcSpec.FontSize = 7
        pbcSpec.FontName = "Arial"
        pbcSpec.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        
        For illoop = 0 To cbcDemo.ListCount - 1 Step 1
            slStr = cbcDemo.List(illoop)
            slNameCode = tgCDemoCode(illoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmMnfSrchKey.iCode = Val(slCode)
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            gConvCustomGroup tmMnf.iGroupNo, tmCustInfo(illoop).sDataType, tmCustInfo(illoop).iDemoIndex
            gSetShow pbcSpec, slStr, tmSCtrls(POPINDEX)
            tmCustInfo(illoop).sName = tmSCtrls(POPINDEX).sShow
        Next illoop
        pbcSpec.FontSize = flFontSize
        pbcSpec.FontName = slFontName
        pbcSpec.FontSize = flFontSize
        pbcSpec.FontBold = True
        tmSCtrls(POPINDEX).sShow = ""
        cbcDemo.Visible = True
    Else
        cbcDemo.Visible = False
    End If
    cbcDemo.AddItem "[Standard Demo]", 0
    cbcDemo.ListIndex = 0
    slType = "DP"
    ilRet = gPopMnfPlusFieldsBox(Research, lbcPlusDemos, tmPlusDemoCode(), smPlusDemoCodeTag, slType)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox: Demo)", Research
        On Error GoTo 0
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDPPop                          *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Daypart list box with *
'*                      names                          *
'*                                                     *
'*******************************************************
Private Sub mDPPop(llRowNo As Long)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim illoop As Integer
    Dim slStr As String
    Dim ilFound As Integer
    If Trim$(tmSaveShow(llRowNo).sSave(1)) = "" Then
        lbcDaypart.Clear
        lbcDPCode.Clear
        Exit Sub
    End If
    gFindMatch Trim$(tmSaveShow(llRowNo).sSave(1)), 0, lbcVehicle
    If gLastFound(lbcVehicle) < 0 Then
        lbcDaypart.Clear
        lbcDPCode.Clear
        Exit Sub
    End If
    slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilRet = gParseItem(slNameCode, 1, "\", slName)
    ilRet = gParseItem(slName, 3, "|", slName)
    If lbcDPCode.Tag = slName Then
        Exit Sub
    End If
    lbcDaypart.Clear
    lbcDPCode.Clear
    lbcDPCode.Tag = slName
    ilVefCode = Val(slCode)
    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        If tgMRif(llRif).iVefCode = ilVefCode Then
            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
            If ilRdf <> -1 Then
                slStr = Trim$(tgMRdf(ilRdf).sName) & "\" & Trim$(Str$(tgMRdf(ilRdf).iCode))
                ilFound = False
                For illoop = 0 To lbcDPCode.ListCount - 1 Step 1
                    If StrComp(slStr, lbcDPCode.List(illoop), 1) = 0 Then
                        ilFound = True
                        Exit For
                    End If
                Next illoop
                If Not ilFound Then
                    lbcDPCode.AddItem slStr
                End If
            End If
        End If
    Next llRif
    For illoop = 0 To lbcDPCode.ListCount - 1 Step 1
        slNameCode = lbcDPCode.List(illoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        lbcDaypart.AddItem slName
    Next illoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slName As String
    If rbcDataType(0).Value Then 'Daypart
        If ilBoxNo < imLBDCtrls Or ilBoxNo > UBound(tmDCtrls) Then
            Exit Sub
        End If
        'If (lmRowNo < (vbcDemo.Value + 1) \ 2) Or (lmRowNo > ((vbcDemo.Value + 1) \ 2 + vbcDemo.LargeChange \ 2)) Then
        If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
            'mSetShow ilBoxNo
            pbcArrow.Visible = False
            lacFrame(0).Visible = False
            mShowDpf False
            Exit Sub
        End If
        lacFrame(0).Move 0, tmDCtrls(DVEHICLEINDEX).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
        lacFrame(0).Visible = True
        pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmDCtrls(DVEHICLEINDEX).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        cmcDuplicate.Enabled = True
        
        Select Case ilBoxNo
            Case DVEHICLEINDEX
                lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
                If smSource <> "I" Then 'Standard Airtime mode
                    edcDropDown.Width = tmDCtrls(DVEHICLEINDEX).fBoxW
                Else 'Podcast Impression mode
                    edcDropDown.Width = tmDCtrls(DVEHICLEINDEX).fBoxW + mAct1ColsWidth
                End If
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)), 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    If lmRowNo > 1 Then
                        gFindMatch Trim$(tmSaveShow(lmRowNo - 1).sSave(DVEHICLEINDEX)), 0, lbcVehicle
                        If (gLastFound(lbcVehicle) >= 0) Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            slName = sgUserDefVehicleName
                            gFindMatch slName, 0, lbcVehicle
                            If gLastFound(lbcVehicle) >= 0 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            Else
                                lbcVehicle.ListIndex = 0
                            End If
                        End If
                    Else
                        slName = sgUserDefVehicleName
                        gFindMatch slName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            lbcVehicle.ListIndex = 0
                        End If
                    End If
                End If
                imComboBoxIndex = lbcVehicle.ListIndex
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            
            Case DACT1CODEINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    edcDropDown.Width = tmDCtrls(ilBoxNo).fBoxW
                    edcDropDown.MaxLength = 11
                    gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                    edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(DACT1CODEINDEX))
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                End If
                
            Case DACT1SETTINGINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    plcACT1Settings.Visible = True
                    edcDropDown.MaxLength = 4
                    gMoveTableCtrl pbcDemo(0), plcACT1Settings, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                    edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(DACT1SETTINGINDEX))
                    If InStr(1, edcDropDown.Text, "T") > 0 Then
                        edcACT1SettingT.Text = "Yes"
                    Else
                        edcACT1SettingT.Text = "No"
                    End If
                    If InStr(1, edcDropDown.Text, "S") > 0 Then
                        edcACT1SettingS.Text = "Yes"
                    Else
                        edcACT1SettingS.Text = "No"
                    End If
                    If InStr(1, edcDropDown.Text, "C") > 0 Then
                        edcACT1SettingC.Text = "Yes"
                    Else
                        edcACT1SettingC.Text = "No"
                    End If
                    If InStr(1, edcDropDown.Text, "F") > 0 Then
                        edcACT1SettingF.Text = "Yes"
                    Else
                        edcACT1SettingF.Text = "No"
                    End If
                    edcACT1SettingT.SetFocus
                End If
                
            Case DDAYPARTINDEX
                mDPPop lmRowNo
                lbcDaypart.Height = gListBoxHeight(lbcDaypart.ListCount, 6)
                edcDropDown.Width = tmDCtrls(DDAYPARTINDEX).fBoxW '- cmcDropDown.Width
                edcDropDown.MaxLength = 30
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(DDAYPARTINDEX).fBoxX, tmDCtrls(DDAYPARTINDEX).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcDaypart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcDaypart.Move edcDropDown.Left, edcDropDown.Top - lbcDaypart.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(DDAYPARTINDEX)), 0, lbcDaypart
                If gLastFound(lbcDaypart) >= 0 Then
                    lbcDaypart.ListIndex = gLastFound(lbcDaypart)
                Else
                    If lbcDaypart.ListCount > 0 Then
                        If lmRowNo > 1 Then
                            If (Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) = Trim$(tmSaveShow(lmRowNo - 1).sSave(DVEHICLEINDEX))) Then
                                gFindMatch Trim$(tmSaveShow(lmRowNo - 1).sSave(DDAYPARTINDEX)), 0, lbcDaypart
                                If (gLastFound(lbcDaypart) >= 0) Then
                                    If gLastFound(lbcDaypart) < lbcDaypart.ListCount - 1 Then
                                        lbcDaypart.ListIndex = gLastFound(lbcDaypart) + 1
                                    Else
                                        lbcDaypart.ListIndex = 0
                                    End If
                                Else
                                    lbcDaypart.ListIndex = 0
                                End If
                            Else
                                lbcDaypart.ListIndex = 0
                            End If
                        Else
                            lbcDaypart.ListIndex = 0
                        End If
                    End If
                End If
                imComboBoxIndex = lbcDaypart.ListIndex
                If lbcDaypart.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDaypart.List(lbcDaypart.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            
            Case DGROUPINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    lbcSocEco.Height = gListBoxHeight(lbcSocEco.ListCount, 6)
                    edcDropDown.Width = tmDCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
                    edcDropDown.MaxLength = 6
                    gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                    cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                    If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                        lbcSocEco.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                    Else
                        lbcSocEco.Move edcDropDown.Left, edcDropDown.Top - lbcSocEco.Height
                    End If
                    imChgMode = True
                    gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(DAIRTIMEGRPNOINDEX)), 1, lbcSocEco
                    If gLastFound(lbcSocEco) > 0 Then
                        lbcSocEco.ListIndex = gLastFound(lbcSocEco)
                    Else
                        lbcSocEco.ListIndex = 0   '[None]
                    End If
                    imComboBoxIndex = lbcSocEco.ListIndex
                    If lbcSocEco.ListIndex < 0 Then
                        edcDropDown.Text = ""
                    Else
                        edcDropDown.Text = lbcSocEco.List(lbcSocEco.ListIndex)
                    End If
                    imChgMode = False
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                    edcDropDown.Visible = True
                    cmcDropDown.Visible = True
                    edcDropDown.SetFocus
                Else 'Podcast Impression mode
                    edcDropDown.Width = tmDCtrls(ilBoxNo).fBoxW
                    edcDropDown.MaxLength = 7
                    gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                    edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(DIMPRESSIONSINDEX))
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                End If
                
            Case DDEMOINDEX To DDEMOINDEX + 17
                edcDropDown.Width = tmDCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 7
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmDCtrls(ilBoxNo).fBoxX, tmDCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(DDEMOINDEX + ilBoxNo - DDEMOINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
        
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        If ilBoxNo < imLBXCtrls Or ilBoxNo > UBound(tmXCtrls) Then
            Exit Sub
        End If
        If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
            pbcArrow.Visible = False
            lacFrame(2).Visible = False
            mShowDpf False
            Exit Sub
        End If
        lacFrame(2).Move 0, tmXCtrls(XVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
        lacFrame(2).Visible = True
        pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmXCtrls(XVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        Select Case ilBoxNo
            Case XVEHICLEINDEX
                lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
                edcDropDown.Width = tmXCtrls(XVEHICLEINDEX).fBoxW
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)), 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    If lmRowNo > 1 Then
                        gFindMatch Trim$(tmSaveShow(lmRowNo - 1).sSave(XVEHICLEINDEX)), 0, lbcVehicle
                        If (gLastFound(lbcVehicle) >= 0) Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            slName = sgUserDefVehicleName
                            gFindMatch slName, 0, lbcVehicle
                            If gLastFound(lbcVehicle) >= 0 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            Else
                                lbcVehicle.ListIndex = 0
                            End If
                        End If
                    Else
                        slName = sgUserDefVehicleName
                        gFindMatch slName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            lbcVehicle.ListIndex = 0
                        End If
                    End If
                End If
                imComboBoxIndex = lbcVehicle.ListIndex
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            
            Case XACT1CODEINDEX
                edcDropDown.Width = tmXCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 11
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(XACT1CODEINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            
            Case XACT1SETTINGINDEX
                plcACT1Settings.Visible = True
                edcDropDown.MaxLength = 4
                gMoveTableCtrl pbcDemo(0), plcACT1Settings, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(XACT1SETTINGINDEX))
                If InStr(1, edcDropDown.Text, "T") > 0 Then
                    edcACT1SettingT.Text = "Yes"
                Else
                    edcACT1SettingT.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "S") > 0 Then
                    edcACT1SettingS.Text = "Yes"
                Else
                    edcACT1SettingS.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "C") > 0 Then
                    edcACT1SettingC.Text = "Yes"
                Else
                    edcACT1SettingC.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "F") > 0 Then
                    edcACT1SettingF.Text = "Yes"
                Else
                    edcACT1SettingF.Text = "No"
                End If
                edcACT1SettingT.SetFocus
            
            Case XTIMEINDEX To XTIMEINDEX + 1
                edcDropDown.MaxLength = 10
                edcDropDown.Width = tmXCtrls(ilBoxNo).fBoxW
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
                End If
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(XTIMEINDEX + ilBoxNo - XTIMEINDEX))
                plcTme.Visible = False
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case XDAYSINDEX
                lbcDays.Height = gListBoxHeight(lbcDays.ListCount, 6)
                edcDropDown.Width = tmXCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top - lbcDays.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(XDAYSINDEX)), 1, lbcDays
                If gLastFound(lbcDays) > 0 Then
                    lbcDays.ListIndex = gLastFound(lbcDays)
                Else
                    lbcDays.ListIndex = 0   'M-F
                End If
                imComboBoxIndex = lbcDays.ListIndex
                If lbcDays.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case XGROUPINDEX
                lbcSocEco.Height = gListBoxHeight(lbcSocEco.ListCount, 6)
                edcDropDown.Width = tmXCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
                edcDropDown.MaxLength = 6
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcSocEco.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcSocEco.Move edcDropDown.Left, edcDropDown.Top - lbcSocEco.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(XGROUPNINDEX)), 1, lbcSocEco
                If gLastFound(lbcSocEco) > 0 Then
                    lbcSocEco.ListIndex = gLastFound(lbcSocEco)
                Else
                    lbcSocEco.ListIndex = 0   '[None]
                End If
                imComboBoxIndex = lbcSocEco.ListIndex
                If lbcSocEco.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSocEco.List(lbcSocEco.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case XDEMOINDEX To XDEMOINDEX + 17
                edcDropDown.Width = tmXCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 7
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmXCtrls(ilBoxNo).fBoxX, tmXCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(XDEMOINDEX + ilBoxNo - XDEMOINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
        
    ElseIf rbcDataType(2).Value Then 'Time
        If ilBoxNo < imLBTCtrls Or ilBoxNo > UBound(tmTCtrls) Then
            Exit Sub
        End If
        If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
            pbcArrow.Visible = False
            lacFrame(0).Visible = False
            mShowDpf False
            Exit Sub
        End If
        lacFrame(0).Move 0, tmTCtrls(TVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
        lacFrame(0).Visible = True
        pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmTCtrls(TVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        Select Case ilBoxNo
            Case TVEHICLEINDEX
                lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
                edcDropDown.Width = tmTCtrls(TVEHICLEINDEX).fBoxW
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmTCtrls(ilBoxNo).fBoxX, tmTCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(TVEHICLEINDEX)), 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    If lmRowNo > 1 Then
                        gFindMatch Trim$(tmSaveShow(lmRowNo - 1).sSave(TVEHICLEINDEX)), 0, lbcVehicle
                        If (gLastFound(lbcVehicle) >= 0) Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            slName = sgUserDefVehicleName
                            gFindMatch slName, 0, lbcVehicle
                            If gLastFound(lbcVehicle) >= 0 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            Else
                                lbcVehicle.ListIndex = 0
                            End If
                        End If
                    Else
                        slName = sgUserDefVehicleName
                        gFindMatch slName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            lbcVehicle.ListIndex = 0
                        End If
                    End If
                End If
                imComboBoxIndex = lbcVehicle.ListIndex
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case TTIMEINDEX To TTIMEINDEX + 1
                edcDropDown.MaxLength = 10
                edcDropDown.Width = tmTCtrls(ilBoxNo).fBoxW
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmTCtrls(ilBoxNo).fBoxX, tmTCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
                End If
                If (TTIMEINDEX + 1 = ilBoxNo) And (tmSaveShow(lmRowNo).sSave(TTIMEINDEX + 1 + ilBoxNo - TTIMEINDEX) = "") Then
                    tmSaveShow(lmRowNo).sSave(TTIMEINDEX + 1 + ilBoxNo - TTIMEINDEX) = tmSaveShow(lmRowNo).sSave(TTIMEINDEX + 1 + ilBoxNo - TTIMEINDEX - 1)
                End If
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(TTIMEINDEX + ilBoxNo - TTIMEINDEX))
                plcTme.Visible = False
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case TDAYSINDEX
                lbcDays.Height = gListBoxHeight(lbcDays.ListCount, 6)
                edcDropDown.Width = tmTCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmTCtrls(ilBoxNo).fBoxX, tmTCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top - lbcDays.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(TDAYSINDEX)), 1, lbcDays
                If gLastFound(lbcDays) > 0 Then
                    lbcDays.ListIndex = gLastFound(lbcDays)
                Else
                    lbcDays.ListIndex = 0   'M-F
                End If
                imComboBoxIndex = lbcDays.ListIndex
                If lbcDays.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case TDEMOINDEX To TDEMOINDEX + 17
                edcDropDown.Width = tmTCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 7
                gMoveTableCtrl pbcDemo(0), edcDropDown, tmTCtrls(ilBoxNo).fBoxX, tmTCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(TDEMOINDEX + ilBoxNo - TDEMOINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
        
    Else 'Vehicle
        If ilBoxNo < imLBVCtrls Or ilBoxNo > UBound(tmVCtrls) Then
            Exit Sub
        End If
        If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
            pbcArrow.Visible = False
            lacFrame(1).Visible = False
            mShowDpf False
            Exit Sub
        End If
        lacFrame(1).Move 0, tmVCtrls(VVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
        lacFrame(1).Visible = True
        pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmVCtrls(VVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        Select Case ilBoxNo
            Case VVEHICLEINDEX
                lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
                edcDropDown.Width = tmVCtrls(VVEHICLEINDEX).fBoxW
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcDemo(1), edcDropDown, tmVCtrls(ilBoxNo).fBoxX, tmVCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)), 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    If lmRowNo > 1 Then
                        gFindMatch Trim$(tmSaveShow(lmRowNo - 1).sSave(VVEHICLEINDEX)), 0, lbcVehicle
                        If (gLastFound(lbcVehicle) >= 0) Then
                            If gLastFound(lbcVehicle) < lbcVehicle.ListCount - 1 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle) + 1
                            Else
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            End If
                        Else
                            slName = sgUserDefVehicleName
                            gFindMatch slName, 0, lbcVehicle
                            If gLastFound(lbcVehicle) >= 0 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            Else
                                lbcVehicle.ListIndex = 0
                            End If
                        End If
                    Else
                        slName = sgUserDefVehicleName
                        gFindMatch slName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        Else
                            lbcVehicle.ListIndex = 0
                        End If
                    End If
                End If
                imComboBoxIndex = lbcVehicle.ListIndex
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case VACT1CODEINDEX
                edcDropDown.Width = tmVCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 11
                gMoveTableCtrl pbcDemo(1), edcDropDown, tmVCtrls(ilBoxNo).fBoxX, tmVCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(VACT1CODEINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            
            Case VACT1SETTINGINDEX
                plcACT1Settings.Visible = True
                edcDropDown.MaxLength = 4
                gMoveTableCtrl pbcDemo(1), plcACT1Settings, tmVCtrls(ilBoxNo).fBoxX, tmVCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(VACT1SETTINGINDEX))
                If InStr(1, edcDropDown.Text, "T") > 0 Then
                    edcACT1SettingT.Text = "Yes"
                Else
                    edcACT1SettingT.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "S") > 0 Then
                    edcACT1SettingS.Text = "Yes"
                Else
                    edcACT1SettingS.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "C") > 0 Then
                    edcACT1SettingC.Text = "Yes"
                Else
                    edcACT1SettingC.Text = "No"
                End If
                If InStr(1, edcDropDown.Text, "F") > 0 Then
                    edcACT1SettingF.Text = "Yes"
                Else
                    edcACT1SettingF.Text = "No"
                End If
                edcACT1SettingT.SetFocus
            
            Case VDAYSINDEX
                lbcDays.Height = gListBoxHeight(lbcDays.ListCount, 6)
                edcDropDown.Width = tmVCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcDemo(1), edcDropDown, tmVCtrls(ilBoxNo).fBoxX, tmVCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If lmRowNo - vbcDemo.Value <= vbcDemo.LargeChange \ 2 Then
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top - lbcDays.Height
                End If
                imChgMode = True
                gFindMatch Trim$(tmSaveShow(lmRowNo).sSave(VDAYSINDEX)), 1, lbcDays
                If gLastFound(lbcDays) > 0 Then
                    lbcDays.ListIndex = gLastFound(lbcDays)
                Else
                    lbcDays.ListIndex = 0   'M-F
                End If
                imComboBoxIndex = lbcDays.ListIndex
                If lbcDays.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
                
            Case VDEMOINDEX To VDEMOINDEX + 17
                edcDropDown.Width = tmVCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 7
                gMoveTableCtrl pbcDemo(1), edcDropDown, tmVCtrls(ilBoxNo).fBoxX, tmVCtrls(ilBoxNo).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(VDEMOINDEX + ilBoxNo - VDEMOINDEX))
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
    End If
    mShowDpf False
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetRec                         *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move records to be viewed      *
'*                                                     *
'*******************************************************
Private Sub mGetRec()
    Dim slInfoType As String
    Dim slDataType As String
    Dim ilRecOK As Integer
    Dim llLoop As Long
    Dim llUpper As Long
    Dim llTest As Long
    Dim ilRdf As Integer
    Dim ilVef As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim llTime As Long
    Dim ilDay As Integer
    Dim ilSDay As Integer
    Dim ilEDay As Integer
    Dim ilCustInfoIndex As Integer
    Dim illoop As Integer
    Dim ilDemo As Integer
    Dim llLink As Long
    Dim slStr As String
    Dim slVefName As String * 40
    ReDim tgDrfRec(0 To UBound(tgAllDrf))
    ReDim tgLinkDrfRec(0 To 1) As DRFREC
    If imCustomIndex <= 0 Then
        slDataType = "A"
    Else
        For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
            If illoop < UBound(tmCustInfo) Then
                slStr = Trim$(tmCustInfo(illoop).sName)
                smCustomDemo(illoop - (imCustomIndex - 1)) = slStr  'tmSCtrls(POPINDEX).sShow
            Else
                smCustomDemo(illoop - (imCustomIndex - 1)) = ""
            End If
        Next illoop
    End If
    If (rbcDataType(0).Value) Or (rbcDataType(1).Value) Then 'Daypart or Extra Daypart
        slInfoType = "D"
    ElseIf rbcDataType(2).Value Then 'Time
        slInfoType = "T"
    Else 'Vehicle
        slInfoType = "V"
    End If
    
    llUpper = imLBDrf   'LBound(tgDrfRec)
    For llLoop = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        ilRecOK = False
        If tgAllDrf(llLoop).iStatus <> -1 Then
            If imCustomIndex <= 0 Then
                If (tgAllDrf(llLoop).tDrf.sInfoType = slInfoType) And (tgAllDrf(llLoop).tDrf.sDataType = slDataType) Then
                    ilRecOK = True
                    ilCustInfoIndex = 0
                End If
            Else
                For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
                    If illoop < UBound(tmCustInfo) Then
                        If (tgAllDrf(llLoop).tDrf.sInfoType = slInfoType) And (tgAllDrf(llLoop).tDrf.sDataType = tmCustInfo(illoop).sDataType) Then
                            ilRecOK = True
                            ilCustInfoIndex = illoop
                            Exit For
                        End If
                    End If
                Next illoop
            End If
            If ilRecOK Then
                ilRecOK = False
                If (rbcDataType(0).Value) And (tgAllDrf(llLoop).tDrf.iRdfCode > 0) And (tgAllDrf(llLoop).tDrf.iCount >= 0) Then 'Daypart
                    ilRecOK = True
                ElseIf (rbcDataType(1).Value) And (tgAllDrf(llLoop).tDrf.iRdfCode = 0) And (tgAllDrf(llLoop).tDrf.iCount >= 0) Then 'Extra Daypart
                    ilRecOK = True
                ElseIf rbcDataType(2).Value And (tgAllDrf(llLoop).tDrf.sExStdDP <> "Y") And (tgAllDrf(llLoop).tDrf.sExStdDP <> "X") Then 'Time
                    ilRecOK = True
                ElseIf rbcDataType(3).Value Then 'Vehicle
                    ilRecOK = True
                End If
            End If
        End If
        If ilRecOK Then
            If imCustomIndex <= 0 Then
                mMoveAllToDrfRec llLoop, llUpper
            Else
                Do
                    ilRecOK = False
                    For llTest = imLBDrf To llUpper - 1 Step 1
                        If (tgDrfRec(llTest).tDrf.sDataType = tgAllDrf(llLoop).tDrf.sDataType) Then
                            If (tgDrfRec(llTest).tDrf.iVefCode = tgAllDrf(llLoop).tDrf.iVefCode) Then
                                If (tgDrfRec(llTest).tDrf.iStartTime(0) = tgAllDrf(llLoop).tDrf.iStartTime(0)) And (tgDrfRec(llTest).tDrf.iStartTime(1) = tgAllDrf(llLoop).tDrf.iStartTime(1)) Then
                                    If (tgDrfRec(llTest).tDrf.iEndTime(0) = tgAllDrf(llLoop).tDrf.iEndTime(0)) And (tgDrfRec(llTest).tDrf.iEndTime(1) = tgAllDrf(llLoop).tDrf.iEndTime(1)) Then
                                        If (tgDrfRec(llTest).tDrf.iRdfCode = tgAllDrf(llLoop).tDrf.iRdfCode) And (tgDrfRec(llTest).tDrf.sExStdDP = tgAllDrf(llLoop).tDrf.sExStdDP) Then
                                            ilRecOK = True
                                            For ilDay = 0 To 6 Step 1
                                                If tgDrfRec(llTest).tDrf.sDay(ilDay) <> tgAllDrf(llLoop).tDrf.sDay(ilDay) Then
                                                    ilRecOK = False
                                                End If
                                            Next ilDay
                                            If ilRecOK Then
                                                tgDrfRec(llTest).tDrf.lCode = tgAllDrf(llLoop).tDrf.lCode
                                                tgDrfRec(llTest).lIndex = llLoop
                                                For ilDemo = 1 To 18 Step 1
                                                    tgDrfRec(llTest).tDrf.lDemo(ilDemo - 1) = tgAllDrf(llLoop).tDrf.lDemo(ilDemo - 1)
                                                Next ilDemo
                                                tgAllDrf(llLoop).iStatus = -1
                                                tgAllDrf(llLoop).iModel = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Not ilRecOK Then
                            llLink = tgDrfRec(llTest).lLink
                            Do While llLink <> -1
                                If (tgAllDrf(llLoop).tDrf.sDataType = tgLinkDrfRec(llLink).tDrf.sDataType) Then
                                    If (tgLinkDrfRec(llLink).tDrf.iVefCode = tgAllDrf(llLoop).tDrf.iVefCode) Then
                                        If (tgLinkDrfRec(llLink).tDrf.iStartTime(0) = tgAllDrf(llLoop).tDrf.iStartTime(0)) And (tgLinkDrfRec(llLink).tDrf.iStartTime(1) = tgAllDrf(llLoop).tDrf.iStartTime(1)) Then
                                            If (tgLinkDrfRec(llLink).tDrf.iEndTime(0) = tgAllDrf(llLoop).tDrf.iEndTime(0)) And (tgLinkDrfRec(llLink).tDrf.iEndTime(1) = tgAllDrf(llLoop).tDrf.iEndTime(1)) Then
                                                If (tgLinkDrfRec(llLink).tDrf.iRdfCode = tgAllDrf(llLoop).tDrf.iRdfCode) And (tgLinkDrfRec(llLink).tDrf.sExStdDP = tgAllDrf(llLoop).tDrf.sExStdDP) Then
                                                    ilRecOK = True
                                                    For ilDay = 0 To 6 Step 1
                                                        If tgLinkDrfRec(llLink).tDrf.sDay(ilDay) <> tgAllDrf(llLoop).tDrf.sDay(ilDay) Then
                                                            ilRecOK = False
                                                        End If
                                                    Next ilDay
                                                    If ilRecOK Then
                                                        tgLinkDrfRec(llLink).tDrf.lCode = tgAllDrf(llLoop).tDrf.lCode
                                                        tgLinkDrfRec(llLink).lIndex = llLoop
                                                        For ilDemo = 1 To 18 Step 1
                                                            tgLinkDrfRec(llLink).tDrf.lDemo(ilDemo - 1) = tgAllDrf(llLoop).tDrf.lDemo(ilDemo - 1)
                                                        Next ilDemo
                                                        tgAllDrf(llLoop).iStatus = -1
                                                        tgAllDrf(llLoop).iModel = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                If ilRecOK Then
                                    Exit Do
                                End If
                                llLink = tgLinkDrfRec(llLink).lLink
                            Loop
                        End If
                    Next llTest
                    If Not ilRecOK Then
                        mMoveAllToLinkDrfRec llLoop, llUpper
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next llLoop
    ReDim Preserve tgDrfRec(0 To llUpper) As DRFREC
    mInitNewDrf False, UBound(tgDrfRec)
    If UBound(tgDrfRec) - 1 > 1 Then
        ReDim tmSortDrfRec(0 To UBound(tgDrfRec) - 1) As DRFREC
        For llLoop = 0 To UBound(tmSortDrfRec) Step 1
            tmSortDrfRec(llLoop) = tgDrfRec(llLoop + 1)
        Next llLoop
        ArraySortTyp fnAV(tmSortDrfRec(), 0), UBound(tmSortDrfRec), 0, LenB(tmSortDrfRec(0)), 0, LenB(tmSortDrfRec(0).sKey), 0
        For llLoop = UBound(tmSortDrfRec) To 0 Step -1
            tgDrfRec(llLoop + 1) = tmSortDrfRec(llLoop)
        Next llLoop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    bmResearchSaved = False
    imDataType = 3 'Defaults to Vehicle view

    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    bmIgnoreChg = False
    imLBSCtrls = 1
    imLBDCtrls = 1
    imLBXCtrls = 1
    imLBTCtrls = 1
    imLBVCtrls = 1
    imLBPCtrls = 1
    imLBCDCtrls = 1
    imLBDrf = 1
    imLBDpf = 1
    imLBDef = 1
    imLBMnf = 1
    imLBSaveShow = 1
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    ReDim smStdDemo(0 To 0) As String
    ReDim tgDrfRec(0 To 1) As DRFREC
    ReDim tgDrfDel(0 To 1) As DRFREC
    ReDim tgAllDrf(0 To 1) As DRFREC
    ReDim tgDpfRec(0 To 1) As DPFREC
    ReDim tgDpfDel(0 To 1) As DPFREC
    ReDim tgAllDpf(0 To 1) As DPFREC
    ReDim tgDefRec(0 To 1) As DEFREC
    ReDim tgDefDel(0 To 1) As DEFREC
    ReDim tgLinkDrfRec(0 To 1) As DRFREC
    ReDim tgCDrfPop(0 To 1) As DRF
    ReDim tgGDrfPop(0 To 1) As DRF
    ReDim tmSaveShow(0 To 1) As SAVESHOW
    
    mInitNewDrf True, UBound(tgDrfRec)
    mInitNewDpf
    mInitNewDef
    mInitBox
    gCenterStdAlone Research
    smTotalPop = ""
    imSelectedIndex = -1
    imCustomIndex = -1
    imTestAddStdDemo = True
    hmMnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", Research
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", Research
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    mGetP12Plus
    ilRet = mAddStdDemo()
    lbcDemo.Clear
    smPlusDemoCodeTag = ""
    mDemoPop
    'Research.Show
    Screen.MousePointer = vbHourglass
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imInNewTab = False
    imBoxNo = -1 'Initialize current Box to N/A
    lmRowNo = -1
    imDnfChg = False
    imDrfChg = False
    imPopChg = False
    imDpfChg = False
    imDefChg = False
    lmPlusRowNo = -1
    imPlusBoxNo = -1
    lmPlusRowNo = -1
    imBypassFocus = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imSettingValue = False
    imDragType = -1
    igDnfModel = 0
    imCalType = 0   'Standard
    hmDnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dnf.Btr)", Research
    On Error GoTo 0
    imDnfRecLen = Len(tgDnf)
    hmDrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", Research
    On Error GoTo 0
    imDrfRecLen = Len(tgDrfRec(1).tDrf)
    hmDpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", Research
    On Error GoTo 0
    imDpfRecLen = Len(tmDpf)
    hmDef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", Research
    On Error GoTo 0
    imDefRecLen = Len(tmDef)
    cbcSelect.Clear 'Force list box to be populated
    smFilterStartDate = ""
    lmFilterStartDate = 0
    smFilterEndDate = ""
    lmFilterEndDate = 0
    imFilterVefCode = -1
    smBNCodeTag = ""
    Screen.MousePointer = vbHourglass
    lbcVehicle.Clear 'Force list box to be populated
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    mDaysPop
    ilRet = mObtainSocEco()
    If Not ilRet Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainRcfRifRdf()
    If Not ilRet Then
        imTerminate = True
        Exit Sub
    End If
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    smMonDate = gObtainPrevMonday(slDate)
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, imCurMonth, imCurYear
    imIgnoreRightMove = False
    imEstByLOrU = 0
    If tgSpf.sDemoEstAllowed = "Y" Then
        pbcEstByLorU.Visible = True
        pbcDPorEst_KeyPress Asc("E")
    Else
        pbcDPorEst_KeyPress Asc("P")
    End If
    CSI_ComboBoxMS1.BackColor = &HFFFF00
    CSI_ComboBoxMS1.FontBold = True
    cmcGetBook.Visible = False
    cmcGetBook_Click
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim illoop As Integer
    Dim llMax As Long
    Dim llGap As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llHeight As Long
    Dim llWidthChg As Long
    Dim llLargeChg As Long
    Dim flDemoFieldWidth As Single
    Dim ilDemoStart As Integer
    Dim llWidth As Long
    Dim ilxPos As Single
    Dim ilyPos As Single
    Dim ilWidth As Single
    
    flDemoFieldWidth = pbcSpec.TextWidth("WWWWW.W")  'start size = 570
    flTextHeight = pbcSpec.TextHeight("1") - 35
    cbcSelect.Move 4635, 15
    
    'Position panel and picture areas with panel
    plcSpec.Move 195, cbcSelect.Top + cbcSelect.Height + 30, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    ckcSocEco.Move 195, plcSpec.Top + plcSpec.Height + 30
    plcDataType.Move 4380, ckcSocEco.Top
    pbcKey.Move plcSpec.Left, plcSpec.Top
    plcDemo.Move 195, ckcSocEco.Top + ckcSocEco.Height + 30, pbcDemo(0).Width + fgPanelAdj + vbcDemo.Width, pbcDemo(0).Height + fgPanelAdj
    pbcDemo(0).Move plcDemo.Left + fgBevelX, plcDemo.Top + fgBevelY
    pbcDemo(1).Move pbcDemo(0).Left, pbcDemo(0).Top
    pbcDemo(2).Move pbcDemo(0).Left, pbcDemo(0).Top
    vbcDemo.Move pbcDemo(0).Left + pbcDemo(0).Width - 15, pbcDemo(0).Top
    cbcDemo.Move plcDemo.Left + plcDemo.Width - cbcDemo.Width, plcDemo.Top + plcDemo.Height + 30
    plcPlus.Move 195, Research.Height - pbcPlus.Height - 3 * fgPanelAdj, pbcPlus.Width + fgPanelAdj + vbcPlus.Width, pbcPlus.Height + fgPanelAdj
    pbcPlus.Move plcPlus.Left + fgBevelX, plcPlus.Top + fgBevelY
    pbcDPorEst.Move plcPlus.Left, plcPlus.Top - pbcDPorEst.Height, plcPlus.Width
    pbcEst.Move pbcPlus.Left, pbcPlus.Top
    pbcUSA.Move pbcPlus.Left, pbcPlus.Top
    If tgSpf.sDemoEstAllowed <> "Y" Then
        lacPlusTitle.Caption = "Pre-defined Dayparts"
        pbcEst.Visible = False
        pbcDPorEst.Visible = False
    Else
        pbcPlus.Visible = False
        pbcDPorEst.Visible = True
        lacPlusTitle.Visible = False
        imDPorEst = 1
    End If
    lacPlusTitle.Move 195, plcPlus.Top - lacPlusTitle.Height - 30 ' + 30

    vbcPlus.Move pbcPlus.Left + pbcPlus.Width, pbcPlus.Top + 15
    pbcArrow.Move plcDemo.Left - pbcArrow.Width - 15    'set arrow    'Vehicle
    'Name
    gSetCtrl tmSCtrls(NAMEINDEX), 30, 30, 1815, fgBoxStH
    'Book Date
    gSetCtrl tmSCtrls(DATEINDEX), 1860, tmSCtrls(NAMEINDEX).fBoxY, 705, fgBoxStH
    'Population Source
    gSetCtrl tmSCtrls(POPSRCEINDEX), 30, tmSCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1260, fgBoxStH
    'Qual Population Source
    gSetCtrl tmSCtrls(QUALPOPSRCEINDEX), 1305, tmSCtrls(POPSRCEINDEX).fBoxY, 1260, fgBoxStH
    'Population
    ilDemoStart = 3330 '+ (ilLoop - POPINDEX) * 585  'Original
    For illoop = POPINDEX To POPINDEX + 17
        If illoop <= POPINDEX + 8 Then
            gSetCtrl tmSCtrls(illoop), 3330 + (illoop - POPINDEX) * (flDemoFieldWidth + 15), 375, flDemoFieldWidth, fgBoxGridH
        Else
            gSetCtrl tmSCtrls(illoop), 3330 + (illoop - (POPINDEX + 9)) * (flDemoFieldWidth + 15), 570, flDemoFieldWidth, fgBoxGridH
        End If
    Next illoop
    
    '--------------------------------------------------------------------------
    'Sold Daypart
    'Vehicle
    ilxPos = 30: ilyPos = 375: ilWidth = 4220
    gSetCtrl tmDCtrls(DVEHICLEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'DACT1CODE
    ilxPos = tmDCtrls(DVEHICLEINDEX).fBoxX + tmDCtrls(DVEHICLEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 1245
    gSetCtrl tmDCtrls(DACT1CODEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    mAct1ColsWidth = ilWidth + 15
    'DACT1SETTING
    ilxPos = tmDCtrls(DACT1CODEINDEX).fBoxX + tmDCtrls(DACT1CODEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 555
    gSetCtrl tmDCtrls(DACT1SETTINGINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    mAct1ColsWidth = mAct1ColsWidth + ilWidth + 15
    'Daypart or Times
    ilxPos = tmDCtrls(DACT1SETTINGINDEX).fBoxX + tmDCtrls(DACT1SETTINGINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 1350
    gSetCtrl tmDCtrls(DDAYPARTINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Group or Days
    ilxPos = tmDCtrls(DDAYPARTINDEX).fBoxX + tmDCtrls(DDAYPARTINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 555
        gSetCtrl tmDCtrls(DGROUPINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Demo
    ilxPos = tmDCtrls(DGROUPINDEX).fBoxX + tmDCtrls(DGROUPINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 795
    For illoop = DDEMOINDEX To DDEMOINDEX + 17
        If illoop <= DDEMOINDEX + 8 Then
            'Row 1 of Demo (M)
            gSetCtrl tmDCtrls(illoop), ilxPos + ((illoop - DDEMOINDEX) * (ilWidth + 15)), ilyPos, ilWidth, fgBoxGridH
        Else
            'Row 2 of Demo (W)
            gSetCtrl tmDCtrls(illoop), ilxPos + ((illoop - (DDEMOINDEX + 9)) * (ilWidth + 15)), ilyPos + 195, ilWidth, fgBoxGridH
        End If
    Next illoop
    
    '--------------------------------------------------------------------------
    'Extra Daypart
    'Vehicle
    ilxPos = 30: ilyPos = 375: ilWidth = 3995
    gSetCtrl tmXCtrls(XVEHICLEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'ACT1CODE
    ilxPos = tmXCtrls(XVEHICLEINDEX).fBoxX + tmXCtrls(XVEHICLEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 1245
    gSetCtrl tmXCtrls(XACT1CODEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'ACT1SETTING
    ilxPos = tmXCtrls(XACT1CODEINDEX).fBoxX + tmXCtrls(XACT1CODEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 555
    gSetCtrl tmXCtrls(XACT1SETTINGINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Time1
    ilxPos = tmXCtrls(XACT1SETTINGINDEX).fBoxX + tmXCtrls(XACT1SETTINGINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 495
    gSetCtrl tmXCtrls(XTIMEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH '1350, fgBoxGridH
    'Time2
    ilxPos = tmXCtrls(XTIMEINDEX).fBoxX + tmXCtrls(XTIMEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 495
    gSetCtrl tmXCtrls(XTIMEINDEX + 1), ilxPos, ilyPos, ilWidth, fgBoxGridH '1350, fgBoxGridH
    'Days
    ilxPos = tmXCtrls(XTIMEINDEX + 1).fBoxX + tmXCtrls(XTIMEINDEX + 1).fBoxW + 15: ilyPos = 375: ilWidth = 555
    gSetCtrl tmXCtrls(XDAYSINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Group
    ilxPos = tmXCtrls(XDAYSINDEX).fBoxX + tmXCtrls(XDAYSINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 555
    gSetCtrl tmXCtrls(XGROUPINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Demo
    ilxPos = tmXCtrls(XGROUPINDEX).fBoxX + tmXCtrls(XGROUPINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 795
    For illoop = XDEMOINDEX To XDEMOINDEX + 17
        If illoop <= XDEMOINDEX + 8 Then
            gSetCtrl tmXCtrls(illoop), ilxPos + (illoop - XDEMOINDEX) * (ilWidth + 15), ilyPos, ilWidth, fgBoxGridH
        Else
            gSetCtrl tmXCtrls(illoop), ilxPos + (illoop - (XDEMOINDEX + 9)) * (ilWidth + 15), ilyPos + 195, ilWidth, fgBoxGridH
        End If
    Next illoop
    
    '--------------------------------------------------------------------------
    'Time
    'Vehicle
    ilxPos = 30: ilyPos = 375: ilWidth = 5590
    gSetCtrl tmTCtrls(TVEHICLEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Time 1
    ilxPos = tmTCtrls(TVEHICLEINDEX).fBoxX + tmTCtrls(TVEHICLEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 495
    gSetCtrl tmTCtrls(TTIMEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH '1350, fgBoxGridH
    'Time 2
    ilxPos = tmTCtrls(TTIMEINDEX).fBoxX + tmTCtrls(TTIMEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 495
    gSetCtrl tmTCtrls(TTIMEINDEX + 1), ilxPos, ilyPos, ilWidth, fgBoxGridH  '1350, fgBoxGridH
    'Group or Days
    ilxPos = tmTCtrls(TTIMEINDEX + 1).fBoxX + tmTCtrls(TTIMEINDEX + 1).fBoxW + 15: ilyPos = 375: ilWidth = 1350
    gSetCtrl tmTCtrls(TDAYSINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Demo
    ilxPos = tmTCtrls(TDAYSINDEX).fBoxX + tmTCtrls(TDAYSINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 795
    For illoop = TDEMOINDEX To TDEMOINDEX + 17
        If illoop <= TDEMOINDEX + 8 Then
            gSetCtrl tmTCtrls(illoop), ilxPos + ((illoop - TDEMOINDEX) * (ilWidth + 15)), ilyPos, ilWidth, fgBoxGridH
        Else
            gSetCtrl tmTCtrls(illoop), ilxPos + ((illoop - (TDEMOINDEX + 9)) * (ilWidth + 15)), ilyPos + 195, ilWidth, fgBoxGridH
        End If
    Next illoop
    
    '--------------------------------------------------------------------------
    'Vehicle Type
    'Vehicle
    ilxPos = 30: ilyPos = 375: ilWidth = 4790
    gSetCtrl tmVCtrls(VVEHICLEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'ACT1CODE
    ilxPos = tmVCtrls(VVEHICLEINDEX).fBoxX + tmVCtrls(VVEHICLEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 1245
    gSetCtrl tmVCtrls(VACT1CODEINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'ACT1SETTING
    ilxPos = tmVCtrls(VACT1CODEINDEX).fBoxX + tmVCtrls(VACT1CODEINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 555
    gSetCtrl tmVCtrls(VACT1SETTINGINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Days
    ilxPos = tmVCtrls(VACT1SETTINGINDEX).fBoxX + tmVCtrls(VACT1SETTINGINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 1350
    gSetCtrl tmVCtrls(VDAYSINDEX), ilxPos, ilyPos, ilWidth, fgBoxGridH
    'Demo
    ilxPos = tmVCtrls(VDAYSINDEX).fBoxX + tmVCtrls(VDAYSINDEX).fBoxW + 15: ilyPos = 375: ilWidth = 795
    For illoop = VDEMOINDEX To VDEMOINDEX + 17
        If illoop <= VDEMOINDEX + 8 Then
            gSetCtrl tmVCtrls(illoop), ilxPos + (illoop - VDEMOINDEX) * (ilWidth + 15), ilyPos, ilWidth, fgBoxGridH
        Else
            gSetCtrl tmVCtrls(illoop), ilxPos + (illoop - (VDEMOINDEX + 9)) * (ilWidth + 15), ilyPos + 195, ilWidth, fgBoxGridH
        End If
    Next illoop
    
    '--------------------------------------------------------------------------
    'Plus Data
    'Demo Names
    gSetCtrl tmPCtrls(PDEMOINDEX), 30, 225, 870, fgBoxGridH
    'Audience
    gSetCtrl tmPCtrls(PAUDINDEX), 915, tmPCtrls(PDEMOINDEX).fBoxY, 960, fgBoxGridH
    gSetCtrl tmPCtrls(PPOPINDEX), 1890, tmPCtrls(PDEMOINDEX).fBoxY, 1065, fgBoxGridH
    'Calendar
    For illoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(illoop), 30 + 255 * (illoop - 1), 225, 240, fgBoxGridH
    Next illoop
    
    '--------------------------------------------------------------------------
    'Resize Spec Fields
    llMax = 0
    For illoop = imLBSCtrls To UBound(tmSCtrls) Step 1
        If illoop <= QUALPOPSRCEINDEX Then
            tmSCtrls(illoop).fBoxW = 2 * tmSCtrls(illoop).fBoxW
            Do While (tmSCtrls(illoop).fBoxW Mod 15) <> 0
                tmSCtrls(illoop).fBoxW = tmSCtrls(illoop).fBoxW + 1
            Loop
            If illoop = DATEINDEX Then
                llGap = tmSCtrls(illoop).fBoxW
                Do
                    If tmSCtrls(illoop).fBoxX < tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX + 15
                    ElseIf tmSCtrls(illoop).fBoxX > tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
            If illoop = QUALPOPSRCEINDEX Then
                Do
                    If tmSCtrls(illoop).fBoxX < tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX + 15
                    ElseIf tmSCtrls(illoop).fBoxX > tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        Else
            If illoop <= POPINDEX + 8 Then
                Do
                    If tmSCtrls(illoop).fBoxX < llGap + tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX + 15
                    ElseIf tmSCtrls(illoop).fBoxX > llGap + tmSCtrls(illoop - 1).fBoxX + tmSCtrls(illoop - 1).fBoxW + 15 Then
                        tmSCtrls(illoop).fBoxX = tmSCtrls(illoop).fBoxX - 15
                    Else
                        Exit Do
                    End If
                Loop
                llGap = 0
            Else
                tmSCtrls(illoop).fBoxX = tmSCtrls(illoop - 9).fBoxX
            End If
        End If
        If tmSCtrls(illoop).fBoxX + tmSCtrls(illoop).fBoxW + 15 > llMax Then
            llMax = tmSCtrls(illoop).fBoxX + tmSCtrls(illoop).fBoxW + 15
        End If
    Next illoop
    pbcSpec.Picture = LoadPicture("")
    llWidthChg = tmSCtrls(POPINDEX).fBoxX - ilDemoStart '(llMax - pbcSpec.Width)
    pbcSpec.Width = llMax
    plcSpec.Width = llMax + 2 * fgBevelX + 15
    
    '--------------------------------------------------------------------------
    'Determine Max Width of Demo grids
    For illoop = imLBDCtrls To UBound(tmDCtrls) Step 1
        If tmDCtrls(illoop).fBoxX + tmDCtrls(illoop).fBoxW + 15 > llMax Then
            llMax = tmDCtrls(illoop).fBoxX + tmDCtrls(illoop).fBoxW + 15
        End If
    Next illoop
    For illoop = imLBXCtrls To UBound(tmXCtrls) Step 1
        If tmXCtrls(illoop).fBoxX + tmXCtrls(illoop).fBoxW + 15 > llMax Then
            llMax = tmXCtrls(illoop).fBoxX + tmXCtrls(illoop).fBoxW + 15
        End If
    Next illoop
    For illoop = imLBTCtrls To UBound(tmTCtrls) Step 1
        If tmTCtrls(illoop).fBoxX + tmTCtrls(illoop).fBoxW + 15 > llMax Then
            llMax = tmTCtrls(illoop).fBoxX + tmTCtrls(illoop).fBoxW + 15
        End If
    Next illoop
    For illoop = imLBVCtrls To UBound(tmVCtrls) Step 1
        If tmVCtrls(illoop).fBoxX + tmVCtrls(illoop).fBoxW + 15 > llMax Then
            llMax = tmVCtrls(illoop).fBoxX + tmVCtrls(illoop).fBoxW + 15
        End If
    Next illoop
    
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    
    Me.Width = Me.Width + llMax - pbcDemo(0).Width
    plcDemo.Height = lacPlusTitle.Top - plcDemo.Top - ilSpaceBetweenButtons
    llHeight = ((plcDemo.Height - 2 * fgBevelY) \ (tmDCtrls(1).fBoxH + 15))
    If llHeight Mod 10 = 1 Then llHeight = llHeight - 1
    llLargeChg = llHeight
    llHeight = (tmDCtrls(1).fBoxH + 15) * llHeight
    For illoop = 0 To 2 Step 1
        pbcDemo(illoop).Picture = LoadPicture("")
        pbcDemo(illoop).Width = llMax
        lacFrame(illoop).Width = llMax - 15
        pbcDemo(illoop).Height = llHeight
    Next illoop
    plcDemo.Width = llMax + vbcDemo.Width + 2 * fgBevelX + 15
    plcDemo.Height = pbcDemo(0).Height + 2 * fgBevelY
    vbcDemo.Move pbcDemo(0).Left + pbcDemo(0).Width, pbcDemo(0).Top, vbcDemo.Width, pbcDemo(0).Height
    cmcSetDefault.Left = Research.Width - 2 * imcTrash.Width - cmcSetDefault.Width
    cmcSetDefault.Top = plcPlus.Top + plcPlus.Height - cmcSetDefault.Height
    cmcSocEco.Move cmcSetDefault.Left - cmcSocEco.Width - ilSpaceBetweenButtons, cmcSetDefault.Top
    cmcUndo.Move cmcSocEco.Left - cmcUndo.Width - ilSpaceBetweenButtons, cmcSetDefault.Top
    cmcErase.Move cmcUndo.Left - cmcErase.Width - ilSpaceBetweenButtons, cmcSetDefault.Top
    
    cmcDone.Move cmcErase.Left, cmcErase.Top - cmcDone.Height - ilSpaceBetweenButtons
    cmcCancel.Move cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons, cmcDone.Top
    cmcSave.Move cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons, cmcDone.Top
    cmcAdjust.Move cmcSave.Left + cmcSave.Width + ilSpaceBetweenButtons, cmcDone.Top
    
    imcTrash.Top = cmcDone.Top '+ cmcDone.Height - imcTrash.Height
    imcTrash.Left = Research.Width - (3 * imcTrash.Width) / 2
    
    cbcSelect.Left = plcSpec.Left + plcSpec.Width - cbcSelect.Width
    cbcDemo.Top = plcDemo.Top + plcDemo.Height + 60
    cbcDemo.Left = plcSpec.Left + plcSpec.Width - cbcDemo.Width
    
    cmcDuplicate.Left = Research.Width / 2 - cmcDuplicate.Width / 2 + cmcDuplicate.Width
    cmcDuplicate.Top = cbcDemo.Top
    cmcBaseDuplicate.Left = Research.Width / 2 - cmcBaseDuplicate.Width / 2 - cmcBaseDuplicate.Width
    cmcBaseDuplicate.Top = cbcDemo.Top
    
    vbcDemo.LargeChange = llLargeChg \ 2 - 2
    lmVbcDemoLargeChg = llLargeChg
    
    imcKey.Left = plcScreen.Left + plcScreen.Width
    imcKey.Top = plcScreen.Top
    lacStart.Left = imcKey.Left + imcKey.Width + 120
    edcStart.Left = lacStart.Left + lacStart.Width + 120
    lacEnd.Left = edcStart.Left + edcStart.Width + 120
    edcEnd.Left = lacEnd.Left + lacEnd.Width + 120
    cbcVehicle.Left = edcEnd.Left + edcEnd.Width + 120
    cbcVehicle.AddItem "[All Vehicles]"
    cbcVehicle.ListIndex = 0
    llWidth = (Me.Width - cbcVehicle.Left - cmcGetBook.Width) / 2
    cbcVehicle.Width = (3 * llWidth) / 4
    cmcGetBook.Left = cbcVehicle.Left + cbcVehicle.Width + 120
    cbcSelect.Width = plcSpec.Left + plcSpec.Width - (cmcGetBook.Left + cmcGetBook.Width + 120)
    cbcSelect.Left = cmcGetBook.Left + cmcGetBook.Width + 120
    
    tmDP = tmDCtrls(DDAYPARTINDEX)
    tmGroup = tmDCtrls(DGROUPINDEX)
    
    CSI_ComboBoxMS1.Left = cbcSelect.Left
    CSI_ComboBoxMS1.Top = cbcSelect.Top
    CSI_ComboBoxMS1.Height = cbcSelect.Height
    CSI_ComboBoxMS1.Width = cbcSelect.Width
    CSI_ComboBoxMS1.SetDropDownWidth cbcSelect.Width
    cbcSelect.Visible = False
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewDpf                     *
'*                                                     *
'*             Created:8/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Demo record         *
'*                                                     *
'*******************************************************
Private Sub mInitNewDpf()
    Dim llUpper As Long
    llUpper = UBound(tgDpfRec)
    tgDpfRec(llUpper).sKey = ""
    tgDpfRec(llUpper).iStatus = 0
    tgDpfRec(llUpper).lDpfCode = 0
    tgDpfRec(llUpper).lDrfCode = 0  'Set when row added in pbcPlusTab
    tgDpfRec(llUpper).sDemo = ""
    tgDpfRec(llUpper).sPop = ""
    tgDpfRec(llUpper).lIndex = 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewDef                     *
'*                                                     *
'*             Created:8/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Demo record         *
'*                                                     *
'*******************************************************
Private Sub mInitNewDef()
    Dim llUpper As Long
    llUpper = UBound(tgDefRec)
    tgDefRec(llUpper).sKey = ""
    tgDefRec(llUpper).iStatus = 0
    tgDefRec(llUpper).lDefCode = 0
    tgDefRec(llUpper).sStartDate = ""
    tgDefRec(llUpper).sPop = ""
    tgDefRec(llUpper).sEstPct = ""
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewDrf                    *
'*                                                     *
'*             Created:8/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Demo record         *
'*                                                     *
'*******************************************************
Private Sub mInitNewDrf(ilInitSaveShow As Integer, llUpper As Long)
    Dim llLoop As Long
    If ilInitSaveShow Then
        For llLoop = imLBSaveShow To UBound(tmSaveShow(imLBSaveShow).sSave) Step 1
            tmSaveShow(llUpper).sSave(llLoop) = ""
        Next llLoop
        For llLoop = imLBSaveShow To UBound(tmSaveShow(imLBSaveShow).sShow) Step 1
            tmSaveShow(llUpper).sShow(llLoop) = ""
        Next llLoop
    End If
    tgDrfRec(llUpper).iStatus = 0
    tgDrfRec(llUpper).lRecPos = 0
    tgDrfRec(llUpper).lIndex = 0
    tgDrfRec(llUpper).iModel = False
    tgDrfRec(llUpper).lLink = -1
    tgDrfRec(llUpper).tDrf.lCode = 0
    If imCustomIndex <= 0 Then
        tgDrfRec(llUpper).tDrf.sDataType = "A"    'Standard demos
    Else
        tgDrfRec(llUpper).tDrf.sDataType = smUniqueGroupDataTypes(0)    'Custom Demos- Set as part of save
    End If
    If rbcDataType(0).Value Then 'Daypart
        tgDrfRec(llUpper).tDrf.sInfoType = "D"
        tgDrfRec(llUpper).tDrf.iRdfCode = -1    'A non zero value which indicates type of daypart
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        tgDrfRec(llUpper).tDrf.sInfoType = "D"
        tgDrfRec(llUpper).tDrf.iRdfCode = 0
    ElseIf rbcDataType(2).Value Then 'Time
        tgDrfRec(llUpper).tDrf.sInfoType = "T"
        tgDrfRec(llUpper).tDrf.iRdfCode = 0
    Else 'Vehicle
        tgDrfRec(llUpper).tDrf.sInfoType = "V"
        tgDrfRec(llUpper).tDrf.iRdfCode = 0
    End If
    If imCustomIndex > 0 Then
        For llLoop = 1 To UBound(smUniqueGroupDataTypes) - 1 Step 1
            tgLinkDrfRec(UBound(tgLinkDrfRec)) = tgDrfRec(llUpper)
            tgLinkDrfRec(UBound(tgLinkDrfRec)).tDrf.lCode = 0
            tgLinkDrfRec(UBound(tgLinkDrfRec)).tDrf.sDataType = smUniqueGroupDataTypes(llLoop)    'Custom Demos- Set as part of save
            tgLinkDrfRec(UBound(tgLinkDrfRec)).lIndex = -1
            tgLinkDrfRec(UBound(tgLinkDrfRec)).lLink = -1
            tgLinkDrfRec(UBound(tgLinkDrfRec)).iCustInfoIndex = -1
            If tgDrfRec(llUpper).lLink = -1 Then
                tgDrfRec(llUpper).lLink = UBound(tgLinkDrfRec)
            Else
                tgLinkDrfRec(UBound(tgLinkDrfRec)).lLink = tgDrfRec(llUpper).lLink
                tgDrfRec(llUpper).lLink = UBound(tgLinkDrfRec)
            End If
            ReDim Preserve tgLinkDrfRec(0 To UBound(tgLinkDrfRec) + 1) As DRFREC
        Next llLoop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
'
'   mInitShow
'   Where:
'
    Dim slStr As String
    Dim ilBoxNo As Integer
    Dim llRowNo As Long
    
    For llRowNo = imLBDrf To UBound(tgDrfRec) - 1 Step 1
        If rbcDataType(0).Value Then 'Daypart
            For ilBoxNo = imLBDCtrls To UBound(tmDCtrls) Step 1
                Select Case ilBoxNo
                    Case DVEHICLEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(DVEHICLEINDEX))
                        gSetShow pbcDemo(0), slStr, tmDCtrls(DVEHICLEINDEX)
                        tmSaveShow(llRowNo).sShow(1) = tmDCtrls(DVEHICLEINDEX).sShow
                    
                    Case DACT1CODEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(DACT1CODEINDEX))
                        gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1CODEINDEX)
                        tmSaveShow(llRowNo).sShow(DACT1CODEINDEX) = tmDCtrls(DACT1CODEINDEX).sShow
                        
                    Case DACT1SETTINGINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX))
                        gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1SETTINGINDEX)
                        tmSaveShow(llRowNo).sShow(DACT1SETTINGINDEX) = tmDCtrls(DACT1SETTINGINDEX).sShow
                        
                    Case DDAYPARTINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(DDAYPARTINDEX))
                        gSetShow pbcDemo(0), slStr, tmDCtrls(DDAYPARTINDEX)
                        tmSaveShow(llRowNo).sShow(DDAYPARTINDEX) = tmDCtrls(DDAYPARTINDEX).sShow
                        
                    Case DGROUPINDEX
                        If smSource <> "I" Then 'Standard Airtime mode
                            slStr = Trim$(tmSaveShow(llRowNo).sSave(DAIRTIMEGRPNOINDEX))
                            gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                            tmSaveShow(llRowNo).sShow(DGROUPINDEX) = tmDCtrls(DGROUPINDEX).sShow
                        Else 'Podcast Impression mode
                            slStr = Trim$(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX))
                            gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                            tmSaveShow(llRowNo).sShow(DIMPRESSIONSINDEX) = tmDCtrls(ilBoxNo).sShow
                        End If
                        
                    Case DDEMOINDEX To DDEMOINDEX + 17
                        If smSource <> "I" Then 'Standard Airtime mode
                            slStr = Trim$(tmSaveShow(llRowNo).sSave(DDEMOINDEX + ilBoxNo - DDEMOINDEX))
                            gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                            tmSaveShow(llRowNo).sShow(DDEMOINDEX + ilBoxNo - DDEMOINDEX) = tmDCtrls(ilBoxNo).sShow
                        End If
                End Select
            Next ilBoxNo
            
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            For ilBoxNo = imLBXCtrls To UBound(tmXCtrls) Step 1
                Select Case ilBoxNo
                    Case XVEHICLEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XVEHICLEINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(XVEHICLEINDEX) = tmXCtrls(ilBoxNo).sShow
                    
                    Case XACT1CODEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XACT1CODEINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(XACT1CODEINDEX)
                        tmSaveShow(llRowNo).sShow(DACT1CODEINDEX) = tmXCtrls(XACT1CODEINDEX).sShow
                        
                    Case XACT1SETTINGINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(XACT1SETTINGINDEX)
                        tmSaveShow(llRowNo).sShow(XACT1SETTINGINDEX) = tmXCtrls(XACT1SETTINGINDEX).sShow
                    
                    Case XTIMEINDEX To XTIMEINDEX + 1
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XTIMEINDEX + ilBoxNo - XTIMEINDEX))
                        If gValidTime(slStr) Then
                            If slStr <> "" Then
                                slStr = gFormatTime(slStr, "A", "1")
                            End If
                            gSetShow pbcDemo(0), slStr, tmXCtrls(ilBoxNo)
                            tmSaveShow(llRowNo).sShow(XTIMEINDEX + ilBoxNo - XTIMEINDEX) = tmXCtrls(ilBoxNo).sShow
                        Else
                            tmSaveShow(llRowNo).sShow(XTIMEINDEX + ilBoxNo - XTIMEINDEX) = ""
                        End If
                        
                    Case XDAYSINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XDAYSINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(XDAYSINDEX) = tmXCtrls(ilBoxNo).sShow
                        
                    Case XGROUPINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XGROUPNINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(XGROUPINDEX) = tmXCtrls(ilBoxNo).sShow
                        
                    Case XDEMOINDEX To XDEMOINDEX + 17
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(XDEMOINDEX + ilBoxNo - XDEMOINDEX))
                        gSetShow pbcDemo(0), slStr, tmXCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(XDEMOINDEX + ilBoxNo - XDEMOINDEX) = tmXCtrls(ilBoxNo).sShow
                End Select
            Next ilBoxNo
            
        ElseIf rbcDataType(2).Value Then 'Time
            For ilBoxNo = imLBTCtrls To UBound(tmTCtrls) Step 1
                Select Case ilBoxNo
                    Case TVEHICLEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(TVEHICLEINDEX))
                        gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(TVEHICLEINDEX) = tmTCtrls(ilBoxNo).sShow
                        
                    Case TTIMEINDEX To TTIMEINDEX + 1
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(TTIMEINDEX + ilBoxNo - TTIMEINDEX))
                        If gValidTime(slStr) Then
                            If slStr <> "" Then
                                slStr = gFormatTime(slStr, "A", "1")
                            End If
                            gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                            tmSaveShow(llRowNo).sShow(TTIMEINDEX + ilBoxNo - TTIMEINDEX) = tmTCtrls(ilBoxNo).sShow
                        Else
                            tmSaveShow(llRowNo).sShow(TTIMEINDEX + ilBoxNo - TTIMEINDEX) = ""
                        End If
                        
                    Case TDAYSINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(TDAYSINDEX))
                        gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(TDAYSINDEX) = tmTCtrls(ilBoxNo).sShow
                        
                    Case TDEMOINDEX To TDEMOINDEX + 17
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(TDEMOINDEX + ilBoxNo - TDEMOINDEX))
                        gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(TDEMOINDEX + ilBoxNo - TDEMOINDEX) = tmTCtrls(ilBoxNo).sShow
                End Select
            Next ilBoxNo
            
        Else 'Vehicle
            For ilBoxNo = imLBVCtrls To UBound(tmVCtrls) Step 1
                Select Case ilBoxNo
                    Case VVEHICLEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(VVEHICLEINDEX))
                        gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(VVEHICLEINDEX) = tmVCtrls(ilBoxNo).sShow
                    
                    Case VACT1CODEINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(VACT1CODEINDEX))
                        gSetShow pbcDemo(1), slStr, tmVCtrls(VACT1CODEINDEX)
                        tmSaveShow(llRowNo).sShow(VACT1CODEINDEX) = tmVCtrls(VACT1CODEINDEX).sShow
                        
                    Case VACT1SETTINGINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX))
                        gSetShow pbcDemo(1), slStr, tmVCtrls(VACT1SETTINGINDEX)
                        tmSaveShow(llRowNo).sShow(VACT1SETTINGINDEX) = tmVCtrls(VACT1SETTINGINDEX).sShow
                        
                    Case VDAYSINDEX
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(VDAYSINDEX))
                        gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(VDAYSINDEX) = tmVCtrls(VDAYSINDEX).sShow
                        
                    Case VDEMOINDEX To VDEMOINDEX + 17
                        slStr = Trim$(tmSaveShow(llRowNo).sSave(VDEMOINDEX + ilBoxNo - VDEMOINDEX))
                        gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                        tmSaveShow(llRowNo).sShow(VDEMOINDEX + ilBoxNo - VDEMOINDEX) = tmVCtrls(ilBoxNo).sShow
                End Select
            Next ilBoxNo
        End If
    Next llRowNo
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitSShow                      *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitSShow()
'
'   mInitSShow
'   Where:
'
    Dim ilBoxNo As Integer
    Dim slStr As String
    For ilBoxNo = imLBSCtrls To UBound(tmSCtrls) Step 1
        Select Case ilBoxNo
            Case NAMEINDEX
                slStr = smSSave(NAMEINDEX)
                gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            Case DATEINDEX
                slStr = smSSave(DATEINDEX)
                If gValidDate(slStr) Then
                    slStr = gFormatDate(slStr)
                    gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
                Else
                    tmVCtrls(ilBoxNo).sShow = ""
                End If
            Case POPSRCEINDEX
                slStr = smSSave(POPSRCDESCINDEX)
                gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            Case QUALPOPSRCEINDEX
                slStr = smSSave(QUALSRCDESCINDEX)
                gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            Case POPINDEX To POPINDEX + 17
                slStr = smSSave(POPINDEX + ilBoxNo - POPINDEX)
                gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim ilIndexOffset As Integer 'used to replace hard-coded array position offsets
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRowNo As Long
    Dim slDays As String
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim llTime As Long
    Dim ilRdfCode As Integer
    Dim ilPop As Integer
    Dim ilLoop1 As Integer
    Dim llLink As Long
    Dim ilDay As Integer

    '------------------------
    'Research Header
    '------------------------
    tgDnf.sBookName = smSSave(NAMEINDEX)
    gPackDate smSSave(DATEINDEX), tgDnf.iBookDate(0), tgDnf.iBookDate(1)
    gFindMatch Trim$(smSSave(POPSRCDESCINDEX)), 1, lbcPopSrce
    If gLastFound(lbcPopSrce) >= 1 Then
        tgDnf.iPopDnfCode = lbcPopSrce.ItemData(gLastFound(lbcPopSrce))
        If tgDnf.iCode = tgDnf.iPopDnfCode Then
            tgDnf.iPopDnfCode = 0
        End If
    Else
        tgDnf.iPopDnfCode = 0
    End If
    gFindMatch Trim$(smSSave(QUALSRCDESCINDEX)), 1, lbcPopSrce
    If gLastFound(lbcPopSrce) >= 1 Then
        tgDnf.iQualPopDnfCode = lbcPopSrce.ItemData(gLastFound(lbcPopSrce))
        If tgDnf.iCode = tgDnf.iQualPopDnfCode Then
            tgDnf.iQualPopDnfCode = 0
        End If
    Else
        tgDnf.iQualPopDnfCode = 0
    End If
    If imEstByLOrU = 1 Then
        tgDnf.sEstListenerOrUSA = "U"
    Else
        tgDnf.sEstListenerOrUSA = "L"
    End If
    tgDnf.sForm = smDataForm
    If smSource = "I" Then 'Podcast Impression mode
        tgDnf.sSource = "I"
    End If
    ilIndex = 1
    ilIndexOffset = POPINDEX - 1
    If imCustomIndex <= 0 Then
        For illoop = 1 To 18 Step 1
            If smSSave(ilIndexOffset + illoop) <> "" Then
                If tgSpf.sSAudData = "H" Then
                    tgSDrfPop.lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + illoop), 1)
                ElseIf tgSpf.sSAudData = "N" Then
                    tgSDrfPop.lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + illoop), 2)
                ElseIf tgSpf.sSAudData = "U" Then
                    tgSDrfPop.lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + illoop), 3)
                Else
                    tgSDrfPop.lDemo(ilIndex - 1) = Val(smSSave(ilIndexOffset + illoop))
                End If
            Else
                tgSDrfPop.lDemo(ilIndex - 1) = 0
            End If
            If (smDataForm <> "8") And ((illoop = 9) Or (illoop = 18)) Then
                If illoop = 18 Then
                    tgSDrfPop.lDemo(16) = 0
                    tgSDrfPop.lDemo(17) = 0
                    Exit For
                End If
            Else
                ilIndex = ilIndex + 1
            End If
        Next illoop
    Else
        For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
            If illoop < UBound(tmCustInfo) Then
                For ilPop = imLBDrf To UBound(tgCDrfPop) - 1 Step 1
                    If tgCDrfPop(ilPop).sDataType = tmCustInfo(illoop).sDataType Then
                        ilIndex = tmCustInfo(illoop).iDemoIndex
                        ilLoop1 = illoop - (imCustomIndex - 1) + 1
                        If smSSave(ilIndexOffset + ilLoop1) <> "" Then
                            If tgSpf.sSAudData = "H" Then
                                tgCDrfPop(ilPop).lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + ilLoop1), 1)
                            ElseIf tgSpf.sSAudData = "N" Then
                                tgCDrfPop(ilPop).lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + ilLoop1), 2)
                            ElseIf tgSpf.sSAudData = "U" Then
                                tgCDrfPop(ilPop).lDemo(ilIndex - 1) = gStrDecToLong(smSSave(ilIndexOffset + ilLoop1), 3)
                            Else
                                tgCDrfPop(ilPop).lDemo(ilIndex - 1) = Val(smSSave(ilIndexOffset + ilLoop1))
                            End If
                        Else
                            tgCDrfPop(ilPop).lDemo(ilIndex - 1) = 0
                        End If
                    End If
                Next ilPop
            End If
        Next illoop
    End If
    
    '------------------------
    'Research Data Rows
    '------------------------
    For llRowNo = imLBDrf To UBound(tgDrfRec) - 1 Step 1
        '------------------------
        'Vehicle
        gFindMatch Trim$(tmSaveShow(llRowNo).sSave(1)), 0, lbcVehicle
        If gLastFound(lbcVehicle) >= 0 Then
            slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgDrfRec(llRowNo).tDrf.iVefCode = Val(slCode)
        End If
        
        If (imDataType = 0 Or imDataType = 1 Or imDataType = 3) And smSource <> "I" Then  'Daypart, ExtraDaypart, and Vehicle (Not Time); AND AirTime only, not podcast Impressions
            '------------------------
            'ACT1 Code
            tgDrfRec(llRowNo).tDrf.sACTLineupCode = Trim$(tmSaveShow(llRowNo).sSave(DACT1CODEINDEX))
            '------------------------
            'ACT1 Setting
            slCode = tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX)
            If InStr(1, slCode, "T") > 0 Then
                tgDrfRec(llRowNo).tDrf.sACT1StoredTime = "T"
            Else
                tgDrfRec(llRowNo).tDrf.sACT1StoredTime = ""
            End If
            If InStr(1, slCode, "S") > 0 Then
                tgDrfRec(llRowNo).tDrf.sACT1StoredSpots = "S"
            Else
                tgDrfRec(llRowNo).tDrf.sACT1StoredSpots = ""
            End If
            If InStr(1, slCode, "C") > 0 Then
                tgDrfRec(llRowNo).tDrf.sACT1StoreClearPct = "C"
            Else
                tgDrfRec(llRowNo).tDrf.sACT1StoreClearPct = ""
            End If
            If InStr(1, slCode, "F") > 0 Then
                tgDrfRec(llRowNo).tDrf.sACT1DaypartFilter = "F"
            Else
                tgDrfRec(llRowNo).tDrf.sACT1DaypartFilter = ""
            End If
        End If
        
        '------------------------
        'Daypart/Times
        Select Case imDataType
            Case 0 'Daypart
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode <> 0) Then
                    tgDrfRec(llRowNo).tDrf.sInfoType = "D"
                    If tgDrfRec(llRowNo).iStatus <> 1 Then  'Retain Start/End date imported
                        tgDrfRec(llRowNo).tDrf.iStartTime(0) = 1
                        tgDrfRec(llRowNo).tDrf.iStartTime(1) = 0
                        tgDrfRec(llRowNo).tDrf.iEndTime(0) = 1
                        tgDrfRec(llRowNo).tDrf.iEndTime(1) = 0
                    End If
                    'Daypart
                    ilRdfCode = tgDrfRec(llRowNo).tDrf.iRdfCode
                    tgDrfRec(llRowNo).tDrf.iRdfCode = -1
                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If (tgMRif(llRif).iVefCode = tgDrfRec(llRowNo).tDrf.iVefCode) Then
                            For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                If (tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode) And (StrComp(Trim$(tmSaveShow(llRowNo).sSave(DDAYPARTINDEX)), Trim$(tgMRdf(ilRdf).sName), 1) = 0) Then
                                    tgDrfRec(llRowNo).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                    Exit For
                                End If
                            Next ilRdf
                            If tgDrfRec(llRowNo).tDrf.iRdfCode > 0 Then
                                Exit For
                            End If
                        End If
                    Next llRif
                    If tgDrfRec(llRowNo).tDrf.iRdfCode < 0 Then
                        For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                            If ilRdfCode = tgMRdf(ilRdf).iCode Then
                                tgDrfRec(llRowNo).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                Exit For
                            End If
                        Next ilRdf
                    End If
                    If tgDrfRec(llRowNo).tDrf.iRdfCode < 0 Then
                        For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                            If (StrComp(Trim$(tmSaveShow(llRowNo).sSave(DDAYPARTINDEX)), Trim$(tgMRdf(ilRdf).sName), 1) = 0) Then
                                tgDrfRec(llRowNo).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                Exit For
                            End If
                        Next ilRdf
                    End If
                End If

            Case 1 'Extra Daypart
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0) Then
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    'Start Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(XTIMEINDEX)), tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1)
                    'End Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(XTIMEINDEX + 1)), tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1)
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "T" Then
                    tgDrfRec(llRowNo).tDrf.iRdfCode = 0
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    'Start Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(XTIMEINDEX)), tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1)
                    'End Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(XTIMEINDEX + 1)), tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1)
                    gUnpackTimeLong tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), False, llTime
                    tgDrfRec(llRowNo).tDrf.iQHIndex = llTime \ 900 + 1
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "V" Then
                    tgDrfRec(llRowNo).tDrf.iRdfCode = 0
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    tgDrfRec(llRowNo).tDrf.iStartTime(0) = 1
                    tgDrfRec(llRowNo).tDrf.iStartTime(1) = 0
                    tgDrfRec(llRowNo).tDrf.iEndTime(0) = 1
                    tgDrfRec(llRowNo).tDrf.iEndTime(1) = 0
                    tgDrfRec(llRowNo).tDrf.sProgCode = ""
                End If
            
            Case 2 'Time
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0) Then
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    'Start Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(TTIMEINDEX)), tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1)
                    'End Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(TTIMEINDEX + 1)), tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1)
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "T" Then
                    tgDrfRec(llRowNo).tDrf.iRdfCode = 0
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    'Start Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(TTIMEINDEX)), tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1)
                    'End Time
                    gPackTime Trim$(tmSaveShow(llRowNo).sSave(TTIMEINDEX + 1)), tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1)
                    gUnpackTimeLong tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), False, llTime
                    tgDrfRec(llRowNo).tDrf.iQHIndex = llTime \ 900 + 1
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "V" Then
                    tgDrfRec(llRowNo).tDrf.iRdfCode = 0
                    tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                    tgDrfRec(llRowNo).tDrf.iStartTime(0) = 1
                    tgDrfRec(llRowNo).tDrf.iStartTime(1) = 0
                    tgDrfRec(llRowNo).tDrf.iEndTime(0) = 1
                    tgDrfRec(llRowNo).tDrf.iEndTime(1) = 0
                    tgDrfRec(llRowNo).tDrf.sProgCode = ""
                End If
            Case 3 'Vehicle
        End Select
        
        '------------------------
        'Group #
        Select Case imDataType
            Case 0 'Daypart
                'Group #
                tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                If Trim$(tmSaveShow(llRowNo).sSave(DAIRTIMEGRPNOINDEX)) <> "" Then
                    For illoop = imLBMnf To UBound(tgMnfSocEco) - 1 Step 1
                        If Trim$(tmSaveShow(llRowNo).sSave(DAIRTIMEGRPNOINDEX)) = Trim$(tgMnfSocEco(illoop).sUnitType) Then
                            tgDrfRec(llRowNo).tDrf.iMnfSocEco = tgMnfSocEco(illoop).iCode
                            Exit For
                        End If
                    Next illoop
                End If
            Case 1 'Extra Daypart
                'Group #
                tgDrfRec(llRowNo).tDrf.iMnfSocEco = 0
                If Trim$(tmSaveShow(llRowNo).sSave(XGROUPNINDEX)) <> "" Then
                    For illoop = imLBMnf To UBound(tgMnfSocEco) - 1 Step 1
                        If Trim$(tmSaveShow(llRowNo).sSave(XGROUPNINDEX)) = Trim$(tgMnfSocEco(illoop).sUnitType) Then
                            tgDrfRec(llRowNo).tDrf.iMnfSocEco = tgMnfSocEco(illoop).iCode
                            Exit For
                        End If
                    Next illoop
                End If
            Case 2 'Time
            Case 3 'Vehicle
        End Select
        
        '------------------------
        'Days
        If imDataType = 1 Or imDataType = 2 Or imDataType = 3 Then 'ExtraDaypart,Time,Vehicle (not Daypart)
            If ((tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0)) Or (tgDrfRec(llRowNo).tDrf.sInfoType = "T") Or (tgDrfRec(llRowNo).tDrf.sInfoType = "V") Then
                'Days
                slDays = "YYYYYYY"
                'JW - 9/23/21 - Fix Research List screen: display Act 1 lineup codes and settings - Issue 1: the “Days” value stays as "Mo-Su" per Jason Email [9/20 21 3:46 PM]
                Select Case imDataType
                    Case 1 'Extra Daypart
                        gFindMatch Trim$(tmSaveShow(llRowNo).sSave(XDAYSINDEX)), 0, lbcDays
                    Case 2 'Time
                        gFindMatch Trim$(tmSaveShow(llRowNo).sSave(TDAYSINDEX)), 0, lbcDays
                    Case 3 'Vehicle
                        gFindMatch Trim$(tmSaveShow(llRowNo).sSave(VDAYSINDEX)), 0, lbcDays
                End Select
                If gLastFound(lbcDays) = 0 Then
                    slDays = "YYYYYNN"
                ElseIf gLastFound(lbcDays) = 1 Then
                    slDays = "NNNNNYN"
                ElseIf gLastFound(lbcDays) = 2 Then
                    slDays = "NNNNNNY"
                ElseIf gLastFound(lbcDays) = 3 Then
                    slDays = "YYYYYYN"
                ElseIf gLastFound(lbcDays) = 4 Then
                    slDays = "YYYYYYY"
                ElseIf gLastFound(lbcDays) = 5 Then
                    slDays = "NNNNNYY"
                ElseIf gLastFound(lbcDays) = 6 Then
                    slDays = "NYYYYYY"
                ElseIf gLastFound(lbcDays) = 7 Then
                    slDays = "NYYYYNN"
                ElseIf gLastFound(lbcDays) = 8 Then
                    slDays = "NNYYYYY"
                ElseIf gLastFound(lbcDays) = 9 Then
                    slDays = "YNNNNNN"
                ElseIf gLastFound(lbcDays) = 10 Then
                    slDays = "NYNNNNN"
                ElseIf gLastFound(lbcDays) = 11 Then
                    slDays = "NNYNNNN"
                ElseIf gLastFound(lbcDays) = 12 Then
                    slDays = "NNNYNNN"
                ElseIf gLastFound(lbcDays) = 13 Then
                    slDays = "NNNNYNN"
                ElseIf gLastFound(lbcDays) = 14 Then
                    slDays = "YYYYNNN"
                ElseIf gLastFound(lbcDays) = 15 Then
                    slDays = "YYYNNNN"
                ElseIf gLastFound(lbcDays) = 16 Then
                    slDays = "YYNNNNN"
                ElseIf gLastFound(lbcDays) = 17 Then
                    slDays = "NYYYYYN"
                ElseIf gLastFound(lbcDays) = 18 Then
                    slDays = "NYYYNNN"
                ElseIf gLastFound(lbcDays) = 19 Then
                    slDays = "NYYNNNN"
                ElseIf gLastFound(lbcDays) = 20 Then
                    slDays = "NNYYYYN"
                ElseIf gLastFound(lbcDays) = 21 Then
                    slDays = "NNYYYNN"
                ElseIf gLastFound(lbcDays) = 22 Then
                    slDays = "NNYYNNN"
                ElseIf gLastFound(lbcDays) = 23 Then
                    slDays = "NNNYYYY"
                ElseIf gLastFound(lbcDays) = 24 Then
                    slDays = "NNNYYYN"
                ElseIf gLastFound(lbcDays) = 25 Then
                    slDays = "NNNYYNN"
                ElseIf gLastFound(lbcDays) = 26 Then
                    slDays = "NNNNYYY"
                ElseIf gLastFound(lbcDays) = 27 Then
                    slDays = "NNNNYYN"
                End If
                For illoop = 0 To 6 Step 1
                    tgDrfRec(llRowNo).tDrf.sDay(illoop) = Left$(slDays, 1)
                    slDays = right$(slDays, Len(slDays) - 1)
                Next illoop
            End If
        End If
        
        '------------------------
        'Demos
        Select Case imDataType
            Case 0 'Daypart
                ilIndexOffset = DDEMOINDEX - 1
            Case 1 'Extra Daypart
                ilIndexOffset = XDEMOINDEX - 1
            Case 2 'Time
                ilIndexOffset = TDEMOINDEX - 1
            Case 3 'Vehicle
                ilIndexOffset = VDEMOINDEX - 1
        End Select
        ilIndex = 1
        If imCustomIndex <= 0 Then
            For illoop = 1 To 18 Step 1
                If Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + illoop)) <> "" Then
                    If tgSpf.sSAudData = "H" Then
                        tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + illoop), 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + illoop), 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + illoop), 3)
                    Else
                        tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = Val(Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + illoop)))
                    End If
                Else
                    tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = 0
                End If
                If (smDataForm <> "8") And ((illoop = 9) Or (illoop = 18)) Then
                    If illoop = 18 Then
                        tgSDrfPop.lDemo(16) = 0
                        tgSDrfPop.lDemo(17) = 0
                        Exit For
                    End If
                Else
                    ilIndex = ilIndex + 1
                End If
            Next illoop
        Else
            For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
                ilIndex = -1
                If illoop < UBound(tmCustInfo) Then
                    ilLoop1 = illoop - (imCustomIndex - 1) + 1
                    If tgDrfRec(llRowNo).tDrf.sDataType = tmCustInfo(illoop).sDataType Then
                        ilIndex = tmCustInfo(illoop).iDemoIndex
                        If Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1)) <> "" Then
                            If tgSpf.sSAudData = "H" Then
                                tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 1)
                            ElseIf tgSpf.sSAudData = "N" Then
                                tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 2)
                            ElseIf tgSpf.sSAudData = "U" Then
                                tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 3)
                            Else
                                tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = Val(Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1)))
                            End If
                        Else
                            tgDrfRec(llRowNo).tDrf.lDemo(ilIndex - 1) = 0
                        End If
                    Else
                        llLink = tgDrfRec(llRowNo).lLink
                        Do While llLink <> -1
                            If tgLinkDrfRec(llLink).tDrf.sDataType = tmCustInfo(illoop).sDataType Then
                                ilIndex = tmCustInfo(illoop).iDemoIndex
                                Exit Do
                            End If
                            llLink = tgLinkDrfRec(llLink).lLink
                        Loop
                        If llLink <> -1 Then
                            tgLinkDrfRec(llLink).tDrf.iVefCode = tgDrfRec(llRowNo).tDrf.iVefCode
                            tgLinkDrfRec(llLink).tDrf.sInfoType = tgDrfRec(llRowNo).tDrf.sInfoType
                            tgLinkDrfRec(llLink).tDrf.iRdfCode = tgDrfRec(llRowNo).tDrf.iRdfCode
                            tgLinkDrfRec(llLink).tDrf.iMnfSocEco = tgDrfRec(llRowNo).tDrf.iMnfSocEco
                            tgLinkDrfRec(llLink).tDrf.iStartTime(0) = tgDrfRec(llRowNo).tDrf.iStartTime(0)
                            tgLinkDrfRec(llLink).tDrf.iStartTime(1) = tgDrfRec(llRowNo).tDrf.iStartTime(1)
                            tgLinkDrfRec(llLink).tDrf.iEndTime(0) = tgDrfRec(llRowNo).tDrf.iEndTime(0)
                            tgLinkDrfRec(llLink).tDrf.iEndTime(1) = tgDrfRec(llRowNo).tDrf.iEndTime(1)
                            tgLinkDrfRec(llLink).tDrf.sProgCode = tgDrfRec(llRowNo).tDrf.sProgCode
                            For ilDay = 0 To 6 Step 1
                                tgLinkDrfRec(llLink).tDrf.sDay(ilDay) = tgDrfRec(llRowNo).tDrf.sDay(ilDay)
                            Next ilDay
                            If Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1)) <> "" Then
                                If tgSpf.sSAudData = "H" Then
                                    tgLinkDrfRec(llLink).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 1)
                                ElseIf tgSpf.sSAudData = "N" Then
                                    tgLinkDrfRec(llLink).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 2)
                                ElseIf tgSpf.sSAudData = "U" Then
                                    tgLinkDrfRec(llLink).tDrf.lDemo(ilIndex - 1) = gStrDecToLong(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1), 3)
                                Else
                                    tgLinkDrfRec(llLink).tDrf.lDemo(ilIndex - 1) = Val(Trim$(tmSaveShow(llRowNo).sSave(ilIndexOffset + ilLoop1)))
                                End If
                            Else
                                tgLinkDrfRec(llLink).tDrf.lDemo(ilIndex - 1) = 0
                            End If
                        Else
                        End If
                    End If
                End If
            Next illoop
        End If
    Next llRowNo
    mSetRec 'Move record back into tgAllDrf
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim ilIndexOffset As Integer 'used to replace hard-coded array position offsets
    Dim slStr As String
    Dim llRowNo As Long
    Dim slDays As String
    Dim llUpper As Long
    Dim ilRdf As Integer
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim ilCustomIndex As Integer
    Dim llDemo As Long
    Dim llLink As Long
    Dim llPop As Long
    Dim ilPop As Integer
    
    mGetRec 'Move records from tgAllDrf into tgDrfRec
    llUpper = UBound(tgDrfRec)
    ReDim tmSaveShow(0 To llUpper) As SAVESHOW
    
    '----------------------------------------------------------------
    'Research Header
    '----------------------------------------------------------------
    smSSave(NAMEINDEX) = Trim$(tgDnf.sBookName)
    gUnpackDate tgDnf.iBookDate(0), tgDnf.iBookDate(1), smSSave(DATEINDEX)
    
    'POPSRCE
    smSSave(POPSRCDESCINDEX) = ""
    If tgDnf.iPopDnfCode > 0 Then
        For illoop = 0 To lbcPopSrce.ListCount - 1 Step 1
            If tgDnf.iPopDnfCode = lbcPopSrce.ItemData(illoop) Then
                smSSave(POPSRCDESCINDEX) = lbcPopSrce.List(illoop)
                Exit For
            End If
        Next illoop
    End If
    
    'QUALPOPSRCE
    smSSave(QUALSRCDESCINDEX) = ""
    If tgDnf.iQualPopDnfCode > 0 Then
        For illoop = 0 To lbcPopSrce.ListCount - 1 Step 1
            If tgDnf.iQualPopDnfCode = lbcPopSrce.ItemData(illoop) Then
                smSSave(QUALSRCDESCINDEX) = lbcPopSrce.List(illoop)
                Exit For
            End If
        Next illoop
    End If
    imEstByLOrU = 0
    If tgDnf.sEstListenerOrUSA = "U" Then
        imEstByLOrU = 1
    End If
    pbcEstByLorU.Cls
    pbcEstByLorU_Paint
    
    'Demos
    ilIndex = 1
    ilIndexOffset = POPINDEX - 1
    If imCustomIndex <= 0 Then
        For illoop = 1 To 18 Step 1
            If tgSDrfPop.lDemo(illoop - 1) > 0 Then
                If tgSpf.sSAudData = "H" Then
                    slStr = gLongToStrDec(tgSDrfPop.lDemo(illoop - 1), 1)
                ElseIf tgSpf.sSAudData = "N" Then
                    slStr = gLongToStrDec(tgSDrfPop.lDemo(illoop - 1), 2)
                ElseIf tgSpf.sSAudData = "U" Then
                    slStr = gLongToStrDec(tgSDrfPop.lDemo(illoop - 1), 3)
                Else
                    slStr = Trim$(Str$(tgSDrfPop.lDemo(illoop - 1)))
                End If
            Else
                slStr = ""
            End If
            smSSave(ilIndexOffset + ilIndex) = slStr
            If (smDataForm <> "8") And ((illoop = 8) Or (illoop = 16)) Then
                smSSave(ilIndexOffset + ilIndex + 1) = ""
                ilIndex = ilIndex + 2
                If illoop = 16 Then
                    Exit For
                End If
            Else
                ilIndex = ilIndex + 1
            End If
        Next illoop
    Else
        For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
            llPop = -1
            If illoop < UBound(tmCustInfo) Then
                For ilPop = imLBDrf To UBound(tgCDrfPop) - 1 Step 1
                    If tgCDrfPop(ilPop).sDataType = tmCustInfo(illoop).sDataType Then
                        llPop = tgCDrfPop(ilPop).lDemo(tmCustInfo(illoop).iDemoIndex - 1)
                        Exit For
                    End If
                Next ilPop
                If llPop > 0 Then
                    If tgSpf.sSAudData = "H" Then
                        slStr = gLongToStrDec(llPop, 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        slStr = gLongToStrDec(llPop, 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        slStr = gLongToStrDec(llPop, 3)
                    Else
                        slStr = Trim$(Str$(llPop))
                    End If
                Else
                    slStr = ""
                End If
                smSSave(ilIndexOffset + ilIndex) = slStr
                ilIndex = ilIndex + 1
            Else
                smSSave(ilIndexOffset + ilIndex) = ""
                ilIndex = ilIndex + 1
            End If
        Next illoop
    End If
    
    '----------------------------------------------------------------
    'Research Rows
    '----------------------------------------------------------------
    For llRowNo = imLBDrf To UBound(tgDrfRec) - 1 Step 1
        '------------------------
        'Vehicle
        tmSaveShow(llRowNo).sSave(1) = ""
        ilVefCode = tgDrfRec(llRowNo).tDrf.iVefCode
        ilVef = gBinarySearchVef(tgDrfRec(llRowNo).tDrf.iVefCode)
        If ilVef <> -1 Then
            tmSaveShow(llRowNo).sSave(1) = Trim$(tgMVef(ilVef).sName)
        End If
        
        '------------------------
        'ACT1Code and ACT1Setting
        Select Case imDataType
            Case 0 'Daypart
                'ACT1 CODE
                tmSaveShow(llRowNo).sSave(DACT1CODEINDEX) = ""
                tmSaveShow(llRowNo).sSave(DACT1CODEINDEX) = Trim$(tgDrfRec(llRowNo).tDrf.sACTLineupCode)
                'ACT1 SETTING
                tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) = ""
                If tgDrfRec(llRowNo).tDrf.sACT1StoredTime = "T" Then tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) = "T"
                If tgDrfRec(llRowNo).tDrf.sACT1StoredSpots = "S" Then tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) & "S"
                If tgDrfRec(llRowNo).tDrf.sACT1StoreClearPct = "C" Then tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) & "C"
                If tgDrfRec(llRowNo).tDrf.sACT1DaypartFilter = "F" Then tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(DACT1SETTINGINDEX) & "F"
            
            Case 1 'Extra Daypart
                'ACT1 CODE
                tmSaveShow(llRowNo).sSave(XACT1CODEINDEX) = ""
                tmSaveShow(llRowNo).sSave(XACT1CODEINDEX) = Trim$(tgDrfRec(llRowNo).tDrf.sACTLineupCode)
                'ACT1 SETTING
                tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) = ""
                If tgDrfRec(llRowNo).tDrf.sACT1StoredTime = "T" Then tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) = "T"
                If tgDrfRec(llRowNo).tDrf.sACT1StoredSpots = "S" Then tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) & "S"
                If tgDrfRec(llRowNo).tDrf.sACT1StoreClearPct = "C" Then tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) & "C"
                If tgDrfRec(llRowNo).tDrf.sACT1DaypartFilter = "F" Then tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(XACT1SETTINGINDEX) & "F"
            
            Case 2 'Time
            
            Case 3 'Vehicle
                'ACT1 CODE
                tmSaveShow(llRowNo).sSave(VACT1CODEINDEX) = ""
                tmSaveShow(llRowNo).sSave(VACT1CODEINDEX) = Trim$(tgDrfRec(llRowNo).tDrf.sACTLineupCode)
                'ACT1 SETTING
                tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) = ""
                If tgDrfRec(llRowNo).tDrf.sACT1StoredTime = "T" Then tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) = "T"
                If tgDrfRec(llRowNo).tDrf.sACT1StoredSpots = "S" Then tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) & "S"
                If tgDrfRec(llRowNo).tDrf.sACT1StoreClearPct = "C" Then tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) & "C"
                If tgDrfRec(llRowNo).tDrf.sACT1DaypartFilter = "F" Then tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) = tmSaveShow(llRowNo).sSave(VACT1SETTINGINDEX) & "F"

        End Select
        
        '------------------------
        'Daypart/Times
        Select Case imDataType
            Case 0 'Daypart
                tmSaveShow(llRowNo).sSave(DDAYPARTINDEX) = ""
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode <> 0) Then
                    'Daypart
                    ilRdf = gBinarySearchRdf(tgDrfRec(llRowNo).tDrf.iRdfCode)
                    If ilRdf <> -1 Then
                        tmSaveShow(llRowNo).sSave(DDAYPARTINDEX) = Trim$(tgMRdf(ilRdf).sName)
                    End If
                End If
                
            Case 1 'Extra Daypart
                tmSaveShow(llRowNo).sSave(XTIMEINDEX) = ""
                tmSaveShow(llRowNo).sSave(XTIMEINDEX + 1) = ""
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0) Then
                    'Start Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(XTIMEINDEX) = slStr
                    'End Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(XTIMEINDEX + 1) = slStr
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "T" Then
                    'Start Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(XTIMEINDEX) = slStr
                    'End Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(XTIMEINDEX + 1) = slStr
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "V" Then
                    'I guess we do nothing?
                End If
                
            Case 2 'Time
                tmSaveShow(llRowNo).sSave(TTIMEINDEX) = ""
                tmSaveShow(llRowNo).sSave(TTIMEINDEX + 1) = ""
                If (tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0) Then
                    'Start Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(TTIMEINDEX) = slStr
                    'End Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(TTIMEINDEX + 1) = slStr
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "T" Then
                    'Start Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iStartTime(0), tgDrfRec(llRowNo).tDrf.iStartTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(TTIMEINDEX) = slStr
                    'End Time
                    gUnpackTime tgDrfRec(llRowNo).tDrf.iEndTime(0), tgDrfRec(llRowNo).tDrf.iEndTime(1), "A", "1", slStr
                    tmSaveShow(llRowNo).sSave(TTIMEINDEX + 1) = slStr
                ElseIf tgDrfRec(llRowNo).tDrf.sInfoType = "V" Then
                    'I guess we do nothing?
                End If
                
            Case 3 'Vehicle

        End Select
        
        '------------------------
        'Group #
        Select Case imDataType
            Case 0 'Daypart
                tmSaveShow(llRowNo).sSave(DAIRTIMEGRPNOINDEX) = ""
                For illoop = imLBMnf To UBound(tgMnfSocEco) - 1 Step 1
                    If tgDrfRec(llRowNo).tDrf.iMnfSocEco = tgMnfSocEco(illoop).iCode Then
                        tmSaveShow(llRowNo).sSave(DAIRTIMEGRPNOINDEX) = Trim$(tgMnfSocEco(illoop).sUnitType)
                        Exit For
                    End If
                Next illoop
            Case 1 'Extra Daypart
                tmSaveShow(llRowNo).sSave(XGROUPNINDEX) = ""
                For illoop = imLBMnf To UBound(tgMnfSocEco) - 1 Step 1
                    If tgDrfRec(llRowNo).tDrf.iMnfSocEco = tgMnfSocEco(illoop).iCode Then
                        tmSaveShow(llRowNo).sSave(XGROUPNINDEX) = Trim$(tgMnfSocEco(illoop).sUnitType)
                        Exit For
                    End If
                Next illoop
                
            Case 2 'Time
            
            Case 3 'Vehicle
        End Select
        
        '------------------------
        'Days
        ilIndexOffset = 0
        Select Case imDataType
            Case 0 'Daypart
            Case 1 'Extra Daypart
                ilIndexOffset = XDAYSINDEX
            Case 2 'Time
                ilIndexOffset = TDAYSINDEX
            Case 3 'Vehicle
                ilIndexOffset = VDAYSINDEX
        End Select
        If ilIndexOffset > 0 And ((tgDrfRec(llRowNo).tDrf.sInfoType = "D") And (tgDrfRec(llRowNo).tDrf.iRdfCode = 0)) Or (tgDrfRec(llRowNo).tDrf.sInfoType = "T") Or (tgDrfRec(llRowNo).tDrf.sInfoType = "V") Then
            'Days
            slDays = ""
            For illoop = 0 To 6 Step 1
                slDays = slDays & tgDrfRec(llRowNo).tDrf.sDay(illoop)
            Next illoop
            If slDays = "YYYYYNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(0)
            ElseIf slDays = "NNNNNYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(1)
            ElseIf slDays = "NNNNNNY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(2)
            ElseIf slDays = "YYYYYYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(3)
            ElseIf slDays = "YYYYYYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(4)
            ElseIf slDays = "NNNNNYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(5)
            ElseIf slDays = "NYYYYYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(6)
            ElseIf slDays = "NYYYYNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(7)
            ElseIf slDays = "NNYYYYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(8)
            ElseIf slDays = "YNNNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(9)
            ElseIf slDays = "NYNNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(10)
            ElseIf slDays = "NNYNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(11)
            ElseIf slDays = "NNNYNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(12)
            ElseIf slDays = "NNNNYNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(13)
            ElseIf slDays = "YYYYNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(14)
            ElseIf slDays = "YYYNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(15)
            ElseIf slDays = "YYNNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(16)
            ElseIf slDays = "NYYYYYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(17)
            ElseIf slDays = "NYYYNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(18)
            ElseIf slDays = "NYYNNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(19)
            ElseIf slDays = "NNYYYYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(20)
            ElseIf slDays = "NNYYYNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(21)
            ElseIf slDays = "NNYYNNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(22)
            ElseIf slDays = "NNNYYYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(23)
            ElseIf slDays = "NNNYYYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(24)
            ElseIf slDays = "NNNYYNN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(25)
            ElseIf slDays = "NNNNYYY" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(26)
            ElseIf slDays = "NNNNYYN" Then
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = lbcDays.List(27)
            Else
                tmSaveShow(llRowNo).sSave(ilIndexOffset) = ""
            End If
        End If
        
        '------------------------
        'Demos
        Select Case imDataType
            Case 0 'Daypart
                ilIndexOffset = DDEMOINDEX - 1
            Case 1 'Extra Daypart
                ilIndexOffset = XDEMOINDEX - 1
            Case 2 'Time
                ilIndexOffset = TDEMOINDEX - 1
            Case 3 'Vehicle
                ilIndexOffset = VDEMOINDEX - 1
        End Select
        ilIndex = 1
        If imCustomIndex <= 0 Then
            For illoop = 1 To 18 Step 1
                If tgDrfRec(llRowNo).tDrf.lDemo(illoop - 1) > 0 Then
                    If tgSpf.sSAudData = "H" Then
                        slStr = gLongToStrDec(tgDrfRec(llRowNo).tDrf.lDemo(illoop - 1), 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        slStr = gLongToStrDec(tgDrfRec(llRowNo).tDrf.lDemo(illoop - 1), 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        slStr = gLongToStrDec(tgDrfRec(llRowNo).tDrf.lDemo(illoop - 1), 3)
                    Else
                        slStr = Trim$(Str$(tgDrfRec(llRowNo).tDrf.lDemo(illoop - 1)))
                    End If
                Else
                    slStr = ""
                End If
                tmSaveShow(llRowNo).sSave(ilIndexOffset + ilIndex) = slStr
                If (smDataForm <> "8") And ((illoop = 8) Or (illoop = 16)) Then
                    tmSaveShow(llRowNo).sSave(ilIndexOffset + ilIndex + 1) = ""
                    ilIndex = ilIndex + 2
                    If illoop = 16 Then
                        Exit For
                    End If
                Else
                    ilIndex = ilIndex + 1
                End If
            Next illoop
            mGetImpressions llRowNo
        Else
            For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
                llDemo = -1
                If illoop < UBound(tmCustInfo) Then
                    If tgDrfRec(llRowNo).tDrf.sDataType = tmCustInfo(illoop).sDataType Then
                        llDemo = tgDrfRec(llRowNo).tDrf.lDemo(tmCustInfo(illoop).iDemoIndex - 1)
                    Else
                        llLink = tgDrfRec(llRowNo).lLink
                        Do While llLink <> -1
                            If tgLinkDrfRec(llLink).tDrf.sDataType = tmCustInfo(illoop).sDataType Then
                                llDemo = tgLinkDrfRec(llLink).tDrf.lDemo(tmCustInfo(illoop).iDemoIndex - 1)
                                Exit Do
                            End If
                            llLink = tgLinkDrfRec(llLink).lLink
                        Loop
                    End If
                End If
                If llDemo > 0 Then
                    If tgSpf.sSAudData = "H" Then
                        slStr = gLongToStrDec(llDemo, 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        slStr = gLongToStrDec(llDemo, 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        slStr = gLongToStrDec(llDemo, 3)
                    Else
                        slStr = Trim$(Str$(llDemo))
                    End If
                Else
                    slStr = ""
                End If
                tmSaveShow(llRowNo).sSave(ilIndexOffset + ilIndex) = slStr
                ilIndex = ilIndex + 1
            Next illoop
        End If
    Next llRowNo

    If smSource <> "I" Then 'Standard Airtime mode
        If tgSpf.sDemoEstAllowed = "Y" Then
            pbcDPorEst_KeyPress Asc("E")
        Else 'Podcast Impression mode
            pbcDPorEst_KeyPress Asc("P")
        End If
    End If
    
    imSettingValue = True
    vbcDemo.Min = imLBSaveShow  'LBound(tmSaveShow)
    imSettingValue = True
    If UBound(tmSaveShow) <= vbcDemo.LargeChange + 1 Then ' + 1 Then
        vbcDemo.Max = imLBSaveShow  'LBound(tmSaveShow)
    Else
        vbcDemo.Max = UBound(tmSaveShow) - vbcDemo.LargeChange '- 1
    End If
    imSettingValue = True
    vbcDemo.Value = vbcDemo.Min

    mComputeTotalPop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSocEco                   *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMnfSocEco            *
'*                                                     *
'*******************************************************
Private Function mObtainSocEco() As Integer
'
'   ilRet = mObtainSocEco ()
'   Where:
'       tgMnfSocEco() (I)- MNF record structure
'       ilRet (O)- True = populated; False = error
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim slName As String
    Dim slGroup As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    ReDim tgMnfSocEco(0 To 1) As MNF
    ilRecLen = Len(tmMnf) 'btrRecordLength(hmMnf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmMnf, tmMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainSocEco = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainSocEco = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "F"
    ilOffSet = 2 'gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hmMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0 'gFieldOffset("Mnf", "MnfCode")
    ilRet = btrExtAddField(hmMnf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSocEco = False
        Exit Function
    End If
    ilUpperBound = UBound(tgMnfSocEco)
    ilExtLen = Len(tgMnfSocEco(ilUpperBound))  'Extract operation record size
    ilRet = btrExtGetNext(hmMnf, tgMnfSocEco(ilUpperBound), ilExtLen, llRecPos)
    ilExtLen = Len(tgMnfSocEco(ilUpperBound))  'Extract operation record size
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hmMnf, tgMnfSocEco(ilUpperBound), ilExtLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        ilUpperBound = ilUpperBound + 1
        ReDim Preserve tgMnfSocEco(0 To ilUpperBound) As MNF
        ilRet = btrExtGetNext(hmMnf, tgMnfSocEco(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmMnf, tgMnfSocEco(ilUpperBound), ilExtLen, llRecPos)
        Loop
    Loop
    mObtainSocEco = True
    If ilUpperBound = imLBMnf Then
        gGetSyncDateTime smSyncDate, smSyncTime
        For illoop = 1 To 135 Step 1
            tmMnf.iCode = 0
            tmMnf.sType = "F"
            gGetGroupName illoop, slName, slGroup
            tmMnf.sName = slName
            tmMnf.sRPU = ""
            tmMnf.sUnitType = slGroup
            tmMnf.iMerge = 0
            tmMnf.iGroupNo = 0
            tmMnf.sCodeStn = ""
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            Do
                tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmMnf.iAutoCode = tmMnf.iCode
                gPackDate smSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                gPackTime smSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            tgMnfSocEco(ilUpperBound) = tmMnf
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgMnfSocEco(0 To ilUpperBound) As MNF
        Next illoop
    End If
    For illoop = imLBMnf To UBound(tgMnfSocEco) - 1 Step 1
        lbcSocEcoCode.AddItem Trim$(tgMnfSocEco(illoop).sUnitType) & "\" & Trim$(Str$(tgMnfSocEco(illoop).iCode))
    Next illoop
    For illoop = 0 To lbcSocEcoCode.ListCount - 1 Step 1
        slNameCode = lbcSocEcoCode.List(illoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        lbcSocEco.AddItem slName
    Next illoop
    lbcSocEco.AddItem "[None]", 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slName As String
    slName = Trim$(smSSave(NAMEINDEX)) & ": " & Trim$(smSSave(2))
    gFindMatch slName, 0, cbcSelect    'Determine if name exist
    If gLastFound(cbcSelect) <> -1 Then   'Name found
        If gLastFound(cbcSelect) <> imSelectedIndex Then
            If slName = cbcSelect.List(gLastFound(cbcSelect)) Then
                Beep
                MsgBox "Book Name and Date already defined, enter a different name or date", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                imSBoxNo = NAMEINDEX
                mOKName = False
                Exit Function
            End If
        End If
    End If
    mOKName = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Book name list box    *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilVefCode As Integer
    Dim ilSort As Integer
    Dim ilShow As Integer
    ilIndex = cbcSelect.ListIndex
    If ilIndex > 1 Then
        slName = cbcSelect.List(ilIndex)
    End If

    '2/23/19: Filter books shown
    mApplyFilter
    imChgMode = True
    If ilIndex >= 0 Then
        gFindMatch slName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            cbcSelect.ListIndex = -1
        End If
    Else
        cbcSelect.ListIndex = ilIndex
    End If
    imChgMode = False
    
    Exit Sub
mPopluateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilDnfCode As Integer, ilModel As Integer) As Integer
'
'   iRet = mReadRec (ilDnfCode)
'   Where:
'       ilDnfCode(I)- Dnf code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llUpper As Long
    Dim illoop As Integer
    Dim llLoop As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim tlDrf As DRF
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    ReDim tgDrfRec(0 To 1) As DRFREC
    ReDim tgDrfDel(0 To 1) As DRFREC
    ReDim tgAllDrf(0 To 1) As DRFREC
    ReDim tgDpfRec(0 To 1) As DPFREC
    ReDim tgDpfDel(0 To 1) As DPFREC
    ReDim tgAllDpf(0 To 1) As DPFREC
    ReDim tgDefRec(0 To 1) As DEFREC
    ReDim tgDefDel(0 To 1) As DEFREC
    ReDim tgCDrfPop(0 To 1) As DRF
    ReDim tgGDrfPop(0 To 1) As DRF
    mInitNewDrf False, UBound(tgDrfRec)
    mInitNewDpf
    mInitNewDef
    imDrfChg = False
    imDnfChg = False
    imPopChg = False
    imDpfChg = False
    imDefChg = False
    tmDnfSrchKey.iCode = ilDnfCode
    ilRet = btrGetEqual(hmDnf, tgDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mReadRec = False
        Exit Function
    End If
    smSource = tgDnf.sSource
    imEstByLOrU = 0
    If tgDnf.sEstListenerOrUSA = "U" Then
        imEstByLOrU = 1
    End If
    smDataForm = Trim$(tgDnf.sForm)
    If smDataForm <> "8" Then
        smStdDemo(0) = "M12-17"
        smStdDemo(1) = "M18-24"
        smStdDemo(2) = "M25-34"
        smStdDemo(3) = "M35-44"
        smStdDemo(4) = "M45-49"
        smStdDemo(5) = "M50-54"
        smStdDemo(6) = "M55-64"
        smStdDemo(7) = "M65+"
        smStdDemo(8) = ""
        smStdDemo(9) = "W12-17"
        smStdDemo(10) = "W18-24"
        smStdDemo(11) = "W25-34"
        smStdDemo(12) = "W35-44"
        smStdDemo(13) = "W45-49"
        smStdDemo(14) = "W50-54"
        smStdDemo(15) = "W55-64"
        smStdDemo(16) = "W65+"
        smStdDemo(17) = ""
    Else
        smStdDemo(0) = "M12-17"
        smStdDemo(1) = "M18-20"
        smStdDemo(2) = "M21-24"
        smStdDemo(3) = "M25-34"
        smStdDemo(4) = "M35-44"
        smStdDemo(5) = "M45-49"
        smStdDemo(6) = "M50-54"
        smStdDemo(7) = "M55-64"
        smStdDemo(8) = "M65+"
        smStdDemo(9) = "W12-17"
        smStdDemo(10) = "W18-20"
        smStdDemo(11) = "W21-24"
        smStdDemo(12) = "W25-34"
        smStdDemo(13) = "W35-44"
        smStdDemo(14) = "W45-49"
        smStdDemo(15) = "W50-54"
        smStdDemo(16) = "W55-64"
        smStdDemo(17) = "W65+"
    End If
    lmSDrfPopRecPos = 0
    For illoop = LBound(tgSDrfPop.lDemo) To UBound(tgSDrfPop.lDemo) Step 1
        tgSDrfPop.lDemo(illoop) = 0
    Next illoop
    tmDrfSrchKey.iDnfCode = tgDnf.iCode
    tmDrfSrchKey.sDemoDataType = "P"
    tmDrfSrchKey.iMnfSocEco = 0
    tmDrfSrchKey.iVefCode = 0
    tmDrfSrchKey.sInfoType = ""
    tmDrfSrchKey.iRdfCode = 0
    ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tgDnf.iCode) And (tlDrf.sDemoDataType = "P")
            If tlDrf.iMnfSocEco = 0 Then
                If tlDrf.sDataType = "A" Then
                    ilRet = btrGetPosition(hmDrf, lmSDrfPopRecPos)
                    tgSDrfPop = tlDrf
                ElseIf tlDrf.sDataType <> "C" Then
                    tgCDrfPop(UBound(tgCDrfPop)) = tlDrf
                    ReDim Preserve tgCDrfPop(0 To UBound(tgCDrfPop) + 1) As DRF
                End If
            Else
                tgGDrfPop(UBound(tgGDrfPop)) = tlDrf
                ReDim Preserve tgGDrfPop(0 To UBound(tgGDrfPop) + 1) As DRF
            End If
            ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    llUpper = UBound(tgAllDrf)
    btrExtClear hmDrf   'Clear any previous extend operation
    ilExtLen = Len(tgAllDrf(1).tDrf)  'Extract operation record size
    tmDrfSrchKey.iDnfCode = tgDnf.iCode
    tmDrfSrchKey.sDemoDataType = "D"
    tmDrfSrchKey.iMnfSocEco = 0
    tmDrfSrchKey.iVefCode = 0
    tmDrfSrchKey.sInfoType = ""
    tmDrfSrchKey.iRdfCode = 0
    ilRet = btrGetGreaterOrEqual(hmDrf, tgAllDrf(llUpper).tDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmDrf, llNoRec, -1, "UC", "DRF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Drf", "DrfDnfCode")
        tlIntTypeBuff.iType = tgDnf.iCode    'Val(slCode)
        ilRet = btrExtAddLogicConst(hmDrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Drf.Btr", Research
        On Error GoTo 0
        If ckcSocEco.Value = vbUnchecked Then
            ilOffSet = gFieldOffset("Drf", "DrfMnfSocEco")
            tlIntTypeBuff.iType = 0    'Val(slCode)
            ilRet = btrExtAddLogicConst(hmDrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Drf.Btr", Research
            On Error GoTo 0
        End If
        tlCharTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Drf", "DrfDemoDataType")
        ilRet = btrExtAddLogicConst(hmDrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Drf.Btr", Research
        On Error GoTo 0
        ilRet = btrExtAddField(hmDrf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddField):" & "Drf.Btr", Research
        On Error GoTo 0
        ilRet = btrExtGetNext(hmDrf, tgAllDrf(llUpper).tDrf, ilExtLen, tgAllDrf(llUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Drf.Btr", Research
            On Error GoTo 0
            ilExtLen = Len(tgAllDrf(llUpper).tDrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmDrf, tgAllDrf(llUpper).tDrf, ilExtLen, tgAllDrf(llUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tgAllDrf(llUpper).sKey = ""
                tgAllDrf(llUpper).iStatus = 1
                tgAllDrf(llUpper).lIndex = 0
                tgAllDrf(llUpper).iModel = False
                tgAllDrf(llUpper).lLink = -1
                llUpper = llUpper + 1
                ReDim Preserve tgAllDrf(0 To llUpper) As DRFREC
                ilRet = btrExtGetNext(hmDrf, tgAllDrf(llUpper).tDrf, ilExtLen, tgAllDrf(llUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmDrf, tgAllDrf(llUpper).tDrf, ilExtLen, tgAllDrf(llUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    mGetDef ilDnfCode, ilModel
    If ilModel Then
        'If not model, then dpf read in dynamically (read when referenced in mShowDpf)
        mGetDpf ilDnfCode, ilModel
        '11/15/11: Set Model flags
        bmModelUsed = True
        mSetModelFields
        sgPercentChg = ""
    End If
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mResetStatus                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reset Status as Save Failed    *
'*                                                     *
'*******************************************************
Private Sub mResetStatus()
    Dim llLoop As Long
    Dim slInfoType As String
    Dim slDataType As String
    Dim ilRecOK As Integer
    Dim illoop As Integer
    
    If imCustomIndex <= 0 Then
        slDataType = "A"
    End If
    If (rbcDataType(0).Value) Or (rbcDataType(1).Value) Then 'Daypart or Extra Daypart
        slInfoType = "D"
    ElseIf rbcDataType(2).Value Then 'Time
        slInfoType = "T"
    Else 'Vehicle
        slInfoType = "V"
    End If

    For llLoop = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        ilRecOK = False
        If imCustomIndex <= 0 Then
            If (tgAllDrf(llLoop).tDrf.sInfoType = slInfoType) And (tgAllDrf(llLoop).tDrf.sDataType = slDataType) Then
                ilRecOK = True
            End If
        Else
            For illoop = imCustomIndex - 1 To imCustomIndex + 16 Step 1
                If illoop < UBound(tmCustInfo) Then
                    If (tgAllDrf(llLoop).tDrf.sInfoType = slInfoType) And (tgAllDrf(llLoop).tDrf.sDataType = tmCustInfo(illoop).sDataType) Then
                        ilRecOK = True
                        Exit For
                    End If
                End If
            Next illoop
        End If
        If ilRecOK Then
            ilRecOK = False
            If (rbcDataType(0).Value) And (tgAllDrf(llLoop).tDrf.iRdfCode <> 0) And (tgAllDrf(llLoop).tDrf.iCount >= 0) Then 'Daypart
                ilRecOK = True
            ElseIf (rbcDataType(1).Value) And (tgAllDrf(llLoop).tDrf.iRdfCode = 0) And (tgAllDrf(llLoop).tDrf.iCount >= 0) Then 'Extra Daypart
                ilRecOK = True
            ElseIf rbcDataType(2).Value And (tgAllDrf(llLoop).tDrf.sExStdDP <> "Y") And (tgAllDrf(llLoop).tDrf.sExStdDP <> "X") Then 'Time
                ilRecOK = True
            ElseIf rbcDataType(3).Value Then 'Vehicle
                ilRecOK = True
            End If
        End If
        If ilRecOK Then
            tgAllDrf(llLoop).iStatus = -1   'Available record
            tgAllDrf(llLoop).iModel = False
        End If
    Next llLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim ilMap As Integer
    Dim ilCRet As Integer
    Dim ilDay As Integer
    'Dim ilRowNo As Integer
    Dim llRowNo As Long
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilEffTime0 As Integer
    Dim ilEffTime1 As Integer
    Dim slStr As String
    Dim slMsg As String
    Dim llDpf As Long
    Dim llDef As Long
    Dim tlDnf As DNF
    Dim llDrf As Long
    Dim tlDrf As DRF
    Dim tlDrf1 As MOVEREC
    Dim tlDrf2 As MOVEREC
    Dim tlDnf1 As MOVEREC
    Dim tlDnf2 As MOVEREC
    Dim llDateTest As Long
    Dim ilRes As Integer
    Dim ilPop As Integer
    Dim ilFound As Integer
    Dim llLoop As Long
    Dim llBaseDrfCode As Long
    Dim ilBase As Integer
    Dim ilMapUsed As Integer

    slStr = Format$(gNow(), "m/d/yy")
    gPackDate slStr, ilEffDate0, ilEffDate1
    slStr = Format$(gNow(), "h:m:s AM/PM")
    gPackTime slStr, ilEffTime0, ilEffTime1
    gGetSyncDateTime smSyncDate, smSyncTime
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf True
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    mMoveCtrlToRec
    If mTestSSaveFields() = NO Then
        mResetStatus
        Screen.MousePointer = vbDefault
        mSaveRec = False
        Exit Function
    End If
    For llRowNo = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If tgAllDrf(llRowNo).iStatus <> -1 Then
            If mTestTgFields(llRowNo) = NO Then
                mResetStatus
                Screen.MousePointer = vbDefault
                mSaveRec = False
                Exit Function
            End If
        End If
    Next llRowNo
    If mTestForDuplicateRows() Then
        mResetStatus
        Screen.MousePointer = vbDefault
        mSaveRec = False
        Exit Function
    End If
    For llRowNo = imLBDpf To UBound(tgAllDpf) - 1 Step 1
        If tgAllDpf(llRowNo).iStatus <> -1 Then
            If mTestPlusFields(tgAllDpf(llRowNo)) = NO Then
                mResetStatus
                Screen.MousePointer = vbDefault
                mSaveRec = False
                Exit Function
            End If
        End If
    Next llRowNo
    For llRowNo = imLBDef To UBound(tgDefRec) - 1 Step 1
        If tgDefRec(llRowNo).iStatus <> -1 Then
            If mTestEstFields(tgDefRec(llRowNo)) = NO Then
                mResetStatus
                Screen.MousePointer = vbDefault
                mSaveRec = False
                Exit Function
            End If
            'Check for Duplictae dates
            For llDateTest = llRowNo + 1 To UBound(tgDefRec) - 1 Step 1
                If gDateValue(tgDefRec(llDateTest).sStartDate) = gDateValue(tgDefRec(llRowNo).sStartDate) Then
                    ilRes = MsgBox("Duplicate Start Dates not Allowed in Estimate Population", vbOKOnly + vbExclamation, "Incomplete")
                    imEstBoxNo = EDATEINDEX
                    lmEstRowNo = llRowNo
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    mSaveRec = False
                    Exit Function
                End If
            Next llDateTest
        End If
    Next llRowNo
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        mResetStatus
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    '11/15/11: Reset model flags
    If bmModelUsed Then
        mSetModelFields
    End If
    'Update Book Name
    ilRet = btrBeginTrans(hmDnf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        mResetStatus
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
        imTerminate = True
        mSaveRec = False
        Exit Function
    End If
    If imSelectedIndex > 1 Then
        tmDnfSrchKey.iCode = tgDnf.iCode
        ilRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmDnf)
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mResetStatus
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        LSet tlDnf1 = tlDnf
        LSet tlDnf2 = tgDnf
        If StrComp(tlDnf1.sChar, tlDnf2.sChar, 0) <> 0 Then
            Do
                tgDnf.iUrfCode = tgUrf(0).iCode
                gPackDate smSyncDate, tgDnf.iSyncDate(0), tgDnf.iSyncDate(1)
                gPackTime smSyncTime, tgDnf.iSyncTime(0), tgDnf.iSyncTime(1)
                ilRet = btrUpdate(hmDnf, tgDnf, imDnfRecLen)
                If ilRet = BTRV_ERR_CONFLICT Then
                    tmDnfSrchKey.iCode = tgDnf.iCode
                    ilCRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
    Else
        tgDnf.iCode = 0
        tgDnf.iUrfCode = tgUrf(0).iCode
        tgDnf.sType = "C"
        If igDnfModel > 0 Then
            tgDnf.sType = "M"
        End If
        tgDnf.sExactTime = "N"
        If smSource <> "I" Then 'Standard Airtime mode
            tgDnf.sSource = "M"
        Else 'Podcast Impression mode
            tgDnf.sSource = "I"
        End If
        tgDnf.iRemoteID = tgUrf(0).iRemoteUserID
        tgDnf.iAutoCode = tgDnf.iCode
        ilRet = btrInsert(hmDnf, tgDnf, imDnfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmDnf)
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mResetStatus
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        Do
            tgDnf.iRemoteID = tgUrf(0).iRemoteUserID
            tgDnf.iAutoCode = tgDnf.iCode
            gPackDate smSyncDate, tgDnf.iSyncDate(0), tgDnf.iSyncDate(1)
            gPackTime smSyncTime, tgDnf.iSyncTime(0), tgDnf.iSyncTime(1)
            ilRet = btrUpdate(hmDnf, tgDnf, imDnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    'Update Population
    If lmSDrfPopRecPos <> 0 Then
        ilRet = btrGetDirect(hmDrf, tlDrf, imDrfRecLen, lmSDrfPopRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmDnf)
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mResetStatus
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        LSet tlDrf1 = tlDrf
        LSet tlDrf2 = tgSDrfPop
        If StrComp(tlDrf1.sChar, tlDrf2.sChar, 0) <> 0 Then
            tgSDrfPop.iDemoChgDate(0) = ilEffDate0
            tgSDrfPop.iDemoChgDate(1) = ilEffDate1
            tgSDrfPop.iDemoChgTime(0) = ilEffTime0
            tgSDrfPop.iDemoChgTime(1) = ilEffTime1
            Do
                gPackDate smSyncDate, tgSDrfPop.iSyncDate(0), tgSDrfPop.iSyncDate(1)
                gPackTime smSyncTime, tgSDrfPop.iSyncTime(0), tgSDrfPop.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tgSDrfPop, imDrfRecLen)
                If ilRet = BTRV_ERR_CONFLICT Then
                    ilCRet = btrGetDirect(hmDrf, tlDrf, imDrfRecLen, lmSDrfPopRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilCRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmDnf)
                        If (ilCRet = 30000) Or (ilCRet = 30001) Or (ilCRet = 30002) Or (ilCRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
    Else
        tgSDrfPop.lCode = 0
        tgSDrfPop.iDnfCode = tgDnf.iCode
        tgSDrfPop.sDemoDataType = "P"
        tgSDrfPop.iMnfSocEco = 0
        tgSDrfPop.iVefCode = 0
        tgSDrfPop.sInfoType = ""
        tgSDrfPop.iRdfCode = 0
        tgSDrfPop.sProgCode = ""
        tgSDrfPop.iStartTime(0) = 1
        tgSDrfPop.iStartTime(1) = 0
        tgSDrfPop.iEndTime(0) = 1
        tgSDrfPop.iEndTime(1) = 0
        tgSDrfPop.iStartTime2(0) = 1
        tgSDrfPop.iStartTime2(1) = 0
        tgSDrfPop.iEndTime2(0) = 1
        tgSDrfPop.iEndTime2(1) = 0
        For ilDay = 0 To 6 Step 1
            tgSDrfPop.sDay(ilDay) = "Y"
        Next ilDay
        tgSDrfPop.iQHIndex = 0
        tgSDrfPop.iCount = 0
        tgSDrfPop.sExStdDP = "N"
        tgSDrfPop.sExRpt = "N"
        tgSDrfPop.sDataType = "A"
        tgSDrfPop.iDemoChgDate(0) = ilEffDate0
        tgSDrfPop.iDemoChgDate(1) = ilEffDate1
        tgSDrfPop.iDemoChgTime(0) = ilEffTime0
        tgSDrfPop.iDemoChgTime(1) = ilEffTime1
        tgSDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
        tgSDrfPop.lAutoCode = tgSDrfPop.lCode
        tgSDrfPop.sForm = smDataForm
        ilRet = btrInsert(hmDrf, tgSDrfPop, imDrfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmDnf)
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mResetStatus
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        Do
            tgSDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
            tgSDrfPop.lAutoCode = tgSDrfPop.lCode
            gPackDate smSyncDate, tgSDrfPop.iSyncDate(0), tgSDrfPop.iSyncDate(1)
            gPackTime smSyncTime, tgSDrfPop.iSyncTime(0), tgSDrfPop.iSyncTime(1)
            ilRet = btrUpdate(hmDrf, tgSDrfPop, imDrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmDnf)
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mResetStatus
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
    End If
    'Update Population
    For ilPop = imLBDrf To UBound(tgCDrfPop) - 1 Step 1
        If tgCDrfPop(ilPop).lCode > 0 Then
            tmDrfSrchKey2.lCode = tgCDrfPop(ilPop).lCode
            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            LSet tlDrf1 = tlDrf
            LSet tlDrf2 = tgCDrfPop(ilPop)
            If StrComp(tlDrf1.sChar, tlDrf2.sChar, 0) <> 0 Then
                tgCDrfPop(ilPop).iDemoChgDate(0) = ilEffDate0
                tgCDrfPop(ilPop).iDemoChgDate(1) = ilEffDate1
                tgCDrfPop(ilPop).iDemoChgTime(0) = ilEffTime0
                tgCDrfPop(ilPop).iDemoChgTime(1) = ilEffTime1
                Do
                    gPackDate smSyncDate, tgCDrfPop(ilPop).iSyncDate(0), tgCDrfPop(ilPop).iSyncDate(1)
                    gPackTime smSyncTime, tgCDrfPop(ilPop).iSyncTime(0), tgCDrfPop(ilPop).iSyncTime(1)
                    ilRet = btrUpdate(hmDrf, tgCDrfPop(ilPop), imDrfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmDrfSrchKey2.lCode = tgCDrfPop(ilPop).lCode
                        ilCRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilCRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmDnf)
                            If (ilCRet = 30000) Or (ilCRet = 30001) Or (ilCRet = 30002) Or (ilCRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            mResetStatus
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                            imTerminate = True
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
        Else
            If Trim$(tgCDrfPop(ilPop).sDataType) <> "" Then
                tgCDrfPop(ilPop).lCode = 0
                tgCDrfPop(ilPop).iDnfCode = tgDnf.iCode
                tgCDrfPop(ilPop).sDemoDataType = "P"
                tgCDrfPop(ilPop).iMnfSocEco = 0
                tgCDrfPop(ilPop).iVefCode = 0
                tgCDrfPop(ilPop).sInfoType = ""
                tgCDrfPop(ilPop).iRdfCode = 0
                tgCDrfPop(ilPop).sProgCode = ""
                tgCDrfPop(ilPop).iStartTime(0) = 1
                tgCDrfPop(ilPop).iStartTime(1) = 0
                tgCDrfPop(ilPop).iEndTime(0) = 1
                tgCDrfPop(ilPop).iEndTime(1) = 0
                tgCDrfPop(ilPop).iStartTime2(0) = 1
                tgCDrfPop(ilPop).iStartTime2(1) = 0
                tgCDrfPop(ilPop).iEndTime2(0) = 1
                tgCDrfPop(ilPop).iEndTime2(1) = 0
                For ilDay = 0 To 6 Step 1
                    tgCDrfPop(ilPop).sDay(ilDay) = "Y"
                Next ilDay
                tgCDrfPop(ilPop).iQHIndex = 0
                tgCDrfPop(ilPop).iCount = 0
                tgCDrfPop(ilPop).sExStdDP = "N"
                tgCDrfPop(ilPop).sExRpt = "N"
                tgCDrfPop(ilPop).iDemoChgDate(0) = ilEffDate0
                tgCDrfPop(ilPop).iDemoChgDate(1) = ilEffDate1
                tgCDrfPop(ilPop).iDemoChgTime(0) = ilEffTime0
                tgCDrfPop(ilPop).iDemoChgTime(1) = ilEffTime1
                tgCDrfPop(ilPop).iRemoteID = tgUrf(0).iRemoteUserID
                tgCDrfPop(ilPop).lAutoCode = tgCDrfPop(ilPop).lCode
                tgCDrfPop(ilPop).sForm = smDataForm
                ilRet = btrInsert(hmDrf, tgCDrfPop(ilPop), imDrfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                Do
                    tgCDrfPop(ilPop).iRemoteID = tgUrf(0).iRemoteUserID
                    tgCDrfPop(ilPop).lAutoCode = tgCDrfPop(ilPop).lCode
                    gPackDate smSyncDate, tgCDrfPop(ilPop).iSyncDate(0), tgCDrfPop(ilPop).iSyncDate(1)
                    gPackTime smSyncTime, tgCDrfPop(ilPop).iSyncTime(0), tgCDrfPop(ilPop).iSyncTime(1)
                    ilRet = btrUpdate(hmDrf, tgCDrfPop(ilPop), imDrfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
        End If
    Next ilPop
    '4/17/15: Remove extra images
    For llDrf = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If (tgAllDrf(llDrf).iStatus = -1) And (tgAllDrf(llDrf).tDrf.lCode > 0) Then
            ilFound = False
            For llLoop = imLBDrf To UBound(tgAllDrf) - 1 Step 1
                If tgAllDrf(llLoop).iStatus = 1 Then
                    If tgAllDrf(llLoop).tDrf.lCode = tgAllDrf(llDrf).tDrf.lCode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next llLoop
            If Not ilFound Then
                Do
                    tmDrfSrchKey2.lCode = tgAllDrf(llDrf).tDrf.lCode
                    ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmDnf)
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmDrf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                'TTP 10759 - Research List screen - Impressions book: manually added/edited impressions can disappear or change when saving
                'Deleting a research line was Not deleting the accociated DRF record
                Do
                    tmDpfSrchKey1.lDrfCode = tgAllDrf(llDrf).tDrf.lCode
                    tmDpfSrchKey1.iMnfDemo = imP12PlusMnfCode
                    ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE And ilRet <> 9 Then  ' TTP 10908 JJB 2023-12-20
                        ilCRet = btrAbortTrans(hmDnf)
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmDpf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                For llLoop = imLBDrf To UBound(tgDrfDel) - 1 Step 1
                    If tgDrfDel(llLoop).tDrf.lCode = tgAllDrf(llDrf).tDrf.lCode Then
                        tgDrfDel(llLoop).tDrf.lCode = -1
                        Exit For
                    End If
                Next llLoop
            End If
        End If
    Next llDrf
    ReDim tmDrfMap(0 To 0) As DRFMAP
    For llDrf = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If tgAllDrf(llDrf).iStatus = 1 Then
            slMsg = "mSaveRec (btrGetDirect: Research Demo Data)"
            tmDrfSrchKey2.lCode = tgAllDrf(llDrf).tDrf.lCode
            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            LSet tlDrf1 = tlDrf
            LSet tlDrf2 = tgAllDrf(llDrf).tDrf
            If StrComp(tlDrf1.sChar, tlDrf2.sChar, 0) <> 0 Then
                slMsg = "mSaveRec (btrUpdate: Research Demo Data)"
                tgAllDrf(llDrf).tDrf.iDemoChgDate(0) = ilEffDate0
                tgAllDrf(llDrf).tDrf.iDemoChgDate(1) = ilEffDate1
                tgAllDrf(llDrf).tDrf.iDemoChgTime(0) = ilEffTime0
                tgAllDrf(llDrf).tDrf.iDemoChgTime(1) = ilEffTime1
                Do
                    ilRet = btrDelete(hmDrf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmDrfSrchKey2.lCode = tgAllDrf(llDrf).tDrf.lCode
                        ilCRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilCRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmDnf)
                            If (ilCRet = 30000) Or (ilCRet = 30001) Or (ilCRet = 30002) Or (ilCRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            mResetStatus
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                            imTerminate = True
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                tgAllDrf(llDrf).tDrf.lAutoCode = tgAllDrf(llDrf).tDrf.lCode
                gPackDate smSyncDate, tgAllDrf(llDrf).tDrf.iSyncDate(0), tgAllDrf(llDrf).tDrf.iSyncDate(1)
                gPackTime smSyncTime, tgAllDrf(llDrf).tDrf.iSyncTime(0), tgAllDrf(llDrf).tDrf.iSyncTime(1)
                ilRet = btrInsert(hmDrf, tgAllDrf(llDrf).tDrf, imDrfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
            '4/11/11: Because of the fix in Model added 6/22/10 Pre-defined dayparts not added.  This fixes that problem
            tmDrfMap(UBound(tmDrfMap)).lInitDrfCode = tgAllDrf(llDrf).tDrf.lCode
            tmDrfMap(UBound(tmDrfMap)).lNewDRfCode = tgAllDrf(llDrf).tDrf.lCode
            tmDrfMap(UBound(tmDrfMap)).iRdfCode = 0
            ReDim Preserve tmDrfMap(0 To UBound(tmDrfMap) + 1) As DRFMAP
        ElseIf tgAllDrf(llDrf).iStatus = 0 Then
            slMsg = "mSaveRec (btrInsert: Research Demo Data)"
            llBaseDrfCode = 0
            tgAllDrf(llDrf).tDrf.iDnfCode = tgDnf.iCode
            If Not tgAllDrf(llDrf).iModel Then
                If tgAllDrf(llDrf).tDrf.lCode < 0 Then
                    llBaseDrfCode = tgAllDrf(llDrf).tDrf.lCode
                End If
                tgAllDrf(llDrf).tDrf.sDemoDataType = "D"
                tgAllDrf(llDrf).tDrf.iStartTime2(0) = 1
                tgAllDrf(llDrf).tDrf.iStartTime2(1) = 0
                tgAllDrf(llDrf).tDrf.iEndTime2(0) = 1
                tgAllDrf(llDrf).tDrf.iEndTime2(1) = 0
                tgAllDrf(llDrf).tDrf.sProgCode = ""
                If tgAllDrf(llDrf).tDrf.sInfoType <> "T" Then
                    tgAllDrf(llDrf).tDrf.iQHIndex = 0
                End If
                tgAllDrf(llDrf).tDrf.iCount = 0
                tgAllDrf(llDrf).tDrf.sExStdDP = "N"
                tgAllDrf(llDrf).tDrf.sExRpt = "N"
                tgAllDrf(llDrf).tDrf.sForm = smDataForm
            Else
                '8/14/18
                tmDrfMap(UBound(tmDrfMap)).lInitDrfCode = tgAllDrf(llDrf).lModelDrfCode
            End If
            tgAllDrf(llDrf).tDrf.lCode = 0
            tgAllDrf(llDrf).tDrf.iDemoChgDate(0) = ilEffDate0
            tgAllDrf(llDrf).tDrf.iDemoChgDate(1) = ilEffDate1
            tgAllDrf(llDrf).tDrf.iDemoChgTime(0) = ilEffTime0
            tgAllDrf(llDrf).tDrf.iDemoChgTime(1) = ilEffTime1
            tgAllDrf(llDrf).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
            tgAllDrf(llDrf).tDrf.lAutoCode = tgAllDrf(llDrf).tDrf.lCode
            ilRet = btrInsert(hmDrf, tgAllDrf(llDrf).tDrf, imDrfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            tgAllDrf(llDrf).iStatus = 1
            If tgAllDrf(llDrf).iModel Then
                tmDrfMap(UBound(tmDrfMap)).lNewDRfCode = tgAllDrf(llDrf).tDrf.lCode
                tmDrfMap(UBound(tmDrfMap)).iRdfCode = 0
                ReDim Preserve tmDrfMap(0 To UBound(tmDrfMap) + 1) As DRFMAP
            Else
                '4/11/11: In the current design, you can NOT add Pre-defined daypart to research not saved.
                'Therefore this code is NOT necessary
                If llBaseDrfCode < 0 Then
                    tmDrfMap(UBound(tmDrfMap)).lInitDrfCode = llBaseDrfCode
                    tmDrfMap(UBound(tmDrfMap)).iRdfCode = tgAllDrf(llDrf).tDrf.iRdfCode
                Else
                    tmDrfMap(UBound(tmDrfMap)).lInitDrfCode = tgAllDrf(llDrf).tDrf.lCode
                    tmDrfMap(UBound(tmDrfMap)).iRdfCode = 0
                End If
                tmDrfMap(UBound(tmDrfMap)).lNewDRfCode = tgAllDrf(llDrf).tDrf.lCode
                ReDim Preserve tmDrfMap(0 To UBound(tmDrfMap) + 1) As DRFMAP
            End If
            tgAllDrf(llDrf).iModel = False
            Do
                tgAllDrf(llDrf).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tgAllDrf(llDrf).tDrf.lAutoCode = tgAllDrf(llDrf).tDrf.lCode
                gPackDate smSyncDate, tgAllDrf(llDrf).tDrf.iSyncDate(0), tgAllDrf(llDrf).tDrf.iSyncDate(1)
                gPackTime smSyncTime, tgAllDrf(llDrf).tDrf.iSyncTime(0), tgAllDrf(llDrf).tDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tgAllDrf(llDrf).tDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmDnf)
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                mResetStatus
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mPutImpressions llDrf
    Next llDrf
    '9/29/15: Bypass any deleted records if using model
    If bmModelUsed = False Then
        For llDrf = imLBDrf To UBound(tgDrfDel) - 1 Step 1
            If (tgDrfDel(llDrf).iStatus = 1) And (tgDrfDel(llDrf).tDrf.lCode > 0) Then
                Do
                    slMsg = "mSaveRec (btrGetDirect: Research Demo Data)"
                    tmDrfSrchKey2.lCode = tgDrfDel(llDrf).tDrf.lCode
                    ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmDnf)
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmDrf)
                    slMsg = "mSaveRec (btrDelete: Research Demo Data)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                mRemoveImpressions llDrf
            End If
        Next llDrf
    End If
    
    If imDpfChg And (smSource <> "I") Then
        For llDpf = imLBDpf To UBound(tgAllDpf) - 1 Step 1
            If (tgAllDpf(llDpf).iStatus = 0) Then
                ilFound = False
                For ilMap = 0 To UBound(tmDrfMap) - 1 Step 1
                    If (tmDrfMap(ilMap).lInitDrfCode = tgAllDpf(llDpf).lDrfCode) And (tgAllDpf(llDpf).sSource <> "B") Then
                        tgAllDpf(llDpf).lDrfCode = tmDrfMap(ilMap).lNewDRfCode
                        ilFound = True
                        Exit For
                    ElseIf (tmDrfMap(ilMap).lInitDrfCode = tgAllDpf(llDpf).lDrfCode) And (tgAllDpf(llDpf).sSource = "B") And (tmDrfMap(ilMap).iRdfCode = tgAllDpf(llDpf).iRdfCode) Then
                        tgAllDpf(llDpf).lDrfCode = tmDrfMap(ilMap).lNewDRfCode
                        ilFound = True
                        Exit For
                    End If
                Next ilMap
                If Not ilFound Then
                    tgAllDpf(llDpf).iStatus = -1
                End If
            End If
        Next llDpf
        
        For llDpf = imLBDpf To UBound(tgAllDpf) - 1 Step 1
            If tgAllDpf(llDpf).iStatus = 1 Then
                slMsg = "mSaveRec (btrGetEqual: Research Plus Demo Data)"
                tmDpfSrchKey.lCode = tgAllDpf(llDpf).lDpfCode
                ilRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                slMsg = "mSaveRec (btrUpdate: Research PlusDemo Data)"
                Do
                    ilRet = btrDelete(hmDpf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmDpfSrchKey.lCode = tgAllDpf(llDpf).lDpfCode
                        ilCRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilCRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmDnf)
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            mResetStatus
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                            imTerminate = True
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                mMovePlusToRec llDpf
                ilRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            ElseIf tgAllDpf(llDpf).iStatus = 0 Then
                slMsg = "mSaveRec (btrInsert: Research Demo Data)"
                mMovePlusToRec llDpf
                tmDpf.lCode = 0
                ilRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                tgAllDpf(llDpf).iStatus = 1
            End If
        Next llDpf
        
        For llDpf = imLBDpf To UBound(tgDpfDel) - 1 Step 1
            If tgDpfDel(llDpf).iStatus = 1 Then
                Do
                    slMsg = "mSaveRec (btrGetEqual: Research Plus Demo Data)"
                    tmDpfSrchKey.lCode = tgDpfDel(llDpf).lDpfCode
                    ilRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE And ilRet <> 4 Then ' TTP 10908 JJB 2023-12-20
                        ilCRet = btrAbortTrans(hmDnf)
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmDpf)
                    slMsg = "mSaveRec (btrDelete: Research Plus Demo Data)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE And ilRet <> 8 Then  ' TTP 10908 JJB 2023-12-20
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
        Next llDpf
    End If

    If imDefChg Then
        For llDef = imLBDef To UBound(tgDefRec) - 1 Step 1
            If tgDefRec(llDef).iStatus = 1 Then
                slMsg = "mSaveRec (btrGetEqual: Research Estimate Population Data)"
                tmDefSrchKey.lCode = tgDefRec(llDef).lDefCode
                ilRet = btrGetEqual(hmDef, tmDef, imDefRecLen, tmDefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                slMsg = "mSaveRec (btrUpdate: Research Estimate Population Data)"
                Do
                    ilRet = btrDelete(hmDef)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmDefSrchKey.lCode = tgDefRec(llDef).lDefCode
                        ilCRet = btrGetEqual(hmDef, tmDef, imDefRecLen, tmDefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilCRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmDnf)
                            If (ilCRet = 30000) Or (ilCRet = 30001) Or (ilCRet = 30002) Or (ilCRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            mResetStatus
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                            imTerminate = True
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                mMoveEstToRec llDef
                ilRet = btrInsert(hmDef, tmDef, imDefRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            ElseIf tgDefRec(llDef).iStatus = 0 Then
                slMsg = "mSaveRec (btrInsert: Research Estimate Population Data)"
                mMoveEstToRec llDef
                tmDef.lCode = 0
                ilRet = btrInsert(hmDef, tmDef, imDefRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                tgDefRec(llDef).iStatus = 1
                tgDefRec(llDef).lDefCode = tmDef.lCode
            End If
        Next llDef
        For llDef = imLBDef To UBound(tgDefDel) - 1 Step 1
            If tgDefDel(llDef).iStatus = 1 Then
                Do
                    slMsg = "mSaveRec (btrGetEqual: Research Estimate Population Data)"
                    tmDefSrchKey.lCode = tgDefDel(llDef).lDefCode
                    ilRet = btrGetEqual(hmDef, tmDef, imDefRecLen, tmDefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmDnf)
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mResetStatus
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmDef)
                    slMsg = "mSaveRec (btrDelete: Research Estimate Population Data)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmDnf)
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mResetStatus
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later: Error #:" & ilRet, vbOKOnly + vbExclamation, "Research")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
        Next llDef
    End If

    ilRet = btrEndTrans(hmDnf)
    '11/15/11: Reset Model flag
    bmModelUsed = False
    bmResearchSaved = True
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    If imDnfChg Or imPopChg Or imDrfChg Or imDpfChg Or imDefChg Then
        If ilAsk Then
            If imSelectedIndex > 1 Then
                slMess = "Save Changes to " & smSSave(NAMEINDEX)
            Else
                slMess = "Add " & smSSave(NAMEINDEX)
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
            If ilRes = vbNo Then
                cmcUndo_Click
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSEnableBox                     *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBSCtrls Or ilBoxNo > UBound(tmSCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Width = tmSCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 30
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY
            edcSpecDropDown.Text = smSSave(NAMEINDEX)
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case DATEINDEX
            edcSpecDropDown.Width = tmSCtrls(DATEINDEX).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 10
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSCtrls(DATEINDEX).fBoxX, tmSCtrls(DATEINDEX).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            If edcSpecDropDown.Top + edcSpecDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            Else
                plcCalendar.Move edcSpecDropDown.Left, edcSpecDropDown.Top - plcCalendar.Height
            End If
            If smSSave(DATEINDEX) = "" Then
                slStr = gObtainMondayFromToday()
            Else
                slStr = smSSave(DATEINDEX)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcSpecDropDown.Text = slStr
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            If smSSave(DATEINDEX) = "" Then
                pbcCalendar.Visible = True
            End If
            edcSpecDropDown.SetFocus
        Case POPSRCEINDEX
            lbcPopSrce.Height = gListBoxHeight(lbcPopSrce.ListCount, 10)
            edcSpecDropDown.Width = tmSCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 0
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcPopSrce.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            imChgMode = True
            gFindMatch Trim$(smSSave(POPSRCDESCINDEX)), 0, lbcPopSrce
            If gLastFound(lbcPopSrce) >= 0 Then
                lbcPopSrce.ListIndex = gLastFound(lbcPopSrce)
            Else
                lbcPopSrce.ListIndex = 0
            End If
            imComboBoxIndex = lbcVehicle.ListIndex
            If lbcPopSrce.ListIndex < 0 Then
                edcSpecDropDown.Text = ""
            Else
                edcSpecDropDown.Text = lbcPopSrce.List(lbcPopSrce.ListIndex)
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case QUALPOPSRCEINDEX
            lbcPopSrce.Height = gListBoxHeight(lbcPopSrce.ListCount, 10)
            edcSpecDropDown.Width = tmSCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 0
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcPopSrce.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            imChgMode = True
            gFindMatch Trim$(smSSave(QUALSRCDESCINDEX)), 0, lbcPopSrce
            If gLastFound(lbcPopSrce) >= 0 Then
                lbcPopSrce.ListIndex = gLastFound(lbcPopSrce)
            Else
                lbcPopSrce.ListIndex = 0
            End If
            imComboBoxIndex = lbcVehicle.ListIndex
            If lbcPopSrce.ListIndex < 0 Then
                edcSpecDropDown.Text = ""
            Else
                edcSpecDropDown.Text = lbcPopSrce.List(lbcPopSrce.ListIndex)
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case POPINDEX To POPINDEX + 17
            edcSpecDropDown.Width = tmSCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 8
            gMoveTableCtrl pbcSpec, edcSpecDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY
            edcSpecDropDown.Text = smSSave(POPINDEX + ilBoxNo - POPINDEX)
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilAltered As Integer
    If (imBypassSetting) Or (Not imUpdateAllowed) Then
        Exit Sub
    End If
    ilAltered = imDnfChg
    If Not ilAltered Then
        ilAltered = imPopChg
    End If
    If Not ilAltered Then
        ilAltered = imDrfChg
    End If
    If Not ilAltered Then
        ilAltered = imDpfChg
    End If
    If Not ilAltered Then
        ilAltered = imDefChg
    End If
    If ilAltered Then
        pbcSpec.Enabled = True
        If (rbcDataType(0).Value) Or (rbcDataType(1).Value) Or (rbcDataType(2).Value) Then 'Daypart,Extra Daypart,Time
            pbcDemo(0).Enabled = True
        Else 'Vehicle
            pbcDemo(1).Enabled = True
        End If
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        pbcPlus.Enabled = True
        pbcPlusSTab.Enabled = True
        pbcPlusTab.Enabled = True
        cbcSelect.Enabled = False
        cmcErase.Enabled = False
    Else
        If imSelectedIndex < 0 Then
            pbcSpec.Enabled = False
            If (rbcDataType(0).Value) Or (rbcDataType(1).Value) Or (rbcDataType(2).Value) Then 'Daypart,Extra Daypart,Time
                pbcDemo(0).Enabled = False
            Else 'Vehicle
                pbcDemo(1).Enabled = False
            End If
            pbcSpecSTab.Enabled = False
            pbcSpecTab.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            pbcPlus.Enabled = False
            pbcPlusSTab.Enabled = False
            pbcPlusTab.Enabled = False
            cbcSelect.Enabled = True
            cmcErase.Enabled = False
        Else
            If Me.ActiveControl.Name <> "CSI_ComboboxMS1" Then
                pbcSpec.Enabled = True
            End If
            If (rbcDataType(0).Value) Or (rbcDataType(1).Value) Or (rbcDataType(2).Value) Then 'Daypart,Extra Daypart,Time
                pbcDemo(0).Enabled = True
            Else 'Vehicle
                pbcDemo(1).Enabled = True
            End If
            pbcSpecSTab.Enabled = True
            pbcSpecTab.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            pbcPlus.Enabled = True
            pbcPlusSTab.Enabled = True
            pbcPlusTab.Enabled = True
            cbcSelect.Enabled = True
            If imUpdateAllowed Then
                cmcErase.Enabled = True
            Else
                cmcErase.Enabled = False
            End If
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And ((UBound(tgDrfRec) > 1) Or (UBound(tgAllDrf) > 1)) Then
        If imUpdateAllowed Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    If imSelectedIndex > 1 Then
        cmcSocEco.Enabled = False
    Else
        If Trim$(smSSave(NAMEINDEX)) = "" Then
            cmcSocEco.Enabled = True
        Else
            cmcSocEco.Enabled = False
        End If
    End If
    If imSelectedIndex >= 0 Then
        cmcAdjust.Enabled = True
    Else
        cmcAdjust.Enabled = False
    End If
    If (lmRowNo < vbcDemo.Value) Or (lmRowNo > (vbcDemo.Value + vbcDemo.LargeChange)) Then
        cmcDuplicate.Enabled = False
    Else
        cmcDuplicate.Enabled = True
    End If
    If rbcDataType(0).Value = False Then 'Daypart
        cmcBaseDuplicate.Enabled = False
    Else
        cmcBaseDuplicate.Enabled = True
    End If
    If imSelectedIndex > 1 Then
        cmcSetDefault.Enabled = True
    Else
        cmcSetDefault.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If rbcDataType(0).Value Then 'Daypart
        If ilBoxNo < imLBDCtrls Or ilBoxNo > UBound(tmDCtrls) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case DVEHICLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            
            Case DDAYPARTINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case DGROUPINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    edcDropDown.Visible = True
                    cmcDropDown.Visible = True
                    edcDropDown.SetFocus
                Else 'Podcast Impression mode
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                End If
            Case DDEMOINDEX To DDEMOINDEX + 17
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        If ilBoxNo < imLBXCtrls Or ilBoxNo > UBound(tmXCtrls) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case XVEHICLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case XTIMEINDEX To XTIMEINDEX + 1
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case XDAYSINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case XGROUPINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case XDEMOINDEX To XDEMOINDEX + 17
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
    ElseIf rbcDataType(2).Value Then 'Time
        If ilBoxNo < imLBTCtrls Or ilBoxNo > UBound(tmTCtrls) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case TVEHICLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case TTIMEINDEX To TTIMEINDEX + 1
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case TDAYSINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case TDEMOINDEX To TDEMOINDEX + 17
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
    Else 'Vehicle
        If ilBoxNo < imLBVCtrls Or ilBoxNo > UBound(tmVCtrls) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case VVEHICLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case VDAYSINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case VDEMOINDEX To VDEMOINDEX + 17
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
        End Select
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetRec                         *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move records back into save    *
'*                                                     *
'*******************************************************
Private Sub mSetRec()
    Dim ilCheck As Integer
    Dim ilFound As Integer
    Dim llLoop As Long
    Dim llIndex As Long
    Dim llLink As Long
    Dim blDemoExist As Boolean
    Dim ilDemo As Integer
    
    ilCheck = True
    For llLoop = imLBDrf To UBound(tgDrfRec) - 1 Step 1
        If (tgDrfRec(llLoop).tDrf.lCode > 0) And (Not tgDrfRec(llLoop).iModel) Then
            blDemoExist = True
            tgDrfRec(llLoop).iStatus = 1
        Else
            '6/10/10:  Remove the code that deleted the records without any demo data
            'blDemoExist = False
            'For ilDemo = 1 To 18 Step 1
            '    If tgDrfRec(llLoop).tDrf.lDemo(ilDemo) > 0 Then
                    blDemoExist = True
                    tgDrfRec(llLoop).iStatus = 0
            '        Exit For
            '    End If
            'Next ilDemo
        End If
        If blDemoExist Then
            ilFound = False
            llIndex = tgDrfRec(llLoop).lIndex
            If llIndex < UBound(tgAllDrf) Then
                If tgDrfRec(llLoop).lIndex > 0 Then
                    If tgAllDrf(llIndex).iStatus = -1 Then
                        ilFound = True
                        tgAllDrf(llIndex) = tgDrfRec(llLoop)
                    End If
                End If
            End If
            If Not ilFound Then
                If ilCheck Then
                    'For llIndex = LBound(tgAllDrf) To UBound(tgAllDrf) - 1 Step 1
                    For llIndex = imLBDrf To UBound(tgAllDrf) - 1 Step 1
                        If tgAllDrf(llIndex).iStatus = -1 Then
                            ilFound = True
                            tgAllDrf(llIndex) = tgDrfRec(llLoop)
                            Exit For
                        End If
                    Next llIndex
                End If
            End If
            If Not ilFound Then
                ilCheck = False
                tgAllDrf(UBound(tgAllDrf)) = tgDrfRec(llLoop)
                ReDim Preserve tgAllDrf(0 To UBound(tgAllDrf) + 1) As DRFREC
            End If
        End If
    Next llLoop
    ilCheck = True
    For llLoop = imLBDrf To UBound(tgDrfRec) - 1 Step 1
        llLink = tgDrfRec(llLoop).lLink
        Do While llLink <> -1
            If (tgLinkDrfRec(llLink).tDrf.lCode > 0) And (Not tgLinkDrfRec(llLink).iModel) Then
                blDemoExist = True
                tgLinkDrfRec(llLink).iStatus = 1
            Else
                '6/10/10:  Remove the code that deleted the records without any demo data
                'blDemoExist = False
                'For ilDemo = 1 To 18 Step 1
                '    If tgLinkDrfRec(llLink).tDrf.lDemo(ilDemo) > 0 Then
                        blDemoExist = True
                        tgLinkDrfRec(llLink).iStatus = 0
                '        Exit For
                '    End If
                'Next ilDemo
            End If
            If blDemoExist Then
                'Find hole (iStatus = -1 or Add at end)
                ilFound = False
                llIndex = tgLinkDrfRec(llLink).lIndex
                If llIndex < UBound(tgAllDrf) Then
                    If tgLinkDrfRec(llLink).lIndex > 0 Then
                        If tgAllDrf(llIndex).iStatus = -1 Then
                            ilFound = True
                            tgAllDrf(llIndex) = tgLinkDrfRec(llLink)
                        End If
                    End If
                End If
                If Not ilFound Then
                    If ilCheck Then
                        'For llIndex = LBound(tgAllDrf) To UBound(tgAllDrf) - 1 Step 1
                        For llIndex = imLBDrf To UBound(tgAllDrf) - 1 Step 1
                            If tgAllDrf(llIndex).iStatus = -1 Then
                                ilFound = True
                                tgAllDrf(llIndex) = tgLinkDrfRec(llLink)
                                Exit For
                            End If
                        Next llIndex
                    End If
                End If
                If Not ilFound Then
                    ilCheck = False
                    tgAllDrf(UBound(tgAllDrf)) = tgLinkDrfRec(llLink)
                    ReDim Preserve tgAllDrf(0 To UBound(tgAllDrf) + 1) As DRFREC
                End If
            End If
            llLink = tgLinkDrfRec(llLink).lLink
        Loop
    Next llLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer, ilClearFrame As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If lmRowNo = -1 Then
        Exit Sub
    End If
    
    If rbcDataType(0).Value Then 'Daypart
        If ilClearFrame Then
            pbcArrow.Visible = False
            lacFrame(0).Visible = False
        End If
        If (ilBoxNo < imLBDCtrls) Or (ilBoxNo > UBound(tmDCtrls)) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case DVEHICLEINDEX
                lbcVehicle.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcVehicle.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(1) = tmDCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case DACT1CODEINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    edcDropDown.Visible = False  'Set visibility
                    slStr = edcDropDown.Text
                    gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1CODEINDEX)
                    tmSaveShow(lmRowNo).sShow(DACT1CODEINDEX) = tmDCtrls(DACT1CODEINDEX).sShow
                    slStr = edcDropDown.Text
                    If Trim$(tmSaveShow(lmRowNo).sSave(DACT1CODEINDEX)) <> slStr Then
                        tmSaveShow(lmRowNo).sSave(DACT1CODEINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                    If Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) <> "" Then
                        If lmRowNo >= UBound(tmSaveShow) Then
                            imDrfChg = True
                            ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                            ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                    End If
                End If
                    
            Case DACT1SETTINGINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    plcACT1Settings.Visible = False
                    edcDropDown.Visible = False  'Set visibility
                    edcDropDown.Text = ""
                    If edcACT1SettingT.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "T"
                    If edcACT1SettingS.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "S"
                    If edcACT1SettingC.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "C"
                    If edcACT1SettingF.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "F"
                    slStr = edcDropDown.Text
                    gSetShow pbcDemo(0), slStr, tmDCtrls(DACT1SETTINGINDEX)
                    tmSaveShow(lmRowNo).sShow(DACT1SETTINGINDEX) = tmDCtrls(DACT1SETTINGINDEX).sShow
                    slStr = edcDropDown.Text
                    If Trim$(tmSaveShow(lmRowNo).sSave(DACT1SETTINGINDEX)) <> slStr Then
                        tmSaveShow(lmRowNo).sSave(DACT1SETTINGINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                    If Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) <> "" Then
                        If lmRowNo >= UBound(tmSaveShow) Then
                            imDrfChg = True
                            ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                            ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                    End If
                End If
                
            Case DDAYPARTINDEX
                lbcDaypart.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcDaypart.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcDaypart.List(lbcDaypart.ListIndex)
                End If
                gSetShow pbcDemo(0), slStr, tmDCtrls(DDAYPARTINDEX)
                tmSaveShow(lmRowNo).sShow(DDAYPARTINDEX) = tmDCtrls(DDAYPARTINDEX).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(DDAYPARTINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(DDAYPARTINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
            
            Case DGROUPINDEX
                If smSource <> "I" Then 'Standard Airtime mode
                    lbcSocEco.Visible = False
                    cmcDropDown.Visible = False
                    edcDropDown.Visible = False
                    If lbcSocEco.ListIndex <= 0 Then
                        slStr = ""
                    Else
                        slStr = lbcSocEco.List(lbcSocEco.ListIndex)
                    End If
                    gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                    tmSaveShow(lmRowNo).sShow(DGROUPINDEX) = tmDCtrls(DGROUPINDEX).sShow
                    If Trim$(tmSaveShow(lmRowNo).sSave(DAIRTIMEGRPNOINDEX)) <> slStr Then
                        tmSaveShow(lmRowNo).sSave(DAIRTIMEGRPNOINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                Else 'Podcast Impression mode
                    edcDropDown.Visible = False  'Set visibility
                    slStr = edcDropDown.Text
                    If tgSpf.sSAudData = "H" Then
                        gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
                    End If
                    If tgSpf.sSAudData = "N" Then
                        gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                    End If
                    If tgSpf.sSAudData = "U" Then
                        gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
                    End If
                    gSetShow pbcDemo(0), slStr, tmDCtrls(DGROUPINDEX)
                    tmSaveShow(lmRowNo).sShow(DIMPRESSIONSINDEX) = tmDCtrls(ilBoxNo).sShow
                    If gCompNumberStr(Trim$(tmSaveShow(lmRowNo).sSave(DIMPRESSIONSINDEX)), slStr) <> 0 Then
                        tmSaveShow(lmRowNo).sSave(DIMPRESSIONSINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                    If Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) <> "" Then
                        If lmRowNo >= UBound(tmSaveShow) Then
                            imDrfChg = True
                            ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                            ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                    End If
                End If
            
            Case DDEMOINDEX To DDEMOINDEX + 17
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If tgSpf.sSAudData = "H" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
                End If
                If tgSpf.sSAudData = "N" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                End If
                If tgSpf.sSAudData = "U" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
                End If
                gSetShow pbcDemo(0), slStr, tmDCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(DDEMOINDEX + ilBoxNo - DDEMOINDEX) = tmDCtrls(ilBoxNo).sShow
                slStr = edcDropDown.Text
                If gCompNumberStr(Trim$(tmSaveShow(lmRowNo).sSave(DDEMOINDEX + ilBoxNo - DDEMOINDEX)), slStr) <> 0 Then
                    tmSaveShow(lmRowNo).sSave(DDEMOINDEX + ilBoxNo - DDEMOINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
        End Select
    
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        pbcArrow.Visible = False
        lacFrame(2).Visible = False
        If (ilBoxNo < imLBXCtrls) Or (ilBoxNo > UBound(tmXCtrls)) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case XVEHICLEINDEX
                lbcVehicle.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcVehicle.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                gSetShow pbcDemo(2), slStr, tmXCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(XVEHICLEINDEX) = tmXCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case XACT1CODEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                gSetShow pbcDemo(2), slStr, tmXCtrls(XACT1CODEINDEX)
                tmSaveShow(lmRowNo).sShow(XACT1CODEINDEX) = tmXCtrls(XACT1CODEINDEX).sShow
                slStr = edcDropDown.Text
                If Trim$(tmSaveShow(lmRowNo).sSave(XACT1CODEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(XACT1CODEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
                
            Case XACT1SETTINGINDEX
                plcACT1Settings.Visible = False
                edcDropDown.Visible = False  'Set visibility
                edcDropDown.Text = ""
                If edcACT1SettingT.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "T"
                If edcACT1SettingS.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "S"
                If edcACT1SettingC.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "C"
                If edcACT1SettingF.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "F"
                slStr = edcDropDown.Text
                gSetShow pbcDemo(2), slStr, tmXCtrls(XACT1SETTINGINDEX)
                tmSaveShow(lmRowNo).sShow(XACT1SETTINGINDEX) = tmXCtrls(XACT1SETTINGINDEX).sShow
                slStr = edcDropDown.Text
                If Trim$(tmSaveShow(lmRowNo).sSave(XACT1SETTINGINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(XACT1SETTINGINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
            
            Case XTIMEINDEX To XTIMEINDEX + 1
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gValidTime(slStr) Then
                    If slStr <> "" Then
                        slStr = gFormatTime(slStr, "A", "1")
                    End If
                    gSetShow pbcDemo(2), slStr, tmXCtrls(ilBoxNo)
                    tmSaveShow(lmRowNo).sShow(XTIMEINDEX + ilBoxNo - XTIMEINDEX) = tmXCtrls(ilBoxNo).sShow
                    If Trim$(tmSaveShow(lmRowNo).sSave(XTIMEINDEX + ilBoxNo - XTIMEINDEX)) <> slStr Then
                        tmSaveShow(lmRowNo).sSave(XTIMEINDEX + ilBoxNo - XTIMEINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                Else
                    Beep
                    edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(XTIMEINDEX + ilBoxNo - XTIMEINDEX))
                End If
                
            Case XDAYSINDEX
                lbcDays.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcDays.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcDays.List(lbcDays.ListIndex)
                End If
                gSetShow pbcDemo(2), slStr, tmXCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(XDAYSINDEX) = tmXCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(XDAYSINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(XDAYSINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case XGROUPINDEX
                lbcSocEco.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcSocEco.ListIndex <= 0 Then
                    slStr = ""
                Else
                    slStr = lbcSocEco.List(lbcSocEco.ListIndex)
                End If
                gSetShow pbcDemo(2), slStr, tmXCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(XGROUPINDEX) = tmXCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(XGROUPNINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(XGROUPNINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case XDEMOINDEX To XDEMOINDEX + 17
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If tgSpf.sSAudData = "H" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
                End If
                If tgSpf.sSAudData = "N" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                End If
                If tgSpf.sSAudData = "U" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
                End If
                gSetShow pbcDemo(2), slStr, tmXCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(XDEMOINDEX + ilBoxNo - XDEMOINDEX) = tmXCtrls(ilBoxNo).sShow
                slStr = edcDropDown.Text
                If gCompNumberStr(Trim$(tmSaveShow(lmRowNo).sSave(XDEMOINDEX + ilBoxNo - XDEMOINDEX)), slStr) <> 0 Then
                    tmSaveShow(lmRowNo).sSave(XDEMOINDEX + ilBoxNo - XDEMOINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
        End Select
        
    ElseIf rbcDataType(2).Value Then 'Time
        pbcArrow.Visible = False
        lacFrame(0).Visible = False
        If (ilBoxNo < imLBTCtrls) Or (ilBoxNo > UBound(tmTCtrls)) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case TVEHICLEINDEX
                lbcVehicle.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcVehicle.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(TVEHICLEINDEX) = tmTCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(TVEHICLEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(TVEHICLEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
            
            Case TTIMEINDEX To TTIMEINDEX + 1
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gValidTime(slStr) Then
                    If slStr <> "" Then
                        slStr = gFormatTime(slStr, "A", "1")
                    End If
                    gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                    tmSaveShow(lmRowNo).sShow(TTIMEINDEX + ilBoxNo - TTIMEINDEX) = tmTCtrls(ilBoxNo).sShow
                    If Trim$(tmSaveShow(lmRowNo).sSave(TTIMEINDEX + ilBoxNo - TTIMEINDEX)) <> slStr Then
                        tmSaveShow(lmRowNo).sSave(TTIMEINDEX + ilBoxNo - TTIMEINDEX) = slStr
                        If lmRowNo < UBound(tmSaveShow) Then
                            imDrfChg = True
                        End If
                    End If
                Else
                    Beep
                    edcDropDown.Text = Trim$(tmSaveShow(lmRowNo).sSave(TTIMEINDEX + ilBoxNo - TTIMEINDEX))
                End If
                
            Case TDAYSINDEX
                lbcDays.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcDays.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcDays.List(lbcDays.ListIndex)
                End If
                gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(TDAYSINDEX) = tmTCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(TDAYSINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(TDAYSINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case TDEMOINDEX To TDEMOINDEX + 17
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If tgSpf.sSAudData = "H" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
                End If
                If tgSpf.sSAudData = "N" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                End If
                If tgSpf.sSAudData = "U" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
                End If
                gSetShow pbcDemo(0), slStr, tmTCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(TDEMOINDEX + ilBoxNo - TDEMOINDEX) = tmTCtrls(ilBoxNo).sShow
                slStr = edcDropDown.Text
                If gCompNumberStr(Trim$(tmSaveShow(lmRowNo).sSave(TDEMOINDEX + ilBoxNo - TDEMOINDEX)), slStr) <> 0 Then
                    tmSaveShow(lmRowNo).sSave(TDEMOINDEX + ilBoxNo - TDEMOINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(TVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
        End Select
        
    Else 'Vehicle
        pbcArrow.Visible = False
        lacFrame(1).Visible = False
        If (ilBoxNo < imLBVCtrls) Or (ilBoxNo > UBound(tmVCtrls)) Then
            Exit Sub
        End If
        Select Case ilBoxNo
            Case VVEHICLEINDEX
                lbcVehicle.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcVehicle.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(VVEHICLEINDEX) = tmVCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                                
            Case VACT1CODEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                gSetShow pbcDemo(1), slStr, tmVCtrls(VACT1CODEINDEX)
                tmSaveShow(lmRowNo).sShow(VACT1CODEINDEX) = tmVCtrls(VACT1CODEINDEX).sShow
                slStr = edcDropDown.Text
                If Trim$(tmSaveShow(lmRowNo).sSave(VACT1CODEINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(VACT1CODEINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
                
            Case VACT1SETTINGINDEX
                plcACT1Settings.Visible = False
                edcDropDown.Visible = False  'Set visibility
                edcDropDown.Text = ""
                If edcACT1SettingT.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "T"
                If edcACT1SettingS.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "S"
                If edcACT1SettingC.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "C"
                If edcACT1SettingF.Text = "Yes" Then edcDropDown.Text = edcDropDown.Text & "F"
                slStr = edcDropDown.Text
                gSetShow pbcDemo(1), slStr, tmVCtrls(VACT1SETTINGINDEX)
                tmSaveShow(lmRowNo).sShow(VACT1SETTINGINDEX) = tmVCtrls(VACT1SETTINGINDEX).sShow
                slStr = edcDropDown.Text
                If Trim$(tmSaveShow(lmRowNo).sSave(VACT1SETTINGINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(VACT1SETTINGINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
                               
            Case VDAYSINDEX
                lbcDays.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False
                If lbcDays.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcDays.List(lbcDays.ListIndex)
                End If
                gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(VDAYSINDEX) = tmVCtrls(ilBoxNo).sShow
                If Trim$(tmSaveShow(lmRowNo).sSave(VDAYSINDEX)) <> slStr Then
                    tmSaveShow(lmRowNo).sSave(VDAYSINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                
            Case VDEMOINDEX To VDEMOINDEX + 17
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If tgSpf.sSAudData = "H" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
                End If
                If tgSpf.sSAudData = "N" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                End If
                If tgSpf.sSAudData = "U" Then
                    gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
                End If
                gSetShow pbcDemo(1), slStr, tmVCtrls(ilBoxNo)
                tmSaveShow(lmRowNo).sShow(VDEMOINDEX + ilBoxNo - VDEMOINDEX) = tmVCtrls(ilBoxNo).sShow
                slStr = edcDropDown.Text
                If gCompNumberStr(Trim$(tmSaveShow(lmRowNo).sSave(VDEMOINDEX + ilBoxNo - VDEMOINDEX)), slStr) <> 0 Then
                    tmSaveShow(lmRowNo).sSave(VDEMOINDEX + ilBoxNo - VDEMOINDEX) = slStr
                    If lmRowNo < UBound(tmSaveShow) Then
                        imDrfChg = True
                    End If
                End If
                If Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)) <> "" Then
                    If lmRowNo >= UBound(tmSaveShow) Then
                        imDrfChg = True
                        ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                        ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                        mInitNewDrf True, UBound(tgDrfRec)
                    End If
                End If
        End Select
    End If
    mSetCommands
End Sub

Private Sub mShowDpf(ilFromSave As Integer)
    Dim ilRet As Integer
    Dim slDemoStr As String
    Dim slPopStr As String
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilUpper As Integer
    Dim llLoop As Long
    Dim llIndex As Long
    Dim llUpper As Long

    If smSource = "I" Then 'Podcast Impression mode
        Exit Sub
    End If
    mPlusSetShow imPlusBoxNo
    imPlusBoxNo = -1
    lmPlusRowNo = -1
    If (lmPlusRowNo = lmRowNo) And (Not ilFromSave) Then
        If lmRowNo = -1 Then
            pbcPlus.Cls
            ReDim tgDpfRec(0 To 1) As DPFREC
            mInitNewDpf
            If (tgSpf.sDemoEstAllowed = "Y") And (imDPorEst = 0) Then
                pbcDPorEst_MouseUp vbLeftButton, 0, 0, 0
            End If
        Else
            If (tgSpf.sDemoEstAllowed = "Y") And (imDPorEst = 1) Then
                pbcDPorEst_MouseUp vbLeftButton, 0, 0, 0
            End If
        End If
        Exit Sub
    End If
    lmPlusRowNo = lmRowNo
    For llLoop = imLBDpf To UBound(tgDpfRec) - 1 Step 1
        llIndex = tgDpfRec(llLoop).lIndex
        If llIndex > 0 Then
            tgAllDpf(llIndex) = tgDpfRec(llLoop)
        Else
            llUpper = UBound(tgAllDpf)
            tgAllDpf(llUpper) = tgDpfRec(llLoop)
            ReDim Preserve tgAllDpf(0 To llUpper + 1) As DPFREC
        End If
    Next llLoop
    pbcPlus.Cls
    ReDim tgDpfRec(0 To 1) As DPFREC
    mInitNewDpf

    If imCustomIndex > 0 Then
        Exit Sub
    End If
    If ilFromSave Then
        Exit Sub
    End If
    If (lmRowNo < vbcDemo.Value) Or (lmRowNo > vbcDemo.Value + vbcDemo.LargeChange) Then
        If (tgSpf.sDemoEstAllowed = "Y") And (imDPorEst = 0) Then
            pbcDPorEst_MouseUp vbLeftButton, 0, 0, 0
        End If
        Exit Sub
    End If
    If (tgDrfRec(lmRowNo).iStatus <> 1) And (Not tgDrfRec(lmRowNo).iModel) Then
        If tgDrfRec(lmRowNo).tDrf.lCode >= 0 Then   'Negative number indicates from base duplicate
            Exit Sub
        End If
    End If
    If (tgSpf.sDemoEstAllowed = "Y") And (imDPorEst = 1) Then
        pbcDPorEst_MouseUp vbLeftButton, 0, 0, 0
    End If
    ilFound = False
    For llLoop = imLBDpf To UBound(tgAllDpf) - 1 Step 1
        '8/14/18: When modeling, use the saved drfcode value
        If tgDrfRec(lmRowNo).iModel = True Then
            If tgDrfRec(lmRowNo).lModelDrfCode = tgAllDpf(llLoop).lDrfCode Then
                ilFound = True
                Exit For
            End If
        Else
            If tgDrfRec(lmRowNo).tDrf.lCode = tgAllDpf(llLoop).lDrfCode Then
                ilFound = True
                Exit For
            End If
        End If
    Next llLoop
    If Not ilFound Then
        For illoop = imLBDpf To UBound(tgDpfDel) - 1 Step 1
            If tgDrfRec(lmRowNo).tDrf.lCode = tgDpfDel(illoop).lDrfCode Then
                ilFound = True
                Exit For
            End If
        Next illoop
    End If
    If Not ilFound Then
        tmDpfSrchKey1.lDrfCode = tgDrfRec(lmRowNo).tDrf.lCode
        tmDpfSrchKey1.iMnfDemo = 0
        ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tgDrfRec(lmRowNo).tDrf.lCode)
            tmMnfSrchKey.iCode = tmDpf.iMnfDemo
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgSpf.sSAudData = "H" Then
                slDemoStr = gLongToStrDec(tmDpf.lDemo, 1)
                slPopStr = gLongToStrDec(tmDpf.lPop, 1)
            ElseIf tgSpf.sSAudData = "N" Then
                slDemoStr = gLongToStrDec(tmDpf.lDemo, 2)
                slPopStr = gLongToStrDec(tmDpf.lPop, 2)
            ElseIf tgSpf.sSAudData = "U" Then
                slDemoStr = gLongToStrDec(tmDpf.lDemo, 3)
                slPopStr = gLongToStrDec(tmDpf.lPop, 3)
            Else
                slDemoStr = Trim$(Str$(tmDpf.lDemo))
                slPopStr = Trim$(Str$(tmDpf.lPop))
            End If
            llUpper = UBound(tgAllDpf)
            tgAllDpf(llUpper).sKey = Trim$(tmMnf.sName)
            tgAllDpf(llUpper).iStatus = 1
            tgAllDpf(llUpper).lDpfCode = tmDpf.lCode
            tgAllDpf(llUpper).lDrfCode = tmDpf.lDrfCode
            tgAllDpf(llUpper).sDemo = slDemoStr
            tgAllDpf(llUpper).sPop = slPopStr
            tgAllDpf(llUpper).lIndex = llUpper
            tgAllDpf(llUpper).sSource = "D"
            tgAllDpf(llUpper).iRdfCode = 0
            ReDim Preserve tgAllDpf(0 To llUpper + 1) As DPFREC
            ilRet = btrGetNext(hmDpf, tmDpf, imDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If

    For llLoop = imLBDpf To UBound(tgAllDpf) - 1 Step 1
        '8/14/18: When modeling, use the saved drfcode value
        If tgDrfRec(lmRowNo).iModel = True Then
            If tgDrfRec(lmRowNo).lModelDrfCode = tgAllDpf(llLoop).lDrfCode Then
                tgDpfRec(UBound(tgDpfRec)) = tgAllDpf(llLoop)
                tgDpfRec(UBound(tgDpfRec)).lIndex = llLoop
                ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) + 1) As DPFREC
            End If
        Else
            If (tgDrfRec(lmRowNo).tDrf.lCode = tgAllDpf(llLoop).lDrfCode) And (tgAllDpf(llLoop).sSource <> "B") Then
                tgDpfRec(UBound(tgDpfRec)) = tgAllDpf(llLoop)
                tgDpfRec(UBound(tgDpfRec)).lIndex = llLoop
                ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) + 1) As DPFREC
            ElseIf (tgDrfRec(lmRowNo).tDrf.lCode = tgAllDpf(llLoop).lDrfCode) And (tgAllDpf(llLoop).sSource = "B") And (tgDrfRec(lmRowNo).tDrf.iRdfCode = tgAllDpf(llLoop).iRdfCode) Then
                tgDpfRec(UBound(tgDpfRec)) = tgAllDpf(llLoop)
                tgDpfRec(UBound(tgDpfRec)).lIndex = llLoop
                ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) + 1) As DPFREC
            End If
        End If
    Next llLoop
    mInitNewDpf
    If UBound(tgDpfRec) - 1 > 1 Then
        ReDim tmSortDpfRec(0 To UBound(tgDpfRec) - 1) As DPFREC
        For llLoop = 0 To UBound(tmSortDpfRec) Step 1
            tmSortDpfRec(llLoop) = tgDpfRec(llLoop + 1)
        Next llLoop
        ArraySortTyp fnAV(tmSortDpfRec(), 0), UBound(tmSortDpfRec), 0, LenB(tmSortDpfRec(0)), 0, LenB(tmSortDpfRec(0).sKey), 0
        For llLoop = UBound(tmSortDpfRec) To 0 Step -1
            tgDpfRec(llLoop + 1) = tmSortDpfRec(llLoop)
        Next llLoop
    End If
    mSetDpfScrollBar
    pbcPlus_Paint
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetFocus                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBSCtrls Or ilBoxNo > UBound(tmSCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case DATEINDEX
            If smSSave(DATEINDEX) = "" Then
                pbcCalendar.Visible = True
            End If
            edcSpecDropDown.SetFocus
        Case POPSRCEINDEX
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case QUALPOPSRCEINDEX
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case POPINDEX To POPINDEX + 17
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetShow                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Select Case ilBoxNo
        Case NAMEINDEX
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            If smSSave(NAMEINDEX) <> edcSpecDropDown.Text Then
                smSSave(NAMEINDEX) = edcSpecDropDown.Text
                imDnfChg = True
            End If
        Case DATEINDEX
            plcCalendar.Visible = False
            cmcSpecDropDown.Visible = False
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If gValidDate(slStr) Then
                If smSSave(DATEINDEX) <> slStr Then
                    imDnfChg = True
                End If
                smSSave(DATEINDEX) = slStr
                slStr = gFormatDate(slStr)
                gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            Else
                Beep
                edcSpecDropDown.Text = smSSave(DATEINDEX)
            End If
        Case POPSRCEINDEX
            edcSpecDropDown.Visible = False  'Set visibility
            cmcSpecDropDown.Visible = False
            lbcPopSrce.Visible = False
            If lbcPopSrce.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = edcSpecDropDown.Text
            End If
            gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            If smSSave(POPSRCDESCINDEX) <> edcSpecDropDown.Text Then
                smSSave(POPSRCDESCINDEX) = edcSpecDropDown.Text
                imDnfChg = True
            End If
        Case QUALPOPSRCEINDEX
            edcSpecDropDown.Visible = False  'Set visibility
            cmcSpecDropDown.Visible = False
            lbcPopSrce.Visible = False
            If lbcPopSrce.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = edcSpecDropDown.Text
            End If
            gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            If smSSave(QUALSRCDESCINDEX) <> edcSpecDropDown.Text Then
                smSSave(QUALSRCDESCINDEX) = edcSpecDropDown.Text
                imDnfChg = True
            End If
        Case POPINDEX To POPINDEX + 17
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If tgSpf.sSAudData = "H" Then
                gFormatStr slStr, FMTLEAVEBLANK, 1, slStr
            End If
            If tgSpf.sSAudData = "N" Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            End If
            If tgSpf.sSAudData = "U" Then
                gFormatStr slStr, FMTLEAVEBLANK, 3, slStr
            End If
            gSetShow pbcSpec, slStr, tmSCtrls(ilBoxNo)
            slStr = edcSpecDropDown.Text
            If gCompNumberStr(smSSave(POPINDEX + ilBoxNo - POPINDEX), slStr) <> 0 Then
                smSSave(POPINDEX + ilBoxNo - POPINDEX) = edcSpecDropDown.Text
                imPopChg = True
                mComputeTotalPop
                If pbcEst.Visible Then
                    pbcEst.Cls
                    pbcEst_Paint
                End If
            End If
    End Select
    mSetCommands
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    On Error Resume Next

    gObtainBooksForEstByUSA
    
    If bmResearchSaved Then
        'If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
            'ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
        'End If
    End If
    
    Screen.MousePointer = vbDefault
    igManUnload = YES
    igExitTraffic = 1
    Unload Research
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    If smSSave(NAMEINDEX) = "" Then
        mTestFields = NO
        Exit Function
    End If
    If smSSave(DATEINDEX) = "" Then
        mTestFields = NO
        Exit Function
    End If
    mTestFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mEstPlusFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestEstFields(tlEst As DEFREC) As Integer
    Dim ilRes As Integer
    Dim llEstDate As Long

    If Trim$(tlEst.sStartDate) = "" Then
        Beep
        ilRes = MsgBox("Start Date value must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imEstBoxNo = EDATEINDEX
        mTestEstFields = NO
        Exit Function
    End If
    If Not gValidDate(tlEst.sStartDate) Then
        Beep
        ilRes = MsgBox("Start Date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        imEstBoxNo = EDATEINDEX
        mTestEstFields = NO
        Exit Function
    End If
    llEstDate = gDateValue(tlEst.sStartDate)
    If llEstDate < gDateValue(smSSave(DATEINDEX)) Then
        Beep
        ilRes = MsgBox("Start Date must be after " & smSSave(DATEINDEX), vbOKOnly + vbExclamation, "Incomplete")
        imEstBoxNo = EDATEINDEX
        mTestEstFields = NO
        Exit Function
    End If
    If imEstByLOrU = 1 Then
        If Trim$(tlEst.sEstPct) = "" Then
            Beep
            ilRes = MsgBox("Percent value must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imEstBoxNo = EESTPCTINDEX
            mTestEstFields = NO
            Exit Function
        End If
    Else
        If Trim$(tlEst.sPop) = "" Then
            Beep
            ilRes = MsgBox("Population value must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imEstBoxNo = EPOPINDEX
            mTestEstFields = NO
            Exit Function
        End If
        If Trim$(tlEst.sEstPct) = "" Then
            Beep
            ilRes = MsgBox("Percent value must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imEstBoxNo = EESTPCTINDEX
            mTestEstFields = NO
            Exit Function
        End If
    End If
    mTestEstFields = YES
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestPlusFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestPlusFields(tlPlus As DPFREC) As Integer
'
'   iRet = mTestPlusFields(llRowNo)
'   Where:
'       llRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer

    If Trim$(tlPlus.sKey) = "" Then
        Beep
        ilRes = MsgBox("Demo must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPlusBoxNo = PDEMOINDEX
        mTestPlusFields = NO
        Exit Function
    End If
    If Trim$(tlPlus.sDemo) = "" Then
        Beep
        ilRes = MsgBox("Audience value must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPlusBoxNo = PDEMOINDEX
        mTestPlusFields = NO
        Exit Function
    End If
    If Trim$(tlPlus.sPop) = "" Then
        Beep
        ilRes = MsgBox("Population must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPlusBoxNo = PDEMOINDEX
        mTestPlusFields = NO
        Exit Function
    End If
    mTestPlusFields = YES
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(llRowNo As Long) As Integer
'
'   iRet = mTestSaveFields(llRowNo)
'   Where:
'       llRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If Trim$(tmSaveShow(llRowNo).sSave(1)) = "" Then
        Beep
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = DVEHICLEINDEX
        mTestSaveFields = NO
        Exit Function
    End If

    mTestSaveFields = YES
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSSaveFields                *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSSaveFields() As Integer
'
'   iRet = mTestSSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer
    If smSSave(NAMEINDEX) = "" Then
        Beep
        ilRes = MsgBox("Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSBoxNo = NAMEINDEX
        mTestSSaveFields = NO
        Exit Function
    End If
    If smSSave(DATEINDEX) = "" Then
        Beep
        ilRes = MsgBox("Book Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSBoxNo = DATEINDEX
        mTestSSaveFields = NO
        Exit Function
    End If
    If Not gValidDate(smSSave(DATEINDEX)) Then
        Beep
        ilRes = MsgBox("Book Date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        imSBoxNo = DATEINDEX
        mTestSSaveFields = NO
        Exit Function
    End If
    mTestSSaveFields = YES
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestTgFields                   *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestTgFields(llRowNo As Long) As Integer
'
'   iRet = mTestTgFields(llRowNo)
'   Where:
'       llRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim slMsg As String
    If (tgAllDrf(llRowNo).tDrf.sInfoType = "D") And (tgAllDrf(llRowNo).tDrf.iRdfCode <> 0) Then
        slMsg = "Sold Daypart"
    ElseIf (tgAllDrf(llRowNo).tDrf.sInfoType = "D") And (tgAllDrf(llRowNo).tDrf.iRdfCode = 0) Then
        slMsg = "Extra Daypart"
    ElseIf (tgAllDrf(llRowNo).tDrf.sInfoType = "T") Then
        slMsg = "Time"
    Else
        slMsg = "Vehicle"
    End If
    If tgAllDrf(llRowNo).tDrf.iVefCode <= 0 Then
        Beep
        ilRes = MsgBox("Vehicle must be specified in " & slMsg, vbOKOnly + vbExclamation, "Incomplete")
        mTestTgFields = NO
        Exit Function
    End If
    If (tgAllDrf(llRowNo).tDrf.sInfoType = "D") And (tgAllDrf(llRowNo).tDrf.iRdfCode = -1) Then
        Beep
        ilRes = MsgBox("Daypart must be specified in " & slMsg, vbOKOnly + vbExclamation, "Incomplete")
        mTestTgFields = NO
        Exit Function
    End If
    mTestTgFields = YES
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    If tgSaf(0).sAudByPackage <> "Y" Then
        ilRet = gPopUserVehicleBox(Research, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(Research, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSTDPKG + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Research
        On Error GoTo 0
    End If
    cbcVehicle.Clear
    For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilVef).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        cbcVehicle.AddItem lbcVehicle.List(ilVef)
        cbcVehicle.ItemData(cbcVehicle.NewIndex) = Val(slCode)
    Next ilVef
    cbcVehicle.AddItem "[All Vehicles]", 0
    cbcVehicle.ItemData(cbcVehicle.NewIndex) = 0
    cbcVehicle.ListIndex = 0
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imSBoxNo = DATEINDEX Then
                    edcSpecDropDown.Text = Format$(llDate, "m/d/yy")
                    edcSpecDropDown.SelStart = 0
                    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
                    imBypassFocus = True
                    edcSpecDropDown.SetFocus
                    Exit Sub
                ElseIf (imDPorEst = 1) And (imEstBoxNo = EDATEINDEX) Then
                    edcPlusDropDown.Text = Format$(llDate, "m/d/yy")
                    edcPlusDropDown.SelStart = 0
                    edcPlusDropDown.SelLength = Len(edcPlusDropDown.Text)
                    imBypassFocus = True
                    edcPlusDropDown.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imSBoxNo = DATEINDEX Then
        edcSpecDropDown.SetFocus
    ElseIf (imDPorEst = 1) And (imEstBoxNo = EDATEINDEX) Then
        edcPlusDropDown.SetFocus
    End If
End Sub

Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcDemo_GotFocus(Index As Integer)
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mPlusSetShow imPlusBoxNo
    imPlusBoxNo = -1
    lmPlusRowNo = -1
End Sub

Private Sub pbcDemo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxBox As Integer
    Dim llMaxRow As Long
    Dim llCompRow As Long
    Dim llRow As Long
    Dim llRowNo As Long
    
    Screen.MousePointer = vbDefault
    llCompRow = (vbcDemo.LargeChange + 1)
    If UBound(tgDrfRec) > llCompRow Then
        llMaxRow = llCompRow
    Else
        llMaxRow = UBound(tgDrfRec) ' + 1
    End If
    If rbcDataType(0).Value Then 'Daypart
        'If rbcDemoType(0).Value Then
        If imCustomIndex <= 0 Then
            ilMaxBox = UBound(tmDCtrls)
            If smSource = "I" Then 'Podcast Impression mode
                ilMaxBox = DGROUPINDEX
            End If
        Else
            ilMaxBox = UBound(smCustomDemo) + 4
        End If
        For llRow = 1 To llMaxRow Step 1
            For ilBox = imLBDCtrls To ilMaxBox Step 1
                If (X >= tmDCtrls(ilBox).fBoxX) And (X <= (tmDCtrls(ilBox).fBoxX + tmDCtrls(ilBox).fBoxW)) Then
                    If (Y >= (imRowOffset * (llRow - 1) * (fgBoxGridH + 15) + tmDCtrls(ilBox).fBoxY)) And (Y <= (imRowOffset * (llRow - 1) * (fgBoxGridH + 15) + tmDCtrls(ilBox).fBoxY + tmDCtrls(ilBox).fBoxH)) Then
                        llRowNo = llRow + vbcDemo.Value - 1
                        If llRowNo > UBound(tmSaveShow) Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                        If (imCustomIndex > 0) And (ilBox >= 4) Then
                            If (smCustomDemo(ilBox - 4) = "") Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If smDataForm <> "8" Then
                            If (ilBox = DDEMOINDEX + 8) Or (ilBox = DDEMOINDEX + 17) Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If lmRowNo <> llRow + vbcDemo.Value - 1 Then
                            mSetShow imBoxNo, True
                        Else
                            mSetShow imBoxNo, False
                        End If
                        lmRowNo = llRow + vbcDemo.Value - 1
                        If (lmRowNo = UBound(tmSaveShow)) And (Trim$(tmSaveShow(lmRowNo).sSave(DVEHICLEINDEX)) = "") Then
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                        pbcDemo_Paint 0
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next llRow
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        If imCustomIndex <= 0 Then
            ilMaxBox = UBound(tmXCtrls)
        Else
            ilMaxBox = UBound(smCustomDemo) + 5
        End If
        For llRow = 1 To llMaxRow Step 1
            For ilBox = imLBXCtrls To ilMaxBox Step 1
                If (X >= tmXCtrls(ilBox).fBoxX) And (X <= (tmXCtrls(ilBox).fBoxX + tmXCtrls(ilBox).fBoxW)) Then
                    If (Y >= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmXCtrls(ilBox).fBoxY)) And (Y <= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmXCtrls(ilBox).fBoxY + tmXCtrls(ilBox).fBoxH)) Then
                        llRowNo = llRow + vbcDemo.Value - 1
                        If llRowNo > UBound(tmSaveShow) Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                        If (imCustomIndex > 0) And (ilBox >= 5) Then
                            If (smCustomDemo(ilBox - 5) = "") Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If smDataForm <> "8" Then
                            If (ilBox = XDEMOINDEX + 8) Or (ilBox = XDEMOINDEX + 17) Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If lmRowNo <> llRow + vbcDemo.Value - 1 Then
                            mSetShow imBoxNo, True
                        Else
                            mSetShow imBoxNo, False
                        End If
                        lmRowNo = llRow + vbcDemo.Value - 1
                        If (lmRowNo = UBound(tmSaveShow)) And (Trim$(tmSaveShow(lmRowNo).sSave(XVEHICLEINDEX)) = "") Then
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                        pbcDemo_Paint 2
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next llRow
    ElseIf rbcDataType(2).Value Then 'Time
        If imCustomIndex <= 0 Then
            ilMaxBox = UBound(tmTCtrls)
        Else
            ilMaxBox = UBound(smCustomDemo) + 5
        End If
        For llRow = 1 To llMaxRow Step 1
            For ilBox = imLBTCtrls To ilMaxBox Step 1
                If (X >= tmTCtrls(ilBox).fBoxX) And (X <= (tmTCtrls(ilBox).fBoxX + tmTCtrls(ilBox).fBoxW)) Then
                    If (Y >= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmTCtrls(ilBox).fBoxY)) And (Y <= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmTCtrls(ilBox).fBoxY + tmTCtrls(ilBox).fBoxH)) Then
                        llRowNo = llRow + vbcDemo.Value - 1
                        If llRowNo > UBound(tmSaveShow) Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                        If (imCustomIndex > 0) And (ilBox >= 5) Then
                            If (smCustomDemo(ilBox - 5) = "") Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If smDataForm <> "8" Then
                            If (ilBox = TDEMOINDEX + 8) Or (ilBox = TDEMOINDEX + 17) Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If lmRowNo <> llRow + vbcDemo.Value - 1 Then
                            mSetShow imBoxNo, True
                        Else
                            mSetShow imBoxNo, False
                        End If
                        lmRowNo = llRow + vbcDemo.Value - 1
                        If (lmRowNo = UBound(tmSaveShow)) And (Trim$(tmSaveShow(lmRowNo).sSave(TVEHICLEINDEX)) = "") Then
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                        pbcDemo_Paint 0
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next llRow
    Else 'Vehicle
        If imCustomIndex <= 0 Then
            ilMaxBox = UBound(tmVCtrls)
        Else
            ilMaxBox = UBound(smCustomDemo) + 3
        End If
        For llRow = 1 To llMaxRow Step 1
            For ilBox = imLBVCtrls To ilMaxBox Step 1
                If (X >= tmVCtrls(ilBox).fBoxX) And (X <= (tmVCtrls(ilBox).fBoxX + tmVCtrls(ilBox).fBoxW)) Then
                    If (Y >= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmVCtrls(ilBox).fBoxY)) And (Y <= (2 * (llRow - 1) * (fgBoxGridH + 15) + tmVCtrls(ilBox).fBoxY + tmVCtrls(ilBox).fBoxH)) Then
                        llRowNo = llRow + vbcDemo.Value - 1
                        If llRowNo > UBound(tmSaveShow) Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                        If (imCustomIndex > 0) And (ilBox >= 3) Then
                            If (smCustomDemo(ilBox - 3) = "") Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If smDataForm <> "8" Then
                            If (ilBox = VDEMOINDEX + 8) Or (ilBox = VDEMOINDEX + 17) Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If lmRowNo <> llRow + vbcDemo.Value - 1 Then
                            mSetShow imBoxNo, True
                        Else
                            mSetShow imBoxNo, False
                        End If
                        lmRowNo = llRow + vbcDemo.Value - 1
                        If (lmRowNo = UBound(tmSaveShow)) And (Trim$(tmSaveShow(lmRowNo).sSave(VVEHICLEINDEX)) = "") Then
                            mInitNewDrf True, UBound(tgDrfRec)
                        End If
                        pbcDemo_Paint 1
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next llRow
    End If
    mSetFocus imBoxNo
End Sub

Private Sub pbcDemo_Paint(Index As Integer)
    Dim llColor As Long
    Dim illoop As Integer
    Dim llStartRow As Long
    Dim llEndRow As Long
    Dim llRow As Long
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilMaxBox As Integer
    
    pbcDemo(Index).Cls
    'DoEvents
    llColor = pbcDemo(Index).ForeColor
    If imCustomIndex <= 0 Then
        ilMaxBox = 17
    Else
        ilMaxBox = UBound(smCustomDemo)
    End If
    If smDataForm <> "8" Then
        ilBox = 7
    Else
        ilBox = 8
    End If
    
    'Paint Titles
    mPaintTitles Index
    
    llStartRow = vbcDemo.Value   '+ 1  'Top location
    llEndRow = vbcDemo.Value + vbcDemo.LargeChange + 1 ' + 1
    If llEndRow > UBound(tmSaveShow) Then
        If Trim$(tmSaveShow(UBound(tmSaveShow)).sSave(2)) <> "" Then
            llEndRow = UBound(tmSaveShow) 'include blank row as it might have data
        Else
            llEndRow = UBound(tmSaveShow) ' - 1
        End If
    End If
    llColor = pbcDemo(Index).ForeColor
    For llRow = llStartRow To llEndRow Step 1
        If llRow = UBound(tmSaveShow) Then
            pbcDemo(Index).ForeColor = DARKPURPLE
        Else
            pbcDemo(Index).ForeColor = llColor
        End If
        
        If rbcDataType(0).Value Then 'Daypart
            If imCustomIndex <= 0 Then
                ilMaxBox = UBound(tmDCtrls)
                If (smSource = "I") Then
                    ilMaxBox = DGROUPINDEX
                End If
            Else
                ilMaxBox = UBound(smCustomDemo) + 4
            End If
            For ilBox = imLBDCtrls To ilMaxBox Step 1
                If smSource <> "I" Then 'Standard Airtime mode
                    pbcDemo(Index).CurrentX = tmDCtrls(ilBox).fBoxX + fgBoxInsetX
                    pbcDemo(Index).CurrentY = tmDCtrls(ilBox).fBoxY + (imRowOffset * (llRow - llStartRow)) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    slStr = tmSaveShow(llRow).sShow(ilBox)
                    pbcDemo(Index).Print slStr
                    pbcDemo(Index).ForeColor = llColor
                Else 'Podcast Impression mode
                    If ilBox <> DACT1CODEINDEX And ilBox <> DACT1SETTINGINDEX Then
                        pbcDemo(Index).CurrentX = tmDCtrls(ilBox).fBoxX + fgBoxInsetX
                        pbcDemo(Index).CurrentY = tmDCtrls(ilBox).fBoxY + (imRowOffset * (llRow - llStartRow)) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                        If ilBox = DGROUPINDEX Then
                            slStr = tmSaveShow(llRow).sShow(DIMPRESSIONSINDEX)
                        Else
                            slStr = tmSaveShow(llRow).sShow(ilBox)
                        End If
                        pbcDemo(Index).Print slStr
                        pbcDemo(Index).ForeColor = llColor
                    End If
                End If
            Next ilBox
            
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            If imCustomIndex <= 0 Then
                ilMaxBox = UBound(tmXCtrls)
            Else
                ilMaxBox = UBound(smCustomDemo) + 5
            End If
            For ilBox = imLBXCtrls To ilMaxBox Step 1
                pbcDemo(Index).CurrentX = tmXCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcDemo(Index).CurrentY = tmXCtrls(ilBox).fBoxY + (2 * (llRow - llStartRow)) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = tmSaveShow(llRow).sShow(ilBox)
                pbcDemo(Index).Print slStr
                pbcDemo(Index).ForeColor = llColor
            Next ilBox
            
        ElseIf rbcDataType(2).Value Then 'Time
            If imCustomIndex <= 0 Then
                ilMaxBox = UBound(tmTCtrls)
            Else
                ilMaxBox = UBound(smCustomDemo) + 5
            End If
            For ilBox = imLBTCtrls To ilMaxBox Step 1
                pbcDemo(Index).CurrentX = tmTCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcDemo(Index).CurrentY = tmTCtrls(ilBox).fBoxY + (2 * (llRow - llStartRow)) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = tmSaveShow(llRow).sShow(ilBox)
                pbcDemo(Index).Print slStr
                pbcDemo(Index).ForeColor = llColor
            Next ilBox
            
        Else 'Vehicle
            If imCustomIndex <= 0 Then
                ilMaxBox = UBound(tmVCtrls)
            Else
                ilMaxBox = UBound(smCustomDemo) + 3
            End If
            For ilBox = imLBVCtrls To ilMaxBox Step 1
                pbcDemo(Index).CurrentX = tmVCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcDemo(Index).CurrentY = tmVCtrls(ilBox).fBoxY + (2 * (llRow - llStartRow)) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = tmSaveShow(llRow).sShow(ilBox)
                pbcDemo(Index).Print slStr
                pbcDemo(Index).ForeColor = llColor
            Next ilBox
        End If
        pbcDemo(Index).ForeColor = llColor
    Next llRow
End Sub

Private Sub pbcDPorEst_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    mSetShow imBoxNo, False
    imBoxNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub pbcDPorEst_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("E")) Or (KeyAscii = Asc("e")) Then
        imDPorEst = 1
        pbcPlus.Visible = False
        If imEstByLOrU = 1 Then
            pbcUSA.Visible = True
        Else
            pbcEst.Visible = True
        End If
        mSetDefScrollBar
        If imEstByLOrU = 1 Then
            pbcUSA.Cls
            pbcUSA_Paint
        Else
            pbcEst.Cls
            pbcEst_Paint
        End If
        pbcDPorEst.Cls
        pbcDPorEst_Paint
    ElseIf (KeyAscii = Asc("P")) Or (KeyAscii = Asc("p")) Then
        imDPorEst = 0
        If smSource <> "I" Then 'Standard Airtime mode
            pbcPlus.Visible = True
        End If
        pbcUSA.Visible = False
        pbcEst.Visible = False
        pbcDPorEst.Cls
        pbcDPorEst_Paint
        mShowDpf False
        mSetDpfScrollBar
    End If
    If KeyAscii = Asc(" ") Then
        imDPorEst = imDPorEst + 1
        If imDPorEst > 1 Then
            imDPorEst = 0
        End If
        pbcDPorEst.Cls
        pbcDPorEst_Paint
        If imDPorEst = 0 Then
            If smSource <> "I" Then 'Standard Airtime mode
                pbcPlus.Visible = True
            End If
            pbcEst.Visible = False
            pbcUSA.Visible = False
            mShowDpf False
            mSetDpfScrollBar
        Else
            pbcPlus.Visible = False
            If imEstByLOrU = 1 Then
                pbcUSA.Visible = True
            Else
                pbcEst.Visible = True
            End If
            mSetDefScrollBar
            If imEstByLOrU = 1 Then
                pbcUSA.Cls
                pbcUSA_Paint
            Else
                pbcEst.Cls
                pbcEst_Paint
            End If
        End If
    End If
End Sub

Private Sub pbcDPorEst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imDPorEst = imDPorEst + 1
    If imDPorEst > 1 Then
        imDPorEst = 0
    End If
    pbcDPorEst.Cls
    pbcDPorEst_Paint
    If imDPorEst = 0 Then
        If smSource <> "I" Then 'Standard Airtime mode
            pbcPlus.Visible = True
        End If
        pbcEst.Visible = False
        pbcUSA.Visible = False
        mShowDpf False
        mSetDpfScrollBar
    Else
        pbcPlus.Visible = False
        If imEstByLOrU = 1 Then
            pbcUSA.Visible = True
        Else
            pbcEst.Visible = True
        End If
        mSetDefScrollBar
        If imEstByLOrU = 1 Then
            pbcUSA.Cls
            pbcUSA_Paint
        Else
            pbcEst.Cls
            pbcEst_Paint
        End If
    End If
End Sub

Private Sub pbcDPorEst_Paint()
    pbcDPorEst.CurrentX = fgBoxInsetX \ 20
    pbcDPorEst.CurrentY = 0
    If imDPorEst = 0 Then
        pbcDPorEst.Print "Pre-defined Dayparts"
    Else
        If (smTotalPop = "") Or (imEstByLOrU = 1) Then
            pbcDPorEst.Print "Estimate Populations "
        Else
            pbcDPorEst.Print "Estimate Populations (" & smTotalPop & ")"
        End If
    End If
End Sub

Private Sub pbcEst_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, False
    imBoxNo = -1
End Sub

Private Sub pbcEst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llMaxRow As Long
    Dim llCompRow As Long
    Dim llRow As Long
    Dim llRowNo As Long
    Dim ilMaxBox As Integer

    If imSelectedIndex < 0 Then
        Exit Sub
    End If
    If Trim$(smSSave(DATEINDEX)) = "" Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    llCompRow = (vbcPlus.LargeChange + 1)
    If UBound(tgDefRec) > llCompRow Then
        llMaxRow = llCompRow
    Else
        llMaxRow = UBound(tgDefRec) ' + 1
    End If
    ilMaxBox = UBound(tmPCtrls)
    For llRow = 1 To llMaxRow Step 1
        For ilBox = imLBPCtrls To ilMaxBox Step 1
            If (X >= tmPCtrls(ilBox).fBoxX) And (X <= (tmPCtrls(ilBox).fBoxX + tmPCtrls(ilBox).fBoxW)) Then
                If (Y >= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY)) And (Y <= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY + tmPCtrls(ilBox).fBoxH)) Then
                    llRowNo = llRow + vbcPlus.Value - 1
                    If llRowNo > UBound(tgDefRec) Then
                        Beep
                        mEstSetFocus imEstBoxNo
                        Exit Sub
                    End If
                    mSSetShow imSBoxNo
                    imSBoxNo = -1
                    mSetShow imBoxNo, False
                    imBoxNo = -1
                    mEstSetShow imEstBoxNo
                    If ((ilBox = EPOPINDEX) Or (ilBox = EESTPCTINDEX)) And (Trim$(tgDefRec(llRowNo).sStartDate) = "") Then
                        Beep
                        mPlusSetFocus imEstBoxNo
                        Exit Sub
                    End If
                    lmEstRowNo = llRow + vbcPlus.Value - 1
                    If (lmEstRowNo = UBound(tgDefRec)) And (Trim$(tgDefRec(lmEstRowNo).sStartDate) = "") Then
                        mInitNewDef
                    End If
                    imEstBoxNo = ilBox
                    mEstEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next llRow
    mEstSetFocus imEstBoxNo
End Sub

Private Sub pbcEst_Paint()
    Dim llStartRow As Long
    Dim llEndRow As Long
    Dim llRow As Long
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slStr As String
    llStartRow = vbcPlus.Value   'Top location
    llEndRow = vbcPlus.Value + vbcPlus.LargeChange
    If llEndRow > UBound(tgDefRec) Then
        If Trim$(tgDefRec(UBound(tgDefRec)).sStartDate) <> "" Then
            llEndRow = UBound(tgDefRec)
        Else
            llEndRow = UBound(tgDefRec) - 1
        End If
    End If
    llColor = pbcEst.ForeColor
    For llRow = llStartRow To llEndRow Step 1
        For ilBox = imLBPCtrls To UBound(tmPCtrls) Step 1
            pbcEst.CurrentX = tmPCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcEst.CurrentY = tmPCtrls(ilBox).fBoxY + (llRow - llStartRow) * (fgBoxGridH + 15) - 30
            If llRow = UBound(tgDefRec) Then
                pbcEst.ForeColor = DARKPURPLE
            Else
                pbcEst.ForeColor = llColor
            End If
            If ilBox = EDATEINDEX Then
                pbcEst.Print Trim$(tgDefRec(llRow).sStartDate)
            ElseIf ilBox = EPOPINDEX Then
                pbcEst.Print Trim$(tgDefRec(llRow).sPop)
            Else
                If Trim$(tgDefRec(llRow).sEstPct) <> "" Then
                    slStr = gSubStr(tgDefRec(llRow).sEstPct, "100.00")
                Else
                    slStr = ""
                End If
                pbcEst.Print Trim$(slStr)
            End If
        Next ilBox
        pbcEst.ForeColor = llColor
    Next llRow
End Sub

Private Sub pbcEstByLorU_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    mSetShow imBoxNo, False
    imBoxNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub pbcEstByLorU_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("L")) Or (KeyAscii = Asc("l")) Then
        mClearEst 1
        pbcEstByLorU.Cls
        pbcEstByLorU_Paint
    ElseIf (KeyAscii = Asc("U")) Or (KeyAscii = Asc("u")) Then
        mClearEst 0
        pbcEstByLorU.Cls
        pbcEstByLorU_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imEstByLOrU = 0 Then
            mClearEst 1
        Else
            mClearEst 0
        End If
        pbcEstByLorU.Cls
        pbcEstByLorU_Paint
    End If

End Sub

Private Sub pbcEstByLorU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imEstByLOrU = 0 Then
        mClearEst 1
    Else
        mClearEst 0
    End If
    pbcEstByLorU.Cls
    pbcEstByLorU_Paint
End Sub

Private Sub pbcEstByLorU_Paint()
    If (tgSpf.sDemoEstAllowed = "Y") And (imDPorEst = 1) Then
        pbcEstByLorU.CurrentX = fgBoxInsetX \ 20
        pbcEstByLorU.CurrentY = 0
        If imEstByLOrU = 1 Then
            pbcEstByLorU.Print "USA"
            pbcUSA.Visible = True
            pbcEst.Visible = False
            pbcPlus.Visible = False
        Else
            pbcEstByLorU.Print "Listener"
            pbcEst.Visible = True
            pbcUSA.Visible = False
            pbcPlus.Visible = False
        End If
    Else
        pbcUSA.Visible = False
        pbcEst.Visible = False
        If smSource <> "I" Then 'Standard Airtime mode
            pbcPlus.Visible = True
        End If
    End If
End Sub

Private Sub pbcNewTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLoop                                                                                *
'******************************************************************************************
    pbcSpec.Enabled = True
    Dim ilRet As Integer
    Dim slSvPercentChg As String
    If imInNewTab Then
        Exit Sub
    End If
    If imSelectedIndex > 1 Then
        Exit Sub
    End If
    If imUpdateAllowed = False Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If UBound(tgDrfRec) = imLBDrf Then
        imInNewTab = True
        If smDataForm <> "8" Then
            igDnfModel = 16
        Else
            igDnfModel = 18
        End If
        igResearchModelMethod = 0 'Model
        RSModel.Show vbModal
        DoEvents
        If (igReturn = 1) And (igDnfModel) > 0 Then
            Screen.MousePointer = vbHourglass  'Wait
            pbcSpec.Cls
            
            slSvPercentChg = sgPercentChg
            ilRet = mReadRec(igDnfModel, True)
            mMoveRecToCtrl
            If smSource = "I" Then 'Podcast Impression mode
                sgPercentChg = slSvPercentChg
                mBuildRearchAdjustVehicles True
                mAdjustVehicleFields
                sgPercentChg = ""
            End If
            lmSDrfPopRecPos = 0
            lmCDrfPopRecPos = 0
            If smSource = "I" Then 'Podcast Impression mode
                mSetControls True
                rbcDataType(0).Value = True 'Daypart
            End If
            mInitSShow
            mInitShow
            pbcSpec_Paint
            If bgResearchByImpressions Then
                rbcDataType(0).Value = True 'Daypart
            End If
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
        ElseIf igReturn = 0 Then
            If bgResearchByImpressions Then
                smSource = "I"
                mSetControls True
                rbcDataType(0).Value = True 'Daypart
                pbcDemo_Paint 0
            End If
        End If
        mSetControls True
        Screen.MousePointer = vbDefault
    End If
    imInNewTab = False
    pbcSpecSTab.SetFocus
End Sub

Private Sub pbcPlus_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, False
    imBoxNo = -1
    If (imCustomIndex > 0) Or (lmRowNo = -1) Then
        cmcDone.SetFocus
        Exit Sub
    End If
    If tgDrfRec(lmRowNo).iStatus = 0 Then
        cmcDone.SetFocus
    End If
End Sub

Private Sub pbcPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llMaxRow As Long
    Dim llCompRow As Long
    Dim llRow As Long
    Dim llRowNo As Long
    Dim ilMaxBox As Integer
    If (imCustomIndex > 0) Or (lmRowNo = -1) Then
        cmcDone.SetFocus
        Exit Sub
    End If
    If tgDrfRec(lmRowNo).iStatus = 0 Then
        cmcDone.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    llCompRow = (vbcPlus.LargeChange + 1)
    If UBound(tgDpfRec) > llCompRow Then
        llMaxRow = llCompRow
    Else
        llMaxRow = UBound(tgDpfRec) ' + 1
    End If
    ilMaxBox = UBound(tmPCtrls)
    For llRow = 1 To llMaxRow Step 1
        For ilBox = imLBPCtrls To ilMaxBox Step 1
            If (X >= tmPCtrls(ilBox).fBoxX) And (X <= (tmPCtrls(ilBox).fBoxX + tmPCtrls(ilBox).fBoxW)) Then
                If (Y >= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY)) And (Y <= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY + tmPCtrls(ilBox).fBoxH)) Then
                    llRowNo = llRow + vbcPlus.Value - 1
                    If llRowNo > UBound(tgDpfRec) Then
                        Beep
                        mPlusSetFocus imPlusBoxNo
                        Exit Sub
                    End If
                    mSSetShow imSBoxNo
                    imSBoxNo = -1
                    mSetShow imBoxNo, False
                    imBoxNo = -1
                    mPlusSetShow imPlusBoxNo
                    If (imPlusBoxNo = PDEMOINDEX) Then
                        If mCheckPlusDemoName() Then
                            Beep
                            mPlusSetFocus imPlusBoxNo
                            Exit Sub
                        End If
                    End If
                    lmPlusRowNo = llRow + vbcPlus.Value - 1
                    If (lmPlusRowNo = UBound(tgDpfRec)) And (Trim$(tgDpfRec(lmPlusRowNo).sKey) = "") Then
                        mInitNewDpf
                    End If
                    imPlusBoxNo = ilBox
                    mPlusEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next llRow
    mPlusSetFocus imPlusBoxNo
End Sub

Private Sub pbcPlus_Paint()
    Dim llStartRow As Long
    Dim llEndRow As Long
    Dim llRow As Long
    Dim ilBox As Integer
    Dim llColor As Long
    llStartRow = vbcPlus.Value   'Top location
    llEndRow = vbcPlus.Value + vbcPlus.LargeChange
    If llEndRow > UBound(tgDpfRec) Then
        If Trim$(tgDpfRec(UBound(tgDpfRec)).sKey) <> "" Then
            llEndRow = UBound(tgDpfRec)
        Else
            llEndRow = UBound(tgDpfRec) - 1
        End If
    End If
    llColor = pbcPlus.ForeColor
    For llRow = llStartRow To llEndRow Step 1
        For ilBox = imLBPCtrls To UBound(tmPCtrls) Step 1
            pbcPlus.CurrentX = tmPCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPlus.CurrentY = tmPCtrls(ilBox).fBoxY + (llRow - llStartRow) * (fgBoxGridH + 15) - 30
            If llRow = UBound(tgDpfRec) Then
                pbcPlus.ForeColor = DARKPURPLE
            Else
                pbcPlus.ForeColor = llColor
            End If
            If ilBox = PDEMOINDEX Then
                pbcPlus.Print Trim$(tgDpfRec(llRow).sKey)
            ElseIf ilBox = PAUDINDEX Then
                pbcPlus.Print Trim$(tgDpfRec(llRow).sDemo)
            Else
                pbcPlus.Print Trim$(tgDpfRec(llRow).sPop)
            End If
        Next ilBox
        pbcPlus.ForeColor = llColor
    Next llRow
End Sub

Private Sub pbcPlusSTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcPlusSTab.HWnd Then
        Exit Sub
    End If
    If imDPorEst = 0 Then
        'If (rbcDemoType(1).Value) Or (lmRowNo = -1) Then
        If (imCustomIndex > 0) Or (lmRowNo = -1) Then
            cmcDone.SetFocus
            Exit Sub
        End If
        If tgDrfRec(lmRowNo).iStatus = 0 Then
            cmcDone.SetFocus
            Exit Sub
        End If
        imTabDirection = -1 'Set- Right to left
        mSSetShow imSBoxNo    'Remove focus
        imSBoxNo = -1
        Select Case imPlusBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imSettingValue = True
                vbcPlus.Value = 1
                imSettingValue = False
                lmPlusRowNo = 1
                ilBox = PDEMOINDEX
                imPlusBoxNo = ilBox
                mPlusEnableBox ilBox
                Exit Sub
            Case PDEMOINDEX 'Name (first control within header)
                mPlusSetShow imPlusBoxNo
                If lmPlusRowNo <= 1 Then
                    imPlusBoxNo = -1
                    lmPlusRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                Else
                    ilBox = PPOPINDEX
                    lmPlusRowNo = lmPlusRowNo - 1
                    If lmPlusRowNo < vbcPlus.Value Then
                        imSettingValue = True
                        vbcPlus.Value = vbcPlus.Value - 1
                        imSettingValue = False
                    End If
                    imPlusBoxNo = ilBox
                    mPlusEnableBox ilBox
                    Exit Sub
                End If
            Case Else
                ilBox = imPlusBoxNo - 1
        End Select
        mPlusSetShow imPlusBoxNo
        imPlusBoxNo = ilBox
        mPlusEnableBox ilBox
    Else
        imTabDirection = -1 'Set- Right to left
        mSSetShow imSBoxNo    'Remove focus
        imSBoxNo = -1
        Select Case imEstBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imSettingValue = True
                vbcPlus.Value = 1
                imSettingValue = False
                lmEstRowNo = 1
                ilBox = EDATEINDEX
                imEstBoxNo = ilBox
                mEstEnableBox ilBox
                Exit Sub
            Case EDATEINDEX 'Name (first control within header)
                mEstSetShow imEstBoxNo
                If lmEstRowNo <= 1 Then
                    imEstBoxNo = -1
                    lmEstRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                Else
                    If imEstByLOrU = 1 Then
                        ilBox = EESTPCTINDEX
                    Else
                        ilBox = EESTPCTINDEX    'EPOPINDEX
                    End If
                    lmEstRowNo = lmEstRowNo - 1
                    If lmEstRowNo < vbcPlus.Value Then
                        imSettingValue = True
                        vbcPlus.Value = vbcPlus.Value - 1
                        imSettingValue = False
                    End If
                    imEstBoxNo = ilBox
                    mEstEnableBox ilBox
                    Exit Sub
                End If
            Case Else
                ilBox = imEstBoxNo - 1
        End Select
        mEstSetShow imEstBoxNo
        imEstBoxNo = ilBox
        mEstEnableBox ilBox
    End If
End Sub

Private Sub pbcPlusTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcPlusTab.HWnd Then
        Exit Sub
    End If
    If imDPorEst = 0 Then
        If (imCustomIndex > 0) Or (lmRowNo = -1) Then
            cmcDone.SetFocus
            Exit Sub
        End If
        If tgDrfRec(lmRowNo).iStatus = 0 Then
            cmcDone.SetFocus
            Exit Sub
        End If
        imTabDirection = 0 'Set- Left to right
        Select Case imPlusBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                lmPlusRowNo = UBound(tgDpfRec) - 1
                imSettingValue = True
                If lmPlusRowNo <= vbcPlus.LargeChange + 1 Then
                    vbcPlus.Value = 1
                Else
                    vbcPlus.Value = lmPlusRowNo - vbcPlus.LargeChange - 1
                End If
                imSettingValue = False
                ilBox = PPOPINDEX
                imPlusBoxNo = ilBox
                mPlusEnableBox ilBox
                Exit Sub
            Case PDEMOINDEX
                'Test if row is blank, if so then exit
                mPlusSetShow imPlusBoxNo
                If (lmPlusRowNo = UBound(tgDpfRec)) And (Trim$(tgDpfRec(lmPlusRowNo).sKey) = "") Then
                    lmPlusRowNo = -1
                    If cmcSave.Enabled Then
                        cmcSave.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                If mCheckPlusDemoName() Then
                    Beep
                    ilBox = PDEMOINDEX
                Else
                    ilBox = PAUDINDEX
                End If
                imPlusBoxNo = ilBox
                mPlusEnableBox ilBox
                Exit Sub
            Case PPOPINDEX
                mPlusSetShow imPlusBoxNo
                If mTestPlusFields(tgDpfRec(lmPlusRowNo)) = NO Then
                    mPlusEnableBox imPlusBoxNo
                    Exit Sub
                End If
                If lmPlusRowNo >= UBound(tgDpfRec) Then
                    tgDpfRec(lmPlusRowNo).lDrfCode = tgDrfRec(lmRowNo).tDrf.lCode
                    imDpfChg = True
                    ReDim Preserve tgDpfRec(0 To lmPlusRowNo + 1) As DPFREC
                    mInitNewDpf
                    If UBound(tgDpfRec) <= vbcPlus.LargeChange Then 'was <=
                        vbcPlus.Max = imLBDpf   'LBound(tgDpfRec)
                    Else
                        vbcPlus.Max = UBound(tgDpfRec) - vbcPlus.LargeChange '-1
                    End If
                End If
                lmPlusRowNo = lmPlusRowNo + 1
                If lmPlusRowNo > vbcPlus.Value + vbcPlus.LargeChange Then ' + 1 Then
                    imSettingValue = True
                    vbcPlus.Value = vbcPlus.Value + 1
                    imSettingValue = False
                End If
                If lmPlusRowNo >= UBound(tgDpfRec) Then
                    mSetCommands
                    imPlusBoxNo = 0
                    lacPlus.Move 0, tmPCtrls(PDEMOINDEX).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15) - 30
                    lacPlus.Visible = True
                    pbcPArrow.Move pbcArrow.Left, plcPlus.Top + tmPCtrls(PDEMOINDEX).fBoxY + (lmPlusRowNo - vbcPlus.Value) * (fgBoxGridH + 15) + 45
                    pbcPArrow.Visible = True
                    pbcPArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = 1
                    imPlusBoxNo = ilBox
                    mPlusEnableBox ilBox
                    Exit Sub
                End If
            Case 0
                ilBox = imPlusBoxNo + 1
            Case Else
                ilBox = imPlusBoxNo + 1
        End Select
        mPlusSetShow imPlusBoxNo
        imPlusBoxNo = ilBox
        mPlusEnableBox ilBox
    Else
        imTabDirection = 0 'Set- Left to right
        Select Case imEstBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                lmEstRowNo = UBound(tgDefRec) - 1
                imSettingValue = True
                If lmEstRowNo <= vbcPlus.LargeChange + 1 Then
                    vbcPlus.Value = 1
                Else
                    vbcPlus.Value = lmEstRowNo - vbcPlus.LargeChange - 1
                End If
                imSettingValue = False
                ilBox = EESTPCTINDEX    'EPOPINDEX
                imEstBoxNo = ilBox
                mEstEnableBox ilBox
                Exit Sub
            Case EDATEINDEX
                'Test if row is blank, if so then exit
                mEstSetShow imEstBoxNo
                If (lmEstRowNo = UBound(tgDefRec)) And (Trim$(tgDefRec(lmEstRowNo).sStartDate) = "") Then
                    lmEstRowNo = -1
                    If cmcSave.Enabled Then
                        cmcSave.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                If imEstByLOrU = 1 Then
                    ilBox = EESTPCTINDEX
                Else
                    ilBox = EPOPINDEX
                End If
                imEstBoxNo = ilBox
                mEstEnableBox ilBox
                Exit Sub

            Case EESTPCTINDEX
                mEstSetShow imEstBoxNo
                If mTestEstFields(tgDefRec(lmEstRowNo)) = NO Then
                    mEstEnableBox imEstBoxNo
                    Exit Sub
                End If
                If lmEstRowNo >= UBound(tgDefRec) Then
                    imDefChg = True
                    'ReDim Preserve tgDefRec(1 To lmEstRowNo + 1) As DEFREC
                    ReDim Preserve tgDefRec(0 To lmEstRowNo + 1) As DEFREC
                    mInitNewDef
                    If UBound(tgDefRec) <= vbcPlus.LargeChange Then 'was <=
                        vbcPlus.Max = imLBDef   'LBound(tgDefRec)
                    Else
                        vbcPlus.Max = UBound(tgDefRec) - vbcPlus.LargeChange '-1
                    End If
                End If
                lmEstRowNo = lmEstRowNo + 1
                If lmEstRowNo > vbcPlus.Value + vbcPlus.LargeChange Then ' + 1 Then
                    imSettingValue = True
                    vbcPlus.Value = vbcPlus.Value + 1
                    imSettingValue = False
                End If
                If lmEstRowNo >= UBound(tgDefRec) Then
                    mSetCommands
                    imEstBoxNo = 0
                    lacEst.Move 0, tmPCtrls(EDATEINDEX).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15) - 30
                    lacEst.Visible = True
                    pbcPArrow.Move pbcArrow.Left, plcPlus.Top + tmPCtrls(EDATEINDEX).fBoxY + (lmEstRowNo - vbcPlus.Value) * (fgBoxGridH + 15) + 45
                    pbcPArrow.Visible = True
                    pbcPArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = 1
                    imEstBoxNo = ilBox
                    mEstEnableBox ilBox
                    Exit Sub
                End If
            Case 0
                ilBox = imEstBoxNo + 1
            Case Else
                ilBox = imEstBoxNo + 1
        End Select
        mEstSetShow imEstBoxNo
        imEstBoxNo = ilBox
        mEstEnableBox ilBox
    End If
End Sub

Private Sub pbcSpec_GotFocus()
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pbcSpec.Enabled = False Then Exit Sub
    Dim ilBox As Integer
    Dim ilMaxBox As Integer
    If (imCustomIndex <= 0) Then
        ilMaxBox = UBound(tmSCtrls)
    Else
        ilMaxBox = UBound(smCustomDemo) + 3
    End If
    For ilBox = imLBSCtrls To ilMaxBox Step 1
        If (X >= tmSCtrls(ilBox).fBoxX) And (X <= (tmSCtrls(ilBox).fBoxX + tmSCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmSCtrls(ilBox).fBoxY)) And (Y <= (tmSCtrls(ilBox).fBoxY + tmSCtrls(ilBox).fBoxH)) Then
                If smDataForm <> "8" Then
                    If (ilBox = POPINDEX + 8) Or (ilBox = POPINDEX + 17) Then
                        mSSetFocus imSBoxNo
                        Beep
                        Exit Sub
                    End If
                End If
                mSSetShow imSBoxNo
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub

Private Sub pbcSpec_Paint()
    Dim llColor As Long
    Dim illoop As Integer
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilMaxBox As Integer
    llColor = pbcSpec.ForeColor
    If (imCustomIndex <= 0) Then
        ilMaxBox = 17
    Else
        ilMaxBox = UBound(smCustomDemo)
    End If
    If smDataForm <> "8" Then
        ilBox = 7
    Else
        ilBox = 8
    End If
    mPaintSpecTitle
    If (imCustomIndex <= 0) Then
        ilMaxBox = UBound(tmSCtrls)
    Else
        ilMaxBox = UBound(tmSCtrls) 'UBound(smCustomDemo) + 3
    End If
    For ilBox = imLBSCtrls To ilMaxBox Step 1
        pbcSpec.CurrentX = tmSCtrls(ilBox).fBoxX + fgBoxInsetX
        If ilBox <= QUALPOPSRCEINDEX Then
            pbcSpec.CurrentY = tmSCtrls(ilBox).fBoxY + fgBoxInsetY
        Else
            pbcSpec.CurrentY = tmSCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        End If
        slStr = tmSCtrls(ilBox).sShow
        pbcSpec.Print slStr
    Next ilBox
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.HWnd Then
        Exit Sub
    End If
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    imTabDirection = -1 'Set- Right to left
    Select Case imSBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = NAMEINDEX
            imSBoxNo = ilBox
            mSEnableBox ilBox
            Exit Sub
        Case NAMEINDEX
            mSSetShow imSBoxNo
            imSBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case POPINDEX + 9
            If smDataForm <> "8" Then
                ilBox = imSBoxNo - 2
            Else
                ilBox = imSBoxNo - 1
            End If
        Case Else
            ilBox = imSBoxNo - 1
    End Select
    mSSetShow imSBoxNo
    imSBoxNo = ilBox
    mSEnableBox ilBox
End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    Dim ilMaxDemos As Integer
    If GetFocus() <> pbcSpecTab.HWnd Then
        Exit Sub
    End If
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
    If (imCustomIndex <= 0) Then
        ilMaxDemos = 17
        If smDataForm <> "8" Then
            ilMaxDemos = 16
        End If
    Else
        ilMaxDemos = UBound(smCustomDemo)
        Do While smCustomDemo(ilMaxDemos) = ""
            ilMaxDemos = ilMaxDemos - 1
            If ilMaxDemos = 1 Then
                Exit Do
            End If
        Loop
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imSBoxNo
        Case -1 'Tab from control prior to form area
            If smSource <> "I" Then 'Standard Airtime mode
                ilBox = POPINDEX
            Else 'Podcast Impression mode
                ilBox = DATEINDEX
            End If
            imSBoxNo = ilBox
            mSEnableBox ilBox
            Exit Sub
        Case DATEINDEX
            If smSource <> "I" Then 'Standard Airtime mode
                ilBox = imSBoxNo + 1
            Else 'Podcast Impression mode
                mSSetShow imSBoxNo
                imSBoxNo = -1
                imBoxNo = -1
                pbcSTab.SetFocus
                Exit Sub
            End If
        Case POPINDEX + 7
            If smDataForm <> "8" Then
                ilBox = imSBoxNo + 2
            Else
                ilBox = imSBoxNo + 1
            End If
        Case POPINDEX + ilMaxDemos
            mSSetShow imSBoxNo
            imSBoxNo = -1
            imBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imSBoxNo + 1
    End Select
    mSSetShow imSBoxNo
    imSBoxNo = ilBox
    mSEnableBox ilBox
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilTestIndex As Integer

    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    If rbcDataType(0).Value Then 'Daypart
        ilTestIndex = DDEMOINDEX + 9
        If smSource = "I" Then 'Podcast Impression mode
        ilTestIndex = DGROUPINDEX
        End If
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        ilTestIndex = XDEMOINDEX + 9
    ElseIf rbcDataType(2).Value Then 'Time
        ilTestIndex = TDEMOINDEX + 9
    Else 'Vehicle
        ilTestIndex = VDEMOINDEX + 9
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0 'Set- Left to right
            ilBox = DVEHICLEINDEX
            imBoxNo = ilBox
            lmRowNo = 1
            mEnableBox ilBox
            Exit Sub
        
        Case DVEHICLEINDEX, 0
            mSetShow imBoxNo, True
            If rbcDataType(0).Value Then 'Daypart
                ilBox = DDEMOINDEX
                If smSource = "I" Then 'Podcast Impression mode
                    ilBox = DGROUPINDEX
                End If
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                ilBox = XDEMOINDEX
            ElseIf rbcDataType(2).Value Then 'Time
                ilBox = TDEMOINDEX
            Else 'Vehicle
                ilBox = VDEMOINDEX
            End If
            If lmRowNo <= 1 Then
                imBoxNo = -1
                lmRowNo = -1
                mShowDpf False
                If pbcSpec.Enabled Then
                    pbcSpecTab.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            End If
            lmRowNo = lmRowNo - 1
            If lmRowNo < vbcDemo.Value Then
                imSettingValue = True
                vbcDemo.Value = vbcDemo.Value - 1   '- 2
                imSettingValue = False
            End If
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        
        Case DACT1CODEINDEX
            If smSource <> "I" Then 'Standard Airtime mode
                mSetShow imBoxNo, True
                ilBox = imBoxNo - 1
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
            
        Case DACT1SETTINGINDEX
            If smSource <> "I" Then 'Standard Airtime mode
                mSetShow imBoxNo, True
                ilBox = imBoxNo - 1
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
            
        Case ilTestIndex
            If rbcDataType(0).Value And (smSource = "I") Then 'Daypart
                ilBox = imBoxNo - 1
            Else
                If smDataForm <> "8" Then
                    ilBox = imBoxNo - 2
                Else
                    ilBox = imBoxNo - 1
                End If
            End If
        Case Else
            If smSource = "I" Then 'Podcast Impression mode
                If imBoxNo = 4 Then imBoxNo = 2 'Skip over the ACT1Code and Settings columns
                ilBox = imBoxNo - 1
            Else 'Standard Airtime mode
                ilBox = imBoxNo - 1
            End If
    End Select
    mSetShow imBoxNo, True
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub

Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilMax As Integer
    Dim ilMaxDemos As Integer
    Dim ilTestIndex As Integer

    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    'If rbcDemoType(0).Value Then
    If (imCustomIndex <= 0) Then
        ilMaxDemos = 17
        If smDataForm <> "8" Then
            ilMaxDemos = 16
        End If
    Else
        ilMaxDemos = UBound(smCustomDemo)
        Do While smCustomDemo(ilMaxDemos) = ""
            ilMaxDemos = ilMaxDemos - 1
            If ilMaxDemos = 1 Then
                Exit Do
            End If
        Loop
    End If
    If rbcDataType(0).Value Then 'Daypart or Podcast Impressions
        ilTestIndex = DDEMOINDEX + 7
        ilMax = DDEMOINDEX + ilMaxDemos
        If smSource = "I" Then 'Podcast Impression mode
            ilTestIndex = -2    'bypass testindex
            ilMax = DGROUPINDEX
        End If
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        ilTestIndex = XDEMOINDEX + 7
        ilMax = XDEMOINDEX + ilMaxDemos
    ElseIf rbcDataType(2).Value Then 'Time
        ilTestIndex = TDEMOINDEX + 7
        ilMax = TDEMOINDEX + ilMaxDemos
    Else 'Vehicle
        ilTestIndex = VDEMOINDEX + 7
        ilMax = VDEMOINDEX + ilMaxDemos
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            lmRowNo = UBound(tmSaveShow) - 1
            imSettingValue = True
            If lmRowNo <= vbcDemo.LargeChange + 1 Then
                vbcDemo.Value = 1
            Else
                vbcDemo.Value = lmRowNo - vbcDemo.LargeChange - 1
            End If
            imSettingValue = False
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
        Case ilTestIndex
            If smDataForm <> "8" Then
                ilBox = imBoxNo + 2
            Else
                ilBox = imBoxNo + 1
            End If
        Case ilMax
            mSetShow imBoxNo, True
            If mTestSaveFields(lmRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If lmRowNo >= UBound(tmSaveShow) Then
                imDrfChg = True
                ReDim Preserve tmSaveShow(0 To lmRowNo + 1) As SAVESHOW
                ReDim Preserve tgDrfRec(0 To UBound(tgDrfRec) + 1) As DRFREC
                mInitNewDrf True, UBound(tgDrfRec)
            End If
            If lmRowNo >= UBound(tmSaveShow) - 1 Then
                lmRowNo = lmRowNo + 1
                If UBound(tmSaveShow) <= vbcDemo.LargeChange + 1 Then ' + 1 Then
                    vbcDemo.Max = imLBSaveShow  'LBound(tmSaveShow)
                Else
                    vbcDemo.Max = UBound(tmSaveShow) - vbcDemo.LargeChange ' - 1
                End If
            Else
                lmRowNo = lmRowNo + 1
            End If
            If lmRowNo > vbcDemo.Value + vbcDemo.LargeChange Then
                imSettingValue = True
                If vbcDemo.Value + 1 > vbcDemo.Max Then
                    vbcDemo.Max = vbcDemo.Value + 1
                End If
                vbcDemo.Value = vbcDemo.Value + 1   '+ 2
                imSettingValue = False
            End If
            If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                pbcDemo_Paint 0
            ElseIf rbcDataType(1).Value Then 'Extra Daypart
                pbcDemo_Paint 2
            Else 'Vehicle
                pbcDemo_Paint 1
            End If
            If lmRowNo >= UBound(tmSaveShow) Then
                mShowDpf False
                imBoxNo = 0
                mSetCommands
                pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmDCtrls(DVEHICLEINDEX).fBoxY + imRowOffset * (lmRowNo - vbcDemo.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = DVEHICLEINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
        
            Exit Sub
        Case 0
            ilBox = DVEHICLEINDEX
        Case Else
            If smSource = "I" Then 'Podcast Impression mode
                If imBoxNo = 1 Then
                    mSetShow imBoxNo, True
                    imBoxNo = 3 'Skip over the ACT1Code and Settings columns
                End If
                ilBox = imBoxNo + 1
            Else
                ilBox = imBoxNo + 1
            End If
    End Select
    
    mSetShow imBoxNo, True
    imBoxNo = ilBox
    mEnableBox ilBox
    
End Sub

Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    Select Case imBoxNo
                        Case TTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            gSendKeys edcDropDown, slKey
                        Case TTIMEINDEX + 1
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            gSendKeys edcDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcUSA_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, False
    imBoxNo = -1
End Sub

Private Sub pbcUSA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llMaxRow As Long
    Dim llCompRow As Long
    Dim llRow As Long
    Dim llRowNo As Long
    Dim ilMaxBox As Integer

    If imSelectedIndex < 0 Then
        Exit Sub
    End If
    If Trim$(smSSave(DATEINDEX)) = "" Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    llCompRow = (vbcPlus.LargeChange + 1)
    If UBound(tgDefRec) > llCompRow Then
        llMaxRow = llCompRow
    Else
        llMaxRow = UBound(tgDefRec) ' + 1
    End If
    ilMaxBox = UBound(tmPCtrls)
    For llRow = 1 To llMaxRow Step 1
        For ilBox = imLBPCtrls To ilMaxBox Step 1
            If (X >= tmPCtrls(ilBox).fBoxX) And (X <= (tmPCtrls(ilBox).fBoxX + tmPCtrls(ilBox).fBoxW)) Then
                If (Y >= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY)) And (Y <= ((llRow - 1) * (fgBoxGridH + 15) + tmPCtrls(ilBox).fBoxY + tmPCtrls(ilBox).fBoxH)) Then
                    llRowNo = llRow + vbcPlus.Value - 1
                    If llRowNo > UBound(tgDefRec) Then
                        Beep
                        mEstSetFocus imEstBoxNo
                        Exit Sub
                    End If
                    mSSetShow imSBoxNo
                    imSBoxNo = -1
                    mSetShow imBoxNo, False
                    imBoxNo = -1
                    mEstSetShow imEstBoxNo
                    If imEstByLOrU = 1 Then
                        If (ilBox = EPOPINDEX) Then
                            ilBox = EESTPCTINDEX
                        End If
                        If (ilBox = EESTPCTINDEX) And (Trim$(tgDefRec(llRowNo).sStartDate) = "") Then
                            Beep
                            mPlusSetFocus imEstBoxNo
                            Exit Sub
                        End If
                    Else
                        If ((ilBox = EPOPINDEX) Or (ilBox = EESTPCTINDEX)) And (Trim$(tgDefRec(llRowNo).sStartDate) = "") Then
                            Beep
                            mPlusSetFocus imEstBoxNo
                            Exit Sub
                        End If
                    End If
                    lmEstRowNo = llRow + vbcPlus.Value - 1
                    If (lmEstRowNo = UBound(tgDefRec)) And (Trim$(tgDefRec(lmEstRowNo).sStartDate) = "") Then
                        mInitNewDef
                    End If
                    imEstBoxNo = ilBox
                    mEstEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next llRow
    mEstSetFocus imEstBoxNo
End Sub

Private Sub pbcUSA_Paint()
    Dim llStartRow As Long
    Dim llEndRow As Long
    Dim llRow As Long
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slStr As String
    llStartRow = vbcPlus.Value   'Top location
    llEndRow = vbcPlus.Value + vbcPlus.LargeChange
    If llEndRow > UBound(tgDefRec) Then
        If Trim$(tgDefRec(UBound(tgDefRec)).sStartDate) <> "" Then
            llEndRow = UBound(tgDefRec)
        Else
            llEndRow = UBound(tgDefRec) - 1
        End If
    End If
    llColor = pbcUSA.ForeColor
    For llRow = llStartRow To llEndRow Step 1
        For ilBox = imLBPCtrls To UBound(tmPCtrls) Step 1
            pbcUSA.CurrentX = tmPCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcUSA.CurrentY = tmPCtrls(ilBox).fBoxY + (llRow - llStartRow) * (fgBoxGridH + 15) - 30
            If llRow = UBound(tgDefRec) Then
                pbcUSA.ForeColor = DARKPURPLE
            Else
                pbcUSA.ForeColor = llColor
            End If
            If ilBox = EDATEINDEX Then
                pbcUSA.Print Trim$(tgDefRec(llRow).sStartDate)
            ElseIf ilBox = EPOPINDEX Then
                If Trim$(tgDefRec(llRow).sEstPct) <> "" Then
                    slStr = gSubStr(tgDefRec(llRow).sEstPct, "100.00")
                Else
                    slStr = ""
                End If
                pbcUSA.Print Trim$(slStr)
            End If
        Next ilBox
        pbcUSA.ForeColor = llColor
    Next llRow
End Sub

Private Sub plcDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub rbcDataType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    If bmIgnoreChg Then
        imDataType = Index
        Exit Sub
    End If
    Value = rbcDataType(Index).Value
    'End of coded added
    If Value Then
        Screen.MousePointer = vbHourglass
        mMoveCtrlToRec  'Move values from tgSaveShow to tgDrfRec, then tgDrfRec to tgAllDrf
        imDataType = Index
        mMoveRecToCtrl  'Move tgAllDrf to tgDrfRec, then tgDrfRec to tgSaveShow
        If Index = 0 Then   'Sold Daypart
            pbcDemo(0).Visible = True
            pbcDemo(1).Visible = False
            pbcDemo(2).Visible = False
        ElseIf Index = 1 Then   'Extra Daypart
            pbcDemo(2).Visible = True
            pbcDemo(0).Visible = False
            pbcDemo(1).Visible = False
        ElseIf Index = 2 Then   'Time
            pbcDemo(0).Visible = True
            pbcDemo(1).Visible = False
            pbcDemo(2).Visible = False
        Else    'Vehicle
            pbcDemo(1).Visible = True
            pbcDemo(0).Visible = False
            pbcDemo(2).Visible = False
        End If
        
        pbcSpec.Cls
        pbcDemo(0).Cls
        pbcDemo(1).Cls
        pbcDemo(2).Cls
        mInitSShow
        mInitShow
        pbcSpec_Paint
        If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
            pbcDemo_Paint 0
        ElseIf rbcDataType(1).Value Then 'Extra Daypart
            pbcDemo_Paint 2
        Else 'Vehicle
            pbcDemo_Paint 1
        End If
        Screen.MousePointer = vbDefault
    End If
    mSetCommands
End Sub

Private Sub rbcDataType_GotFocus(Index As Integer)
    lmPlusRowNo = -1
    'Save values, the one below will cause the array to be cleared
    mShowDpf False
    lmPlusRowNo = -1
    If imFirstFocus Then
        imFirstFocus = False
    End If
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mSetShow imBoxNo, True
    imBoxNo = -1
    lmRowNo = -1
    mShowDpf False
    mEstSetShow imEstBoxNo
    imEstBoxNo = -1
    lmEstRowNo = -1
End Sub

Private Sub rbcDataType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
    End Select
End Sub

Private Sub tmcDrag_Timer()
    Dim llCompRow As Long
    Dim llMaxRow As Long
    Dim llRow As Long
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            llCompRow = (vbcDemo.LargeChange + 2) \ 2
            If UBound(tmSaveShow) > llCompRow Then
                llMaxRow = llCompRow
            Else
                llMaxRow = UBound(tmSaveShow)
            End If
            For llRow = 1 To llMaxRow Step 1
                If (fmDragY >= ((llRow - 1) * (fgBoxGridH + 15) + tmDCtrls(DVEHICLEINDEX).fBoxY)) And (fmDragY <= ((llRow - 1) * (fgBoxGridH + 15) + tmDCtrls(DVEHICLEINDEX).fBoxY + tmDCtrls(DVEHICLEINDEX).fBoxH)) Then
                    mSetShow imBoxNo, True
                    imBoxNo = -1
                    lmRowNo = -1
                    mShowDpf False
                    mEstSetShow imEstBoxNo
                    imEstBoxNo = -1
                    lmEstRowNo = -1
                    lmRowNo = llRow + vbcDemo.Value - 1
                    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                        lacFrame(0).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacFrame(0).Move 0, tmDCtrls(DVEHICLEINDEX).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacFrame(0).Visible = True
                    ElseIf rbcDataType(1).Value Then 'Extra Daypart
                        lacFrame(2).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacFrame(2).Move 0, tmDCtrls(DVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacFrame(2).Visible = True
                    Else 'Vehicle
                        lacFrame(1).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacFrame(1).Move 0, tmVCtrls(DVEHICLEINDEX).fBoxY + (2 * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacFrame(1).Visible = True
                    End If
                    pbcArrow.Move pbcArrow.Left, plcDemo.Top + tmDCtrls(DVEHICLEINDEX).fBoxY + (imRowOffset * (lmRowNo - vbcDemo.Value)) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
                        lacFrame(0).Drag vbBeginDrag
                        lacFrame(0).DragIcon = IconTraf!imcIconDrag.DragIcon
                    ElseIf rbcDataType(1).Value Then 'Extra Daypart
                        lacFrame(2).Drag vbBeginDrag
                        lacFrame(2).DragIcon = IconTraf!imcIconDrag.DragIcon
                    Else 'Vehicle
                        lacFrame(1).Drag vbBeginDrag
                        lacFrame(1).DragIcon = IconTraf!imcIconDrag.DragIcon
                    End If
                    Exit Sub
                End If
            Next llRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub vbcDemo_Change()
    Dim ilIndex As Integer
    If rbcDataType(0).Value Or rbcDataType(2).Value Then 'Daypart or Time
        ilIndex = 0
    ElseIf rbcDataType(1).Value Then 'Extra Daypart
        ilIndex = 2
    Else 'Vehicle
        ilIndex = 1
    End If
    If imSettingValue Then
        pbcDemo(ilIndex).Cls
        pbcDemo_Paint ilIndex
        imSettingValue = False
    Else
        mSetShow imBoxNo, True
        pbcDemo(ilIndex).Cls
        pbcDemo_Paint ilIndex
        If (igWinStatus(RESEARCHLIST) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            mEnableBox imBoxNo
        End If
    End If
End Sub

Private Sub vbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Research"
End Sub

Private Sub vbcPlus_Change()
    If imSettingValue Then
        If imDPorEst = 0 Then
            pbcPlus.Cls
            pbcPlus_Paint
            imSettingValue = False
        Else
            If imEstByLOrU = 1 Then
                pbcUSA.Cls
                pbcUSA_Paint
            Else
                pbcEst.Cls
                pbcEst_Paint
            End If
            imSettingValue = False
        End If
    Else
        mSSetShow imSBoxNo
        imSBoxNo = -1
        mSetShow imBoxNo, False
        imBoxNo = -1
        If imDPorEst = 0 Then
            mPlusSetShow imPlusBoxNo
            pbcPlus.Cls
            pbcPlus_Paint
            If (igWinStatus(RESEARCHLIST) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                mPlusEnableBox imPlusBoxNo
            End If
        Else
            mEstSetShow imEstBoxNo
            If imEstByLOrU = 1 Then
                pbcUSA.Cls
                pbcUSA_Paint
            Else
                pbcEst.Cls
                pbcEst_Paint
            End If
            If (igWinStatus(RESEARCHLIST) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                mEstEnableBox imEstBoxNo
            End If
        End If
    End If
End Sub

Private Sub mMoveEstToRec(llRowNo As Long)
    tmDef.lCode = tgDefRec(llRowNo).lDefCode
    tmDef.iDnfCode = tgDnf.iCode
    gPackDate tgDefRec(llRowNo).sStartDate, tmDef.iStartDate(0), tmDef.iStartDate(1)
    If imEstByLOrU = 1 Then
        tmDef.lPopulation = 0
    Else
        If tgSpf.sSAudData = "H" Then
            tmDef.lPopulation = gStrDecToLong(tgDefRec(llRowNo).sPop, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            tmDef.lPopulation = gStrDecToLong(tgDefRec(llRowNo).sPop, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            tmDef.lPopulation = gStrDecToLong(tgDefRec(llRowNo).sPop, 3)
        Else
            tmDef.lPopulation = Val(tgDefRec(llRowNo).sPop)
        End If
    End If
    tmDef.lEstimatePct = gStrDecToLong(tgDefRec(llRowNo).sEstPct, 2)
End Sub

Private Sub mMovePlusToRec(llRowNo As Long)
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    tmDpf.lCode = tgAllDpf(llRowNo).lDpfCode
    tmDpf.iDnfCode = tgDnf.iCode
    gFindMatch Trim$(tgAllDpf(llRowNo).sKey), 0, lbcPlusDemos
    If gLastFound(lbcPlusDemos) >= 0 Then
        slNameCode = tmPlusDemoCode(gLastFound(lbcPlusDemos)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmDpf.iMnfDemo = Val(slCode)
    End If
    If tgSpf.sSAudData = "H" Then
        tmDpf.lDemo = gStrDecToLong(tgAllDpf(llRowNo).sDemo, 1)
        tmDpf.lPop = gStrDecToLong(tgAllDpf(llRowNo).sPop, 1)
    ElseIf tgSpf.sSAudData = "N" Then
        tmDpf.lDemo = gStrDecToLong(tgAllDpf(llRowNo).sDemo, 2)
        tmDpf.lPop = gStrDecToLong(tgAllDpf(llRowNo).sPop, 2)
    ElseIf tgSpf.sSAudData = "U" Then
        tmDpf.lDemo = gStrDecToLong(tgAllDpf(llRowNo).sDemo, 3)
        tmDpf.lPop = gStrDecToLong(tgAllDpf(llRowNo).sPop, 3)
    Else
        tmDpf.lDemo = Val(tgAllDpf(llRowNo).sDemo)
        tmDpf.lPop = Val(tgAllDpf(llRowNo).sPop)
    End If
    tmDpf.lDrfCode = tgAllDpf(llRowNo).lDrfCode
End Sub

Private Function mCheckPlusDemoName() As Integer
    Dim llRow As Long

    If lmPlusRowNo <= 0 Then
        mCheckPlusDemoName = False
        Exit Function
    End If
    For llRow = imLBDpf To UBound(tgDpfRec) - 1 Step 1
        If llRow <> lmPlusRowNo Then
            If Trim$(tgDpfRec(llRow).sKey) = Trim$(tgDpfRec(lmPlusRowNo).sKey) Then
                mCheckPlusDemoName = True
                Exit Function
            End If
        End If
    Next llRow
    'Allow 20 and 21 demo names within 16 buckets
    mCheckPlusDemoName = False
End Function

Private Function mEraseBook() As Integer
    Dim ilRet As Integer
    Dim slMsg As String
    Dim illoop As Integer
    Dim tlDrf As DRF
    Dim tlDef As DEF
    Dim slSQLQuery As String

    Screen.MousePointer = vbHourglass
    slSQLQuery = "Select Count(clfCode) as NumberBooks from DNF_Demo_Rsrch_Names "
    slSQLQuery = slSQLQuery & " Left Outer Join clf_Contract_Line On clfDnfCode = dnfCode"
    slSQLQuery = slSQLQuery & " Where dnfCode = " & tgDnf.iCode
    Set dnf_rst = gSQLSelectCall(slSQLQuery)
    If Not dnf_rst.EOF Then
        If Val(dnf_rst!NumberBooks) > 0 Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references the book name: " & Trim$(tgDnf.sBookName)
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            mEraseBook = False
            Exit Function
        End If
    End If
    
    ilRet = gIICodeRefExist(Research, tgDnf.iCode, "Vef.Btr", "VefDnfCode")
    If ilRet Then
        slMsg = ": Any Vehicles that defaults to this book will no longer have a default until replacement book is added.  Any new Proposal that are entered for vehicles with no default book will have no research.  Thus it is important to add a replacement research book for the one you are deleting as soon as possible. Ok to continue with the removal?"
    Else
        slMsg = ": Ok to remove?"
    End If
    Screen.MousePointer = vbDefault
    ilRet = MsgBox(Trim$(tgDnf.sBookName) & slMsg, vbOKCancel + vbQuestion, "Erase")
    If ilRet = vbCancel Then
        mEraseBook = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    'Remove defaults from vehicle
    For illoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(illoop).iDnfCode = tgDnf.iCode) Or (tgMVef(illoop).iReallDnfCode = tgDnf.iCode) Then
            tmVefSrchKey.iCode = tgMVef(illoop).iCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmVef.iDnfCode = 0
                tmVef.iReallDnfCode = 0
                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            End If
            tgMVef(illoop).iDnfCode = 0
            tgMVef(illoop).iReallDnfCode = 0
            '11/26/17
            gFileChgdUpdate "vef.btr", False
        End If
    Next illoop
    tmDnfSrchKey.iCode = tgDnf.iCode
    ilRet = btrGetEqual(hmDnf, tgDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        Screen.MousePointer = vbHourglass
        tmDrfSrchKey.iDnfCode = tgDnf.iCode
        tmDrfSrchKey.sDemoDataType = ""
        tmDrfSrchKey.iMnfSocEco = 0
        tmDrfSrchKey.iVefCode = 0
        tmDrfSrchKey.sInfoType = ""
        tmDrfSrchKey.iRdfCode = 0
        ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tgDnf.iCode)
            ilRet = btrDelete(hmDrf)
            If ilRet <> BTRV_ERR_NONE Then
                mEraseBook = False
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Function
            End If
            tmDpfSrchKey1.lDrfCode = tlDrf.lCode
            tmDpfSrchKey1.iMnfDemo = imP12PlusMnfCode 'TTP 10759 - Research List screen
            ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tlDrf.lCode)
                ilRet = btrDelete(hmDpf)
                If ilRet <> BTRV_ERR_NONE Then
                    mEraseBook = False
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Function
                End If
                tmDpfSrchKey1.lDrfCode = tlDrf.lCode
                tmDpfSrchKey1.iMnfDemo = imP12PlusMnfCode 'TTP 10759 - Research List screen
                ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Loop
            tmDrfSrchKey.iDnfCode = tgDnf.iCode
            tmDrfSrchKey.sDemoDataType = ""
            tmDrfSrchKey.iMnfSocEco = 0
            tmDrfSrchKey.iVefCode = 0
            tmDrfSrchKey.sInfoType = ""
            tmDrfSrchKey.iRdfCode = 0
            ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Loop
        tmDefSrchKey1.iDnfCode = tgDnf.iCode
        tmDefSrchKey1.iStartDate(0) = 0
        tmDefSrchKey1.iStartDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hmDef, tlDef, imDefRecLen, tmDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tlDef.iDnfCode = tgDnf.iCode)
            ilRet = btrDelete(hmDef)
            If ilRet <> BTRV_ERR_NONE Then
                mEraseBook = False
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Function
            End If
            tmDefSrchKey1.iDnfCode = tgDnf.iCode
            tmDefSrchKey1.iStartDate(0) = 0
            tmDefSrchKey1.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hmDef, tlDef, imDefRecLen, tmDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Loop
        ilRet = btrDelete(hmDnf)
        If ilRet <> BTRV_ERR_NONE Then
            mEraseBook = False
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Function
        End If
    End If
    mEraseBook = True
End Function

'11/15/11: Added the model parameter
Private Sub mGetDef(ilDnfCode As Integer, ilModel As Integer)
    Dim ilRet As Integer
    Dim slDate As String
    Dim slPopStr As String
    Dim llUpper As Long
    Dim llLoop As Long

    If tgSpf.sDemoEstAllowed <> "Y" Then
        Exit Sub
    End If
    tmDefSrchKey1.iDnfCode = ilDnfCode
    tmDefSrchKey1.iStartDate(0) = 0
    tmDefSrchKey1.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmDef, tmDef, imDefRecLen, tmDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmDef.iDnfCode = ilDnfCode)
        If tgSpf.sSAudData = "H" Then
            slPopStr = gLongToStrDec(tmDef.lPopulation, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            slPopStr = gLongToStrDec(tmDef.lPopulation, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            slPopStr = gLongToStrDec(tmDef.lPopulation, 3)
        Else
            slPopStr = Trim$(Str$(tmDef.lPopulation))
        End If
        gUnpackDateForSort tmDef.iStartDate(0), tmDef.iStartDate(1), slDate
        llUpper = UBound(tgDefRec)
        tgDefRec(llUpper).sKey = slDate
        '11/15/11: Add model test
        If ilModel Then
            tgDefRec(llUpper).iStatus = 0
            tgDefRec(llUpper).lDefCode = 0
        Else
            tgDefRec(llUpper).iStatus = 1
            tgDefRec(llUpper).lDefCode = tmDef.lCode
        End If
        gUnpackDate tmDef.iStartDate(0), tmDef.iStartDate(1), slDate
        tgDefRec(llUpper).sStartDate = slDate
        tgDefRec(llUpper).sPop = slPopStr
        tgDefRec(llUpper).sEstPct = gLongToStrDec(tmDef.lEstimatePct, 2)
        ReDim Preserve tgDefRec(0 To llUpper + 1) As DEFREC
        ilRet = btrGetNext(hmDef, tmDef, imDefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    pbcEst.Cls
    pbcUSA.Cls
    mInitNewDef
    If UBound(tgDefRec) - 1 > 1 Then
        ReDim tmSortDefRec(0 To UBound(tgDefRec) - 1) As DEFREC
        For llLoop = 0 To UBound(tmSortDefRec) Step 1
            tmSortDefRec(llLoop) = tgDefRec(llLoop + 1)
        Next llLoop
        ArraySortTyp fnAV(tmSortDefRec(), 0), UBound(tmSortDefRec), 0, LenB(tmSortDefRec(0)), 0, LenB(tmSortDefRec(0).sKey), 0
        For llLoop = UBound(tmSortDefRec) To 0 Step -1
            tgDefRec(llLoop + 1) = tmSortDefRec(llLoop)
        Next llLoop
    End If
    mSetDefScrollBar
    pbcEst_Paint
    pbcUSA_Paint
End Sub

Public Sub mSetDefScrollBar()
    imSettingValue = True
    vbcPlus.Min = imLBDef   'LBound(tgDefRec)
    imSettingValue = True
    If UBound(tgDefRec) <= vbcPlus.LargeChange + 1 Then
        vbcPlus.Max = imLBDef   'LBound(tgDefRec)
    Else
        vbcPlus.Max = UBound(tgDefRec) - vbcPlus.LargeChange
    End If
    imSettingValue = True
    vbcPlus.Value = vbcPlus.Min
    imSettingValue = False
End Sub

Private Sub mSetDpfScrollBar()
    imSettingValue = True
    vbcPlus.Min = imLBDpf   'LBound(tgDpfRec)
    imSettingValue = True
    If UBound(tgDpfRec) <= vbcPlus.LargeChange + 1 Then
        vbcPlus.Max = imLBDpf   'LBound(tgDpfRec)
    Else
        vbcPlus.Max = UBound(tgDpfRec) - vbcPlus.LargeChange
    End If
    imSettingValue = True
    vbcPlus.Value = vbcPlus.Min
    imSettingValue = False
End Sub

Private Sub mComputeTotalPop()
    Dim illoop As Integer
    Dim llRow As Long
    Dim slStr As String

    smTotalPop = 0
    For illoop = POPINDEX To POPINDEX + 17 Step 1
        smTotalPop = gAddStr(smTotalPop, smSSave(POPINDEX + illoop - POPINDEX))
    Next illoop
    If (smTotalPop = "") Or (Val(smTotalPop) = 0) Then
        smTotalPop = ""
    End If
    If (tgSpf.sDemoEstAllowed = "Y") And (imEstByLOrU <> 1) Then
        'Compute Percent
        For llRow = imLBDef To UBound(tgDefRec) - 1 Step 1
            If smTotalPop <> "" Then
                slStr = gMulStr(tgDefRec(llRow).sPop, "100.00")
                slStr = gDivStr(slStr, smTotalPop)
                slStr = gRoundStr(slStr, ".01", 2)
                If gCompNumberStr(tgDefRec(llRow).sEstPct, slStr) <> 0 Then
                    imDefChg = True
                    tgDefRec(llRow).sEstPct = slStr
                End If
            Else
                tgDefRec(llRow).sEstPct = ""
            End If
        Next llRow
        If imDPorEst = 1 Then
            pbcDPorEst.Cls
            pbcDPorEst_Paint
        End If
    End If
End Sub

Private Sub mClearEst(ilSwitch As Integer)
    Dim llRow As Long

    If imEstByLOrU = ilSwitch Then
        Exit Sub
    End If
    If imEstByLOrU = 1 Then
        For llRow = imLBDef To UBound(tgDefRec) - 1 Step 1
            tgDefRec(llRow).sPop = ""
            tgDefRec(llRow).sEstPct = ""
        Next llRow
        If imDPorEst = 1 Then
            pbcEst.Visible = True
            pbcUSA.Visible = False
        End If
    Else
        For llRow = imLBDef To UBound(tgDefRec) - 1 Step 1
            tgDefRec(llRow).sPop = ""
        Next llRow
        If imDPorEst = 1 Then
            pbcUSA.Visible = True
            pbcEst.Visible = False
        End If
    End If
    imEstByLOrU = ilSwitch
End Sub

Private Sub mGetDpf(ilDnfCode As Integer, ilModel As Integer)
    Dim ilRet As Integer
    Dim slDemoStr As String
    Dim slPopStr As String
    Dim llUpper As Long

    If smSource = "I" Then 'Podcast Impression mode
        Exit Sub
    End If
    tmDpfSrchKey2.iDnfCode = ilDnfCode
    tmDpfSrchKey2.iMnfDemo = 0
    ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.iDnfCode = ilDnfCode)
        tmMnfSrchKey.iCode = tmDpf.iMnfDemo
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If tgSpf.sSAudData = "H" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 1)
            slPopStr = gLongToStrDec(tmDpf.lPop, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 2)
            slPopStr = gLongToStrDec(tmDpf.lPop, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 3)
            slPopStr = gLongToStrDec(tmDpf.lPop, 3)
        Else
            slDemoStr = Trim$(Str$(tmDpf.lDemo))
            slPopStr = Trim$(Str$(tmDpf.lPop))
        End If
        llUpper = UBound(tgAllDpf)
        tgAllDpf(llUpper).sKey = Trim$(tmMnf.sName)
        If ilModel Then
            tgAllDpf(llUpper).iStatus = 0
            tgAllDpf(llUpper).lDpfCode = 0
        Else
            tgAllDpf(llUpper).iStatus = 1
            tgAllDpf(llUpper).lDpfCode = tmDpf.lCode
        End If
        tgAllDpf(llUpper).lDrfCode = tmDpf.lDrfCode
        tgAllDpf(llUpper).sDemo = slDemoStr
        tgAllDpf(llUpper).sPop = slPopStr
        tgAllDpf(llUpper).lIndex = llUpper
        If ilModel Then
            tgAllDpf(llUpper).sSource = "M"
        Else
            tgAllDpf(llUpper).sSource = "D"
        End If
        tgAllDpf(llUpper).iRdfCode = 0
        ReDim Preserve tgAllDpf(0 To llUpper + 1) As DPFREC
        ilRet = btrGetNext(hmDpf, tmDpf, imDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If ilModel Then
        imDpfChg = True
    End If
End Sub

Private Sub mDetermineUniqueGroups()
    Dim illoop As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer

    If imCustomIndex <= 0 Then
        ReDim smUniqueGroupDataTypes(0 To 0) As String
        Exit Sub
    End If
    ReDim smUniqueGroupDataTypes(0 To 1) As String
    smUniqueGroupDataTypes(0) = tmCustInfo(imCustomIndex - 1).sDataType
    For illoop = imCustomIndex To imCustomIndex + 16 Step 1
        If illoop < UBound(tmCustInfo) Then
            ilFound = False
            For ilTest = LBound(smUniqueGroupDataTypes) To UBound(smUniqueGroupDataTypes) - 1 Step 1
                If smUniqueGroupDataTypes(ilTest) = tmCustInfo(illoop).sDataType Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
            If Not ilFound Then
                smUniqueGroupDataTypes(UBound(smUniqueGroupDataTypes)) = tmCustInfo(illoop).sDataType
                ReDim Preserve smUniqueGroupDataTypes(0 To UBound(smUniqueGroupDataTypes) + 1) As String
            End If
        End If
    Next illoop
End Sub

Private Sub mMoveAllToDrfRec(llLoop As Long, llUpper As Long)
    Dim ilVef As Integer
    Dim slVefName As String
    Dim ilRdf As Integer
    Dim llTime As Long
    Dim slSTime As String
    Dim slETime As String
    Dim ilSDay As Integer
    Dim ilEDay As Integer
    Dim ilDay As Integer
    
    tgDrfRec(llUpper) = tgAllDrf(llLoop)
    tgDrfRec(llUpper).lIndex = llLoop
    tgDrfRec(llUpper).lLink = -1
    tgDrfRec(llUpper).iCustInfoIndex = 0
    tgAllDrf(llLoop).iStatus = -1   'Available record
    tgAllDrf(llLoop).iModel = False
    'Set vehicle key
    ilVef = gBinarySearchVef(tgDrfRec(llUpper).tDrf.iVefCode)
    If ilVef <> -1 Then
        slVefName = tgMVef(ilVef).sName
    Else
        slVefName = "~~~~~~~~~~"
    End If
    If (tgDrfRec(llUpper).tDrf.sInfoType = "D") And (tgDrfRec(llUpper).tDrf.iRdfCode <> 0) Then
        ilRdf = gBinarySearchRdf(tgDrfRec(llUpper).tDrf.iRdfCode)
        If ilRdf <> -1 Then
            tgDrfRec(llUpper).sKey = slVefName & tgMRdf(ilRdf).sName
        Else
            tgDrfRec(llUpper).sKey = slVefName & "~~~~~~~~~~"
        End If
    Else
        gUnpackTimeLong tgDrfRec(llUpper).tDrf.iStartTime(0), tgDrfRec(llUpper).tDrf.iStartTime(1), False, llTime
        slSTime = Trim(Str$(llTime))
        Do While Len(slSTime) < 6
            slSTime = "0" & slSTime
        Loop
        gUnpackTimeLong tgDrfRec(llUpper).tDrf.iEndTime(0), tgDrfRec(llUpper).tDrf.iEndTime(1), True, llTime
        slETime = Trim(Str$(llTime))
        Do While Len(slETime) < 6
            slETime = "0" & slETime
        Loop
        ilSDay = -1
        For ilDay = 0 To 6 Step 1
            If tgDrfRec(llUpper).tDrf.sDay(ilDay) = "Y" Then
                ilEDay = ilDay + 1
                If ilSDay = -1 Then
                    ilSDay = ilDay + 1
                End If
            End If
        Next ilDay
        tgDrfRec(llUpper).sKey = slVefName & slSTime & slETime & Trim$(Str$(ilSDay)) & Trim$(Str$(ilEDay))
    End If
    llUpper = llUpper + 1
End Sub

Private Sub mMoveAllToLinkDrfRec(llLoop As Long, llUpper As Long)
    Dim llLink As Long
    Dim slDataType As String
    Dim ilDemo As Integer
    
    mInitNewDrf False, llUpper
    llLink = tgDrfRec(llUpper).lLink
    slDataType = tgDrfRec(llUpper).tDrf.sDataType
    mMoveAllToDrfRec llLoop, llUpper
    tgDrfRec(llUpper - 1).lIndex = -1
    tgDrfRec(llUpper - 1).lLink = llLink
    tgDrfRec(llUpper - 1).tDrf.sDataType = slDataType
    tgDrfRec(llUpper - 1).tDrf.lCode = 0
    For ilDemo = 1 To 18
        tgDrfRec(llUpper - 1).tDrf.lDemo(ilDemo - 1) = 0
    Next ilDemo
    Do While llLink <> -1
        slDataType = tgLinkDrfRec(llLink).tDrf.sDataType
        tgLinkDrfRec(llLink).tDrf = tgDrfRec(llUpper - 1).tDrf
        tgLinkDrfRec(llLink).tDrf.lCode = 0
        tgLinkDrfRec(llLink).tDrf.lAutoCode = 0
        tgLinkDrfRec(llLink).tDrf.sDataType = slDataType
        tgLinkDrfRec(llLink).lRecPos = 0
        llLink = tgLinkDrfRec(llLink).lLink
    Loop
End Sub

Private Sub mAddPopIfReq()
    Dim illoop As Integer
    Dim ilPop As Integer
    Dim ilFound As Integer
    Dim ilGroup As Integer
    
    If imCustomIndex <= 0 Then
        Exit Sub
    End If
    For ilGroup = 0 To UBound(smUniqueGroupDataTypes) - 1 Step 1
        ilFound = False
        For ilPop = imLBDrf To UBound(tgCDrfPop) - 1 Step 1
            If tgCDrfPop(ilPop).sDataType = smUniqueGroupDataTypes(ilGroup) Then
                ilFound = True
                Exit For
            End If
        Next ilPop
        If Not ilFound Then
            ilPop = UBound(tgCDrfPop)
            tgCDrfPop(ilPop).lCode = 0
            tgCDrfPop(ilPop).sDataType = smUniqueGroupDataTypes(ilGroup)
            ReDim Preserve tgCDrfPop(0 To ilPop + 1) As DRF
        End If
    Next ilGroup
End Sub

Private Sub mSetModelFields()
    Dim llLoop As Long
    Dim llAdjustValue As Long
    Dim ilAud As Integer
    If Not bmModelUsed Then
        Exit Sub
    End If
    If sgPercentChg = "" Then
        llAdjustValue = 100
    ElseIf Val(sgPercentChg) = 0 Then
        llAdjustValue = 100
    Else
        llAdjustValue = Val(gAddStr("100", sgPercentChg))
    End If
    For llLoop = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If tgAllDrf(llLoop).iStatus >= 0 Then
            tgAllDrf(llLoop).iStatus = 0    'This status indicated that the field should be added
            tgAllDrf(llLoop).lRecPos = 0
            tgAllDrf(llLoop).iModel = True
            '8/14/18: retain the drfcode so that matching up with dpf
            If tgAllDrf(llLoop).tDrf.lCode <> 0 Then
                tgAllDrf(llLoop).lModelDrfCode = tgAllDrf(llLoop).tDrf.lCode
            End If
            '9/29/15: clear drfcode
            tgAllDrf(llLoop).tDrf.lCode = 0
            '2/7/19: Add adjustment
            For ilAud = LBound(tgAllDrf(llLoop).tDrf.lDemo) To UBound(tgAllDrf(llLoop).tDrf.lDemo) Step 1
                If tgAllDrf(llLoop).tDrf.sDemoDataType <> "P" Then
                    tgAllDrf(llLoop).tDrf.lDemo(ilAud) = gRoundStr(gDivStr(gMulStr(Str(llAdjustValue), Str$(tgAllDrf(llLoop).tDrf.lDemo(ilAud))), 100), "1", 0)
                End If
            Next ilAud
        End If
    Next llLoop
    For llLoop = imLBDpf To UBound(tgAllDpf) - 1 Step 1
        If tgAllDpf(llLoop).iStatus >= 0 Then
            tgAllDpf(llLoop).iStatus = 0
            tgAllDpf(llLoop).lDpfCode = 0
            '2/7/19: Add adjustment
            tgAllDpf(llLoop).sDemo = gRoundStr(gDivStr(gMulStr(Str(llAdjustValue), tgAllDpf(llLoop).sDemo), 100), "1", 0)
        End If
    Next llLoop
    For llLoop = imLBDef To UBound(tgDefRec) - 1 Step 1
        If tgDefRec(llLoop).iStatus >= 0 Then
            tgDefRec(llLoop).iStatus = 0
            tgDefRec(llLoop).lDefCode = 0
        End If
    Next llLoop
    For llLoop = imLBDrf To UBound(tgCDrfPop) - 1 Step 1
        tgCDrfPop(llLoop).lCode = 0
    Next llLoop
End Sub

Private Sub mAdjustVehicleFields()
    Dim llLoop As Long
    Dim llAdjustValue As Long
    Dim ilAud As Integer
    Dim ilVef As Integer
    Dim blAdjust As Boolean
    Dim llDrf As Long
    Dim slStr As String
    
    If sgPercentChg = "" Then
        llAdjustValue = 100
    ElseIf Val(sgPercentChg) = 0 Then
        llAdjustValue = 100
    Else
        llAdjustValue = Val(gAddStr("100", sgPercentChg))
    End If
    If llAdjustValue = 100 Then
        Exit Sub
    End If
    For llLoop = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If tgAllDrf(llLoop).iStatus >= 0 Then
            If tgAllDrf(llLoop).tDrf.iVefCode > 0 Then
                blAdjust = False
                For ilVef = 0 To UBound(tgResearchAdjustVehicle) - 1 Step 1
                    If tgResearchAdjustVehicle(ilVef).iVefCode = tgAllDrf(llLoop).tDrf.iVefCode Then
                        blAdjust = tgResearchAdjustVehicle(ilVef).bAdjust
                        Exit For
                    End If
                Next ilVef
                If blAdjust Then
                    For ilAud = LBound(tgAllDrf(llLoop).tDrf.lDemo) To UBound(tgAllDrf(llLoop).tDrf.lDemo) Step 1
                        If tgAllDrf(llLoop).tDrf.sDemoDataType <> "P" Then
                            imDrfChg = True
                            If tgSpf.sSAudData = "H" Then
                                slStr = gLongToStrDec(tgAllDrf(llLoop).tDrf.lDemo(ilAud), 1)
                            ElseIf tgSpf.sSAudData = "N" Then
                                slStr = gLongToStrDec(tgAllDrf(llLoop).tDrf.lDemo(ilAud), 2)
                            ElseIf tgSpf.sSAudData = "U" Then
                                slStr = gLongToStrDec(tgAllDrf(llLoop).tDrf.lDemo(ilAud), 3)
                            Else
                                slStr = Trim$(Str$(tgAllDrf(llLoop).tDrf.lDemo(ilAud)))
                            End If
                            
                            slStr = gDivStr(gMulStr(Str(llAdjustValue), slStr), "100.00")
                            If tgSpf.sSAudData = "H" Then
                                tgAllDrf(llLoop).tDrf.lDemo(ilAud) = gStrDecToLong(slStr, 1)
                            ElseIf tgSpf.sSAudData = "N" Then
                                tgAllDrf(llLoop).tDrf.lDemo(ilAud) = gStrDecToLong(slStr, 2)
                            ElseIf tgSpf.sSAudData = "U" Then
                                tgAllDrf(llLoop).tDrf.lDemo(ilAud) = gStrDecToLong(slStr, 3)
                            Else
                                tgAllDrf(llLoop).tDrf.lDemo(ilAud) = Val(slStr)
                            End If
                        End If
                    Next ilAud
                End If
            End If
        End If
    Next llLoop
    
    For llLoop = imLBDpf To UBound(tgAllDpf) - 1 Step 1
        If tgAllDpf(llLoop).iStatus >= 0 Then
            For llDrf = imLBDrf To UBound(tgAllDrf) - 1 Step 1
                If tgAllDpf(llLoop).lDrfCode = tgAllDrf(llDrf).tDrf.lCode Then
                    blAdjust = False
                    For ilVef = 0 To UBound(tgResearchAdjustVehicle) - 1 Step 1
                        If tgResearchAdjustVehicle(ilVef).iVefCode = tgAllDrf(llDrf).tDrf.iVefCode Then
                            blAdjust = tgResearchAdjustVehicle(ilVef).bAdjust
                            Exit For
                        End If
                    Next ilVef
                    If blAdjust Then
                        imDpfChg = True
                        If tgSpf.sSAudData = "H" Then
                            slStr = tgAllDpf(llLoop).sDemo  'gLongToStrDec(tgAllDpf(llLoop).sDemo, 1)
                        ElseIf tgSpf.sSAudData = "N" Then
                            slStr = tgAllDpf(llLoop).sDemo  'gLongToStrDec(tgAllDpf(llLoop).sDemo, 2)
                        ElseIf tgSpf.sSAudData = "U" Then
                            slStr = tgAllDpf(llLoop).sDemo  'gLongToStrDec(tgAllDpf(llLoop).sDemo, 3)
                        Else
                            slStr = tgAllDpf(llLoop).sDemo  'Trim$(str$(tgAllDpf(llLoop).sDemo))
                        End If
                        
                        slStr = gDivStr(gMulStr(Str(llAdjustValue), slStr), "100.00")
                        
                        tgAllDpf(llLoop).sDemo = slStr
                        If tgSpf.sSAudData = "H" Then
                            gFormatStr slStr, FMTLEAVEBLANK, 1, tgAllDpf(llLoop).sDemo
                        End If
                        If tgSpf.sSAudData = "N" Then
                            gFormatStr slStr, FMTLEAVEBLANK, 2, tgAllDpf(llLoop).sDemo
                        End If
                        If tgSpf.sSAudData = "U" Then
                            gFormatStr slStr, FMTLEAVEBLANK, 3, tgAllDpf(llLoop).sDemo
                        End If
                        
                    End If
                    Exit For
                End If
            Next llDrf
        End If
    Next llLoop
    
    If smSource = "I" Then 'Podcast Impression mode
        For llDrf = imLBDrf To UBound(tgAllDrf) - 1 Step 1
            blAdjust = False
            For ilVef = 0 To UBound(tgResearchAdjustVehicle) - 1 Step 1
                If tgResearchAdjustVehicle(ilVef).iVefCode = tgAllDrf(llDrf).tDrf.iVefCode Then
                    blAdjust = tgResearchAdjustVehicle(ilVef).bAdjust
                    Exit For
                End If
            Next ilVef
            If blAdjust Then
                imDpfChg = True
                slStr = tmSaveShow(llDrf).sSave(DIMPRESSIONSINDEX)
                slStr = gDivStr(gMulStr(Str(llAdjustValue), slStr), "100.00")
                tmSaveShow(llDrf).sSave(DIMPRESSIONSINDEX) = slStr
                gSetShow pbcDemo(0), slStr, tmDCtrls(DGROUPINDEX)
                tmSaveShow(llDrf).sShow(DIMPRESSIONSINDEX) = tmDCtrls(DGROUPINDEX).sShow
            End If
        Next llDrf
    End If
    mSetCommands
End Sub

Private Function mTestForDuplicateRows() As Boolean
    Dim blRet As Boolean
    Dim ilRes As Integer
    
    mTestForDuplicateRows = False
    blRet = mFindOrRemoveDuplicates(True)
    If blRet Then
        ilRes = MsgBox("Duplicated Specifications exist, Remove those that are Duplicate?", vbOKCancel, "Incomplete")
        If ilRes = vbCancel Then
            mTestForDuplicateRows = True
            Exit Function
        End If
        blRet = mFindOrRemoveDuplicates(False)
    End If
End Function

Private Sub mPaintTitles(Index As Integer)
    Dim llColor As Long
    Dim llFillColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilMaxBox As Integer
    Dim ilBox As Integer
    Dim illoop As Integer
    Dim ilCol As Integer
    Dim ilLineCount As Integer
    Dim llTop As Long
    Dim ilColorCount As Integer
    
    llColor = pbcDemo(Index).ForeColor
    slFontName = pbcDemo(Index).FontName
    flFontSize = pbcDemo(Index).FontSize
    pbcDemo(Index).ForeColor = BLUE
    pbcDemo(Index).FontBold = False
    pbcDemo(Index).FontSize = 7
    pbcDemo(Index).FontName = "Arial"
    pbcDemo(Index).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    'If rbcDemoType(0).Value Then
    If imCustomIndex <= 0 Then
        ilMaxBox = 17
    Else
        ilMaxBox = UBound(smCustomDemo)
    End If
    If smDataForm <> "8" Then
        ilBox = 7
    Else
        ilBox = 8
    End If
    
    '-------------------------------------------------------------------
    'Paint Data grid title boxes (Except for the Demo Boxes)
    llFillColor = pbcDemo(Index).FillColor
    pbcDemo(Index).FillColor = vbWhite
    ilCol = Switch(Index = 0, DDEMOINDEX, Index = 1, VDEMOINDEX, Index = 2, XDEMOINDEX)
    If rbcDataType(2).Value Then ilCol = TDEMOINDEX 'Time
    For illoop = 1 To ilCol - 1 Step 1
        Select Case imDataType
            Case 0 'Daypart
                If smSource <> "I" Then 'Standard Airtime mode
                    pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxY - 30), BLUE, B
                Else 'Podcast Impression mode
                    If illoop <> DACT1CODEINDEX And illoop <> DACT1SETTINGINDEX Then 'Skip drawing the Act1 code and setting columns in podcast Impression Mode
                        If illoop = DVEHICLEINDEX Then
                            'make the vehicle column Wider, since we dont have a Act1 Code and Setting Column
                            pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, 15)-Step(tmDCtrls(illoop).fBoxW + 15 + mAct1ColsWidth, tmDCtrls(illoop).fBoxY - 30), BLUE, B
                        Else
                            'Draw the rest of the columns normal
                            pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxY - 30), BLUE, B
                        End If
                    End If
                End If
            Case 1 'ExtraDaypart
                If illoop <> XTIMEINDEX And illoop <> XTIMEINDEX + 1 Then
                    pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, 15)-Step(tmXCtrls(illoop).fBoxW + 15, tmXCtrls(illoop).fBoxY - 30), BLUE, B
                Else
                    If illoop = XTIMEINDEX Then
                        pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, 15)-Step(tmXCtrls(illoop).fBoxW + tmXCtrls(illoop + 1).fBoxW + 30, tmXCtrls(illoop).fBoxY - 30), BLUE, B
                    End If
                End If
            Case 2 'Time
                If illoop <> TTIMEINDEX And illoop <> TTIMEINDEX + 1 Then
                    pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, 15)-Step(tmTCtrls(illoop).fBoxW + 15, tmXCtrls(illoop).fBoxY - 30), BLUE, B
                Else
                    If illoop = TTIMEINDEX Then
                        pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, 15)-Step(tmTCtrls(illoop).fBoxW + tmTCtrls(illoop + 1).fBoxW + 30, tmTCtrls(illoop).fBoxY - 30), BLUE, B
                    End If
                End If
            Case 3 'Vehicle
                pbcDemo(Index).Line (tmVCtrls(illoop).fBoxX - 15, 15)-Step(tmVCtrls(illoop).fBoxW + 15, tmVCtrls(illoop).fBoxY - 30), BLUE, B
        End Select
    Next illoop
    
    '-------------------------------------------------------------------
    'Draw Data Grid Demo Boxes
    If smSource <> "I" Then 'Standard Airtime mode
        For illoop = 0 To ilMaxBox Step 1
            If illoop <= 8 Then     'ilBox Then
                Select Case imDataType
                    Case 0 'Daypart
                        If tmDCtrls(DDEMOINDEX + illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmDCtrls(DDEMOINDEX + illoop).fBoxX - 15, 15)-Step(tmDCtrls(DDEMOINDEX + illoop).fBoxW + 15, tmDCtrls(DDEMOINDEX + illoop).fBoxY - 30), BLUE, B
                    Case 1 'ExtraDaypart
                        If tmXCtrls(XDEMOINDEX + illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmXCtrls(XDEMOINDEX + illoop).fBoxX - 15, 15)-Step(tmXCtrls(XDEMOINDEX + illoop).fBoxW + 15, tmXCtrls(XDEMOINDEX + illoop).fBoxY - 30), BLUE, B
                    Case 2 'Time
                        If tmTCtrls(TDEMOINDEX + illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmTCtrls(TDEMOINDEX + illoop).fBoxX - 15, 15)-Step(tmTCtrls(TDEMOINDEX + illoop).fBoxW + 15, tmTCtrls(TDEMOINDEX + illoop).fBoxY - 30), BLUE, B
                    Case 3 'Vehicle
                        If tmVCtrls(VDEMOINDEX + illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmVCtrls(VDEMOINDEX + illoop).fBoxX - 15, 15)-Step(tmVCtrls(VDEMOINDEX + illoop).fBoxW + 15, tmVCtrls(VDEMOINDEX + illoop).fBoxY - 30), BLUE, B
                End Select
                pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            Else
                pbcDemo(Index).CurrentY = 180 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            End If
            Select Case imDataType
                Case 0 'Daypart
                    pbcDemo(Index).CurrentX = tmDCtrls(DDEMOINDEX + illoop).fBoxX + 15 'fgBoxInsetX
                Case 1 'ExtraDaypart
                    pbcDemo(Index).CurrentX = tmVCtrls(VDEMOINDEX + illoop).fBoxX + 15 'fgBoxInsetX
                Case 2 'Time
                    pbcDemo(Index).CurrentX = tmTCtrls(TDEMOINDEX + illoop).fBoxX + 15 'fgBoxInsetX
                Case 3 'Vehicle
                    pbcDemo(Index).CurrentX = tmXCtrls(XDEMOINDEX + illoop).fBoxX + 15 'fgBoxInsetX
            End Select
            If imCustomIndex <= 0 Then
                pbcDemo(Index).Print smStdDemo(illoop)
            Else
                pbcDemo(Index).Print smCustomDemo(illoop)
            End If
        Next illoop
    End If
    
    '-------------------------------------------------------------------
    'Titles
    Select Case imDataType
        Case 0 'Daypart
            pbcDemo(Index).CurrentX = tmDCtrls(DVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Vehicle Name"
            If smSource <> "I" Then 'Standard Airtime mode
                pbcDemo(Index).CurrentX = tmDCtrls(DACT1CODEINDEX).fBoxX + 15  'fgBoxInsetX
                pbcDemo(Index).CurrentY = 30 - 15
                pbcDemo(Index).Print "ACT1 Code"
                pbcDemo(Index).CurrentX = tmDCtrls(DACT1SETTINGINDEX).fBoxX + 15  'fgBoxInsetX
                pbcDemo(Index).CurrentY = 30 - 15
                pbcDemo(Index).Print "Setting"
            End If
            pbcDemo(Index).CurrentX = tmDCtrls(DDAYPARTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Daypart"
            pbcDemo(Index).CurrentX = tmDCtrls(DGROUPINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            If smSource <> "I" Then 'Standard Airtime mode
                pbcDemo(Index).Print "Group #"
            Else 'Podcast Impression mode
                pbcDemo(Index).Print "Impressions"
            End If
        
        Case 1 'ExtraDaypart
            pbcDemo(Index).CurrentX = tmXCtrls(XVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Vehicle Name"
            pbcDemo(Index).CurrentX = tmXCtrls(XACT1CODEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "ACT1 Code"
            pbcDemo(Index).CurrentX = tmXCtrls(XACT1SETTINGINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Setting"
            pbcDemo(Index).CurrentX = tmXCtrls(XTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Time"
            pbcDemo(Index).CurrentX = tmXCtrls(XDAYSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcDemo(Index).Print "Days"
            pbcDemo(Index).CurrentX = tmXCtrls(XGROUPINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcDemo(Index).Print "Group #"
            
        Case 2 'Time
            pbcDemo(Index).CurrentX = tmTCtrls(TVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Vehicle Name"
            pbcDemo(Index).CurrentX = tmTCtrls(TTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Time"
            pbcDemo(Index).CurrentX = tmTCtrls(TDAYSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            If smSource <> "I" Then 'Standard Airtime mode
                pbcDemo(Index).Print "Days"
            Else 'Podcast Impression mode
                pbcDemo(Index).Print "Impressions"
            End If
        
        Case 3 'Vehicle
            pbcDemo(Index).CurrentX = tmVCtrls(VVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Vehicle Name"
            pbcDemo(Index).CurrentX = tmVCtrls(VACT1CODEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "ACT1 Code"
            pbcDemo(Index).CurrentX = tmVCtrls(VACT1SETTINGINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15
            pbcDemo(Index).Print "Setting"
            pbcDemo(Index).CurrentX = tmVCtrls(VDAYSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDemo(Index).CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcDemo(Index).Print "Days"
    End Select
    
    '-------------------------------------------------------------------
    'Paint Data Grid Lines
    ilLineCount = 0
    Select Case imDataType
        Case 0 'Daypart
            llTop = tmDCtrls(1).fBoxY
            Do
                For illoop = imLBDCtrls To UBound(tmDCtrls) Step 1
                    If illoop < DDEMOINDEX Then
                        If smSource <> "I" Then 'Standard Airtime mode
                            pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxH + 15), BLUE, B
                        Else 'Podcast Impression mode
                            If illoop <> DACT1CODEINDEX And illoop <> DACT1SETTINGINDEX Then
                                If illoop = DVEHICLEINDEX Then
                                    'Make the Vehicle column Wider since the Act1 columns are not shown in Podcast impression mode
                                    pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmDCtrls(illoop).fBoxW + 15 + mAct1ColsWidth, tmDCtrls(illoop).fBoxH + 15), BLUE, B
                                Else
                                    pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxH + 15), BLUE, B
                                End If
                            End If
                        End If
                    Else
                        If smSource <> "I" Then 'Standard Airtime mode
                            If ilColorCount <= 1 Then
                                If tmDCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxH + 15), BLUE, B
                            Else
                                pbcDemo(Index).FillColor = LIGHTERGREEN   'vbGreen
                                If tmDCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmDCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmDCtrls(illoop).fBoxW + 15, tmDCtrls(illoop).fBoxH + 15), BLUE, B
                                If ilColorCount = 3 Then ilColorCount = -1
                            End If
                        End If
                    End If
                Next illoop
                pbcDemo(Index).FillColor = vbWhite
                ilColorCount = ilColorCount + 1
                ilLineCount = ilLineCount + 1
                llTop = llTop + tmDCtrls(1).fBoxH + 15
            Loop While llTop + tmDCtrls(1).fBoxH < pbcDemo(Index).Height
        
        Case 1 'ExtraDaypart
            llTop = tmXCtrls(1).fBoxY
            Do
                For illoop = imLBXCtrls To UBound(tmXCtrls) Step 1
                    If illoop <> XTIMEINDEX And illoop <> XTIMEINDEX + 1 Then
                        If illoop < XDEMOINDEX Then
                            pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmXCtrls(illoop).fBoxW + 15, tmXCtrls(illoop).fBoxH + 15), BLUE, B
                        Else
                            If ilColorCount <= 1 Then
                                If tmXCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmXCtrls(illoop).fBoxW + 15, tmXCtrls(illoop).fBoxH + 15), BLUE, B
                            Else
                                pbcDemo(Index).FillColor = LIGHTERGREEN   'vbGreen
                                If tmXCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmXCtrls(illoop).fBoxW + 15, tmXCtrls(illoop).fBoxH + 15), BLUE, B
                                If ilColorCount = 3 Then ilColorCount = -1
                            End If
                        End If
                    Else
                        If illoop = XTIMEINDEX Then
                            pbcDemo(Index).Line (tmXCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmXCtrls(illoop).fBoxW + tmXCtrls(illoop + 1).fBoxW + 30, tmXCtrls(illoop).fBoxH + 15), BLUE, B
                        End If
                    End If
                Next illoop
                pbcDemo(Index).FillColor = vbWhite
                ilColorCount = ilColorCount + 1
                ilLineCount = ilLineCount + 1
                llTop = llTop + tmXCtrls(1).fBoxH + 15
            Loop While llTop + tmXCtrls(1).fBoxH < pbcDemo(Index).Height
        
        Case 2 'Time
            llTop = tmTCtrls(1).fBoxY
            Do
                For illoop = imLBTCtrls To UBound(tmTCtrls) Step 1
                    If illoop <> TTIMEINDEX And illoop <> TTIMEINDEX + 1 Then
                        If illoop < TDEMOINDEX Then
                            pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmTCtrls(illoop).fBoxW + 15, tmTCtrls(illoop).fBoxH + 15), BLUE, B
                        Else
                            If smSource <> "I" Then 'Standard Airtime mode
                                If ilColorCount <= 1 Then
                                    If tmTCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmTCtrls(illoop).fBoxW + 15, tmTCtrls(illoop).fBoxH + 15), BLUE, B
                                Else
                                    pbcDemo(Index).FillColor = LIGHTERGREEN   'vbGreen
                                    If tmTCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmTCtrls(illoop).fBoxW + 15, tmTCtrls(illoop).fBoxH + 15), BLUE, B
                                    If ilColorCount = 3 Then ilColorCount = -1
                                End If
                            End If
                        End If
                    Else
                        If illoop = TTIMEINDEX Then
                            pbcDemo(Index).Line (tmTCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmTCtrls(illoop).fBoxW + tmTCtrls(illoop + 1).fBoxW + 30, tmTCtrls(illoop).fBoxH + 15), BLUE, B
                        End If
                    End If
                
                Next illoop
                pbcDemo(Index).FillColor = vbWhite
                ilColorCount = ilColorCount + 1
                ilLineCount = ilLineCount + 1
                llTop = llTop + tmTCtrls(1).fBoxH + 15
            Loop While llTop + tmTCtrls(1).fBoxH < pbcDemo(Index).Height
            
        Case 3 'Vehicle
            llTop = tmVCtrls(1).fBoxY
            Do
                For illoop = imLBVCtrls To UBound(tmVCtrls) Step 1
                    If illoop < VDEMOINDEX Then
                        pbcDemo(Index).Line (tmVCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmVCtrls(illoop).fBoxW + 15, tmVCtrls(illoop).fBoxH + 15), BLUE, B
                    Else
                        If ilColorCount <= 1 Then
                            If tmVCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmVCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmVCtrls(illoop).fBoxW + 15, tmVCtrls(illoop).fBoxH + 15), BLUE, B
                        Else
                            pbcDemo(Index).FillColor = LIGHTERGREEN  'vbGreen
                            If tmVCtrls(illoop).fBoxX <> 0 Then pbcDemo(Index).Line (tmVCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmVCtrls(illoop).fBoxW + 15, tmVCtrls(illoop).fBoxH + 15), BLUE, B
                            If ilColorCount = 3 Then ilColorCount = -1
                        End If
                    End If
                Next illoop
                pbcDemo(Index).FillColor = vbWhite
                ilColorCount = ilColorCount + 1
                ilLineCount = ilLineCount + 1
                llTop = llTop + tmVCtrls(1).fBoxH + 15
            Loop While llTop + tmVCtrls(1).fBoxH < pbcDemo(Index).Height
    End Select
    pbcDemo(Index).FontSize = flFontSize
    pbcDemo(Index).FontName = slFontName
    pbcDemo(Index).FontSize = flFontSize
    pbcDemo(Index).ForeColor = llColor
    pbcDemo(Index).FontBold = True
    pbcDemo(Index).FillColor = llFillColor
End Sub

Private Sub mPaintSpecTitle()
    Dim llColor As Long
    Dim llFillColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilMaxBox As Integer
    Dim ilBox As Integer
    Dim illoop As Integer
    Dim ilCol As Integer
    Dim ilLineCount As Integer
    Dim llTop As Long
    
    llFillColor = pbcSpec.FillColor
    pbcSpec.FillColor = vbWhite
    llColor = pbcSpec.ForeColor
    slFontName = pbcSpec.FontName
    flFontSize = pbcSpec.FontSize
    pbcSpec.ForeColor = BLUE
    pbcSpec.FontBold = False
    pbcSpec.FontSize = 7
    pbcSpec.FontName = "Arial"
    pbcSpec.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If imCustomIndex <= 0 Then
        ilMaxBox = 17
    Else
        ilMaxBox = UBound(smCustomDemo)
    End If
    If smDataForm <> "8" Then
        ilBox = 7
    Else
        ilBox = 8
    End If
    '-------------------------------------------------------------------
    'Paint Title lines
    pbcSpec.FillColor = vbWhite
    For illoop = 0 To QUALPOPSRCEINDEX Step 1
        If illoop <= DATEINDEX Then
            If tmSCtrls(illoop).fBoxX <> 0 Then pbcSpec.Line (tmSCtrls(illoop).fBoxX - 15, 15)-Step(tmSCtrls(illoop).fBoxW + 15, tmSCtrls(illoop).fBoxH + 15), BLUE, B
        Else
            If tmSCtrls(illoop).fBoxX <> 0 Then pbcSpec.Line (tmSCtrls(illoop).fBoxX - 15, tmSCtrls(illoop).fBoxY - 15)-Step(tmSCtrls(illoop).fBoxW + 15, tmSCtrls(illoop).fBoxY - 30), BLUE, B
        End If
    Next illoop
    '-------------------------------------------------------------------
    'Paint Demo lines
    For illoop = 0 To ilMaxBox Step 1
        If illoop <= 8 Then 'ilBox Then
            pbcSpec.Line (tmSCtrls(POPINDEX + illoop).fBoxX - 15, 15)-Step(tmSCtrls(POPINDEX + illoop).fBoxW + 15, tmSCtrls(POPINDEX + illoop).fBoxY - 30), BLUE, B
            pbcSpec.CurrentY = 30 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Else
            pbcSpec.CurrentY = 180 - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        End If
        pbcSpec.CurrentX = tmSCtrls(POPINDEX + illoop).fBoxX + 15 'fgBoxInsetX
        If imCustomIndex <= 0 Then
            pbcSpec.Print smStdDemo(illoop)
        Else
            pbcSpec.Print smCustomDemo(illoop)
        End If
    Next illoop
    '-------------------------------------------------------------------
    'Paint Titles
    pbcSpec.CurrentX = tmSCtrls(NAMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSpec.CurrentY = 30 - 15
    pbcSpec.Print "Name"
    pbcSpec.CurrentX = tmSCtrls(DATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSpec.CurrentY = 30 - 15
    pbcSpec.Print "Book Date"
    pbcSpec.CurrentX = tmSCtrls(POPSRCEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSpec.CurrentY = tmSCtrls(POPSRCEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcSpec.Print "Population Source"
    pbcSpec.CurrentX = tmSCtrls(QUALPOPSRCEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSpec.CurrentY = tmSCtrls(QUALPOPSRCEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcSpec.Print "Qual Pop Source"
    ilLineCount = 0
    llTop = tmSCtrls(POPSRCEINDEX).fBoxY
    Do
        For illoop = POPINDEX To UBound(tmSCtrls) Step 1
            If tmSCtrls(illoop).fBoxX <> 0 Then pbcSpec.Line (tmSCtrls(illoop).fBoxX - 15, llTop - 15)-Step(tmSCtrls(illoop).fBoxW + 15, tmSCtrls(illoop).fBoxH + 15), BLUE, B
        Next illoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + fgBoxGridH + 15
    Loop While ilLineCount <= 2 'llTop + tmSCtrls(1).fBoxH < pbcSpec.Height
    pbcSpec.CurrentX = tmSCtrls(POPINDEX).fBoxX - pbcSpec.TextWidth("Population") - 60 'fgBoxInsetX
    pbcSpec.CurrentY = 30 - 15
    pbcSpec.Print "Population"
    pbcSpec.CurrentX = tmSCtrls(POPINDEX).fBoxX - pbcSpec.TextWidth("Population") - 60   'fgBoxInsetX
    pbcSpec.CurrentY = 180 - 15
    pbcSpec.Print "(000)"
    pbcSpec.FontSize = flFontSize
    pbcSpec.FontName = slFontName
    pbcSpec.FontSize = flFontSize
    pbcSpec.ForeColor = llColor
    pbcSpec.FontBold = True
    pbcSpec.FillColor = llFillColor
End Sub

Private Sub mApplyFilter()
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim llDate As Long
    Dim blOk As Boolean
    Dim slSQLQuery As String
    
    cbcSelect.Clear
    CSI_ComboBoxMS1.Clear
    lbcPopSrce.Clear
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
        lbcPopSrce.AddItem Trim$(dnf_rst!dnfBookName)
        lbcPopSrce.ItemData(lbcPopSrce.NewIndex) = dnf_rst!dnfCode
        blOk = True
        If imFilterVefCode > 0 Then
            slSQLQuery = "Select drfVefCode from DRF_Demo_Rsrch_Data where drfDnfCode = " & dnf_rst!dnfCode
            slSQLQuery = slSQLQuery & " And drfVefCode = " & imFilterVefCode
            '3/27/20: bypass population records and bad records
            slSQLQuery = slSQLQuery & " And drfDemoDataType <> 'P'"
            slSQLQuery = slSQLQuery & " And drfDemoDataType <> ''"
            Set drf_rst = gSQLSelectCall(slSQLQuery)
            If drf_rst.EOF Then
                blOk = False
            End If
        End If
        If blOk Then
            cbcSelect.AddItem Trim$(dnf_rst!dnfBookName) & ": " & dnf_rst!dnfBookDate
            cbcSelect.ItemData(cbcSelect.NewIndex) = dnf_rst!dnfCode
            CSI_ComboBoxMS1.AddItem Trim$(dnf_rst!dnfBookName) & ": " & dnf_rst!dnfBookDate
        End If
        dnf_rst.MoveNext
    Loop
    cbcSelect.AddItem "[New with Demo 18-24]", 0
    CSI_ComboBoxMS1.AddItem "[New with Demo 18-24]"
    cbcSelect.AddItem "[New with Demo 18-20 + 21-24]", 0
    CSI_ComboBoxMS1.AddItem "[New with Demo 18-20 + 21-24]"
    lbcPopSrce.AddItem "[This Book]", 0
    imChgMode = True
    cbcSelect.ListIndex = 0
    imChgMode = False
    If cbcSelect.ListCount < 20 Then
        gSetComboboxDropdownHeight Research, cbcSelect, cbcSelect.ListCount
    Else
        gSetComboboxDropdownHeight Research, cbcSelect, 20
    End If
End Sub

Private Function mFindOrRemoveDuplicates(blFindDuplicates As Boolean) As Boolean
    Dim llRowNo As Long
    Dim llCheckRow As Long
    Dim ilDay As Integer
    Dim blMatch As Boolean
    Dim ilRes As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim slVehicleName As String
    Dim slDayPartName As String
    Dim slDays As String
    Dim llStartTime1 As Long
    Dim llEndTime1 As Long
    Dim llStartTime2 As Long
    Dim llEndTime2 As Long
    Dim ilRdfCode As Integer
    Dim ilRdfCount As Integer
    Dim llExtraRowNo As Long
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    ReDim slDay(0 To 6) As String
    Dim tlRdf As RDF

    mFindOrRemoveDuplicates = False
    If blFindDuplicates = False Then
        For llRowNo = imLBDrf To UBound(tgAllDrf) - 1 Step 1
            tgAllDrf(llRowNo).sKey = tgAllDrf(llRowNo).tDrf.lCode
            Do While Len(Trim(tgAllDrf(llRowNo).sKey)) < Len(tgAllDrf(llRowNo).sKey)
                tgAllDrf(llRowNo).sKey = "0" & Trim(tgAllDrf(llRowNo).sKey)
            Loop
        Next llRowNo
        'Sort in descending to retain the newest
        ArraySortTyp fnAV(tgAllDrf(), 0), UBound(tgAllDrf), 1, LenB(tgAllDrf(0)), 0, LenB(tgAllDrf(0).sKey), 0
    End If
    For llRowNo = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If (tgAllDrf(llRowNo).iStatus = 0) Or (tgAllDrf(llRowNo).iStatus = 1) Then
        If tgAllDrf(llRowNo).iStatus = 0 Then
        llRowNo = llRowNo
        End If
            gUnpackTimeLong tgAllDrf(llRowNo).tDrf.iStartTime(0), tgAllDrf(llRowNo).tDrf.iStartTime(1), False, llStartTime1
            gUnpackTimeLong tgAllDrf(llRowNo).tDrf.iEndTime(0), tgAllDrf(llRowNo).tDrf.iEndTime(1), False, llEndTime1
            blMatch = False
            For llCheckRow = llRowNo + 1 To UBound(tgAllDrf) - 1 Step 1
                If llRowNo <> llCheckRow Then
                    If (tgAllDrf(llCheckRow).iStatus = 0) Or (tgAllDrf(llCheckRow).iStatus = 1) Then
                        If tgAllDrf(llRowNo).tDrf.iVefCode = tgAllDrf(llCheckRow).tDrf.iVefCode Then
                            If tgAllDrf(llRowNo).tDrf.sInfoType = tgAllDrf(llCheckRow).tDrf.sInfoType Then
                                If tgAllDrf(llRowNo).tDrf.sDataType = tgAllDrf(llCheckRow).tDrf.sDataType Then
                                    If (tgAllDrf(llRowNo).tDrf.sInfoType = "D") Or (tgAllDrf(llRowNo).tDrf.sInfoType = "V") Or (tgAllDrf(llRowNo).tDrf.sInfoType = "T") Then
                                        If (tgAllDrf(llRowNo).tDrf.sInfoType = "D") Then
                                            If (tgAllDrf(llRowNo).tDrf.iRdfCode = tgAllDrf(llCheckRow).tDrf.iRdfCode) Then
                                                If tgAllDrf(llRowNo).tDrf.iRdfCode <> 0 Then
                                                    blMatch = True
                                                Else
                                                    'Extra daypart
                                                    blMatch = True
                                                    For ilDay = 0 To 6 Step 1
                                                        If (tgAllDrf(llRowNo).tDrf.sDay(ilDay) <> tgAllDrf(llCheckRow).tDrf.sDay(ilDay)) Then
                                                            blMatch = False
                                                            Exit For
                                                        End If
                                                    Next ilDay
                                                    If blMatch Then
                                                        gUnpackTimeLong tgAllDrf(llCheckRow).tDrf.iStartTime(0), tgAllDrf(llCheckRow).tDrf.iStartTime(1), False, llStartTime2
                                                        gUnpackTimeLong tgAllDrf(llCheckRow).tDrf.iEndTime(0), tgAllDrf(llCheckRow).tDrf.iEndTime(1), False, llEndTime2
                                                        If (llStartTime1 <> llStartTime2) Or (llEndTime1 <> llEndTime2) Then
                                                            blMatch = False
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                'Compare daypart against Extra
                                                If (tgAllDrf(llRowNo).tDrf.iRdfCode = 0) Or (tgAllDrf(llCheckRow).tDrf.iRdfCode = 0) Then
                                                    If (tgAllDrf(llRowNo).tDrf.iRdfCode <> 0) Then
                                                        ilRdfCode = tgAllDrf(llRowNo).tDrf.iRdfCode
                                                        llExtraRowNo = llCheckRow
                                                    Else
                                                        ilRdfCode = tgAllDrf(llCheckRow).tDrf.iRdfCode
                                                        llExtraRowNo = llRowNo
                                                    End If
                                                    ilRdf = gBinarySearchRdf(ilRdfCode)
                                                    If ilRdf <> -1 Then
                                                        tlRdf = tgMRdf(ilRdf)
                                                        ilRdfCount = 0
                                                        For ilRdf = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1
                                                            If (tlRdf.iStartTime(0, ilRdf) <> 1) Or (tlRdf.iStartTime(1, ilRdf) <> 0) Then
                                                                ilRdfCount = ilRdfCount + 1
                                                                gUnpackTimeLong tlRdf.iStartTime(0, ilRdf), tlRdf.iStartTime(1, ilRdf), False, llRdfStartTime
                                                                gUnpackTimeLong tlRdf.iEndTime(0, ilRdf), tlRdf.iEndTime(1, ilRdf), False, llRdfEndTime
                                                                For ilDay = 1 To 7 Step 1
                                                                    If tlRdf.sWkDays(ilRdf, ilDay - 1) = "Y" Then
                                                                        slDay(ilDay - 1) = "Y"
                                                                    Else
                                                                        slDay(ilDay - 1) = "N"
                                                                    End If
                                                                Next ilDay
                                                            End If
                                                        Next ilRdf
                                                        If ilRdfCount = 1 Then
                                                            gUnpackTimeLong tgAllDrf(llExtraRowNo).tDrf.iStartTime(0), tgAllDrf(llExtraRowNo).tDrf.iStartTime(1), False, llStartTime2
                                                            gUnpackTimeLong tgAllDrf(llExtraRowNo).tDrf.iEndTime(0), tgAllDrf(llExtraRowNo).tDrf.iEndTime(1), False, llEndTime2
                                                            blMatch = True
                                                             For ilDay = 0 To 6 Step 1
                                                                 If (tgAllDrf(llExtraRowNo).tDrf.sDay(ilDay) <> slDay(ilDay)) Then
                                                                     blMatch = False
                                                                     Exit For
                                                                 End If
                                                             Next ilDay
                                                             If blMatch Then
                                                                 If (llRdfStartTime <> llStartTime2) Or (llRdfEndTime <> llEndTime2) Then
                                                                     blMatch = False
                                                                 End If
                                                             End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf (tgAllDrf(llRowNo).tDrf.sInfoType = "V") Then
                                            blMatch = True
                                            For ilDay = 0 To 6 Step 1
                                                If (tgAllDrf(llRowNo).tDrf.sDay(ilDay) <> tgAllDrf(llCheckRow).tDrf.sDay(ilDay)) Then
                                                    blMatch = False
                                                    Exit For
                                                End If
                                            Next ilDay
                                        ElseIf (tgAllDrf(llRowNo).tDrf.sInfoType = "T") Then
                                            blMatch = True
                                            For ilDay = 0 To 6 Step 1
                                                If (tgAllDrf(llRowNo).tDrf.sDay(ilDay) <> tgAllDrf(llCheckRow).tDrf.sDay(ilDay)) Then
                                                    blMatch = False
                                                    Exit For
                                                End If
                                            Next ilDay
                                            If blMatch Then
                                                gUnpackTimeLong tgAllDrf(llCheckRow).tDrf.iStartTime(0), tgAllDrf(llCheckRow).tDrf.iStartTime(1), False, llStartTime2
                                                gUnpackTimeLong tgAllDrf(llCheckRow).tDrf.iEndTime(0), tgAllDrf(llCheckRow).tDrf.iEndTime(1), False, llEndTime2
                                                If (llStartTime1 <> llStartTime2) Or (llEndTime1 <> llEndTime2) Then
                                                    blMatch = False
                                                End If
                                            End If
                                        End If
                                        If blMatch Then
                                            If blFindDuplicates Then
                                                mFindOrRemoveDuplicates = True
                                                Exit Function
                                            Else
                                                'Remove duplicate record, Copy to delele array not required as mSave wikk test for status = -1 and lCode > 0
                                                tgAllDrf(llCheckRow).iStatus = -1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next llCheckRow
        End If
    Next llRowNo
End Function

Private Sub mSetControls(blFromModel As Boolean)
    imRowOffset = 2
    If smSource = "I" Then 'Podcast Impression mode
        imRowOffset = 1
    End If
    If (imSelectedIndex > 1) Or blFromModel Then
        If smSource <> "I" Then 'Standard Airtime mode
            pbcSpec.Width = tmSCtrls(POPINDEX + 17).fBoxX + tmSCtrls(POPINDEX + 17).fBoxW + 30
            pbcSpec.Height = tmSCtrls(POPINDEX + 17).fBoxY + tmSCtrls(POPINDEX + 17).fBoxH + 15 '30
            plcSpec.Width = pbcSpec.Width + fgPanelAdj
            plcSpec.Height = pbcSpec.Height + fgPanelAdj
            plcDataType.Visible = True
            ckcSocEco.Visible = True
            tmDCtrls(DDAYPARTINDEX) = tmDP
            tmDCtrls(DGROUPINDEX) = tmGroup
            
            pbcDemo(0).Width = tmDCtrls(DDEMOINDEX + 17).fBoxX + tmDCtrls(DDEMOINDEX + 17).fBoxW + 15
            plcDemo.Width = pbcDemo(0).Width + vbcDemo.Width + fgPanelAdj
            vbcDemo.Left = plcDemo.Left + plcDemo.Width - vbcDemo.Width - 30 '- fgPanelAdj
            If imDPorEst = 0 Then
                'If smSource <> "I" Then 'Standard Airtime mode
                pbcPlus.Visible = True
                'End If
                plcPlus.Visible = True
                vbcPlus.Visible = True
                pbcEst.Visible = False
                pbcUSA.Visible = False
            Else
                pbcPlus.Visible = False
                plcPlus.Visible = False
                vbcPlus.Visible = False
                If imEstByLOrU = 1 Then
                    pbcUSA.Visible = True
                Else
                    pbcEst.Visible = True
                End If
            End If
            If tgSpf.sDemoEstAllowed <> "Y" Then
                lacPlusTitle.Visible = True
            Else
                lacPlusTitle.Visible = False
            End If
            If cbcDemo.ListCount > 0 Then
                cbcDemo.Visible = True
            End If
            cmcDuplicate.Visible = True
            cmcBaseDuplicate.Visible = True
        Else    'Podcast Impression mode
            pbcSpec.Width = tmSCtrls(DATEINDEX).fBoxX + tmSCtrls(DATEINDEX).fBoxW + 30
            pbcSpec.Height = tmSCtrls(NAMEINDEX).fBoxY + tmSCtrls(NAMEINDEX).fBoxH + 15 '30
            plcSpec.Width = pbcSpec.Width + fgPanelAdj
            plcSpec.Height = pbcSpec.Height + fgPanelAdj
            plcDataType.Visible = False
            rbcDataType(0).Value = True 'Daypart
            ckcSocEco.Visible = False
            tmDCtrls(DDAYPARTINDEX).fBoxW = 2 * tmDP.fBoxW
            tmDCtrls(DGROUPINDEX).fBoxX = tmDP.fBoxX + tmDCtrls(DDAYPARTINDEX).fBoxW + 15
            tmDCtrls(DGROUPINDEX).fBoxW = tmDCtrls(DDAYPARTINDEX).fBoxW
            pbcDemo(0).Width = tmDCtrls(DGROUPINDEX).fBoxX + tmDCtrls(DGROUPINDEX).fBoxW + 15
            plcDemo.Width = pbcDemo(0).Width + vbcDemo.Width + fgPanelAdj
            vbcDemo.Left = plcDemo.Left + plcDemo.Width - vbcDemo.Width - 30 '- fgPanelAdj
            lacPlusTitle.Visible = False
            pbcPlus.Visible = False
            plcPlus.Visible = False
            vbcPlus.Visible = False
            pbcEst.Visible = False
            pbcUSA.Visible = False
            cbcDemo.Visible = False
            cmcDuplicate.Visible = False
            cmcBaseDuplicate.Visible = False
        End If
    Else
        pbcSpec.Width = tmSCtrls(POPINDEX + 17).fBoxX + tmSCtrls(POPINDEX + 17).fBoxW + 30
        pbcSpec.Height = tmSCtrls(POPINDEX + 17).fBoxY + tmSCtrls(POPINDEX + 17).fBoxH + 15 '30
        plcSpec.Width = pbcSpec.Width + fgPanelAdj
        plcSpec.Height = pbcSpec.Height + fgPanelAdj
        plcDataType.Visible = True
        ckcSocEco.Visible = True
        tmDCtrls(DDAYPARTINDEX) = tmDP
        tmDCtrls(DGROUPINDEX) = tmGroup
        pbcDemo(0).Width = tmDCtrls(DDEMOINDEX + 17).fBoxX + tmDCtrls(DDEMOINDEX + 17).fBoxW + 15
        plcDemo.Width = pbcDemo(0).Width + vbcDemo.Width + fgPanelAdj
        vbcDemo.Left = plcDemo.Left + plcDemo.Width - vbcDemo.Width - 30 '- fgPanelAdj
        If imDPorEst = 0 Then
            If smSource <> "I" Then 'Standard Airtime mode
                pbcPlus.Visible = True
            End If
            plcPlus.Visible = True
            vbcPlus.Visible = True
            pbcEst.Visible = False
            pbcUSA.Visible = False
        Else
            pbcPlus.Visible = False
            plcPlus.Visible = False
            vbcPlus.Visible = False
            If imEstByLOrU = 1 Then
                pbcUSA.Visible = True
            Else
                pbcEst.Visible = True
            End If
        End If
        If tgSpf.sDemoEstAllowed <> "Y" Then
            lacPlusTitle.Visible = True
        Else
            lacPlusTitle.Visible = False
        End If
        If cbcDemo.ListCount > 0 Then
            cbcDemo.Visible = True
        End If
        cmcDuplicate.Visible = True
        cmcBaseDuplicate.Visible = True
    End If
    lacFrame(0).Height = 2 * (fgBoxGridH) + 75
    lacFrame(0).Width = pbcDemo(0).Width - 15
    vbcDemo.LargeChange = lmVbcDemoLargeChg \ imRowOffset - imRowOffset
    If smSource = "I" Then 'Podcast Impression mode
        lacFrame(0).Height = (fgBoxGridH) + 60
    End If
End Sub

Private Sub mGetImpressions(llRowNo As Long)
    Dim ilRet As Integer
    Dim slDemoStr As String
    Dim llDrfCode As Long
    
    If smSource <> "I" Then 'Standard Airtime mode
        Exit Sub
    End If
    If tgDrfRec(llRowNo).iModel Then
        llDrfCode = tgDrfRec(llRowNo).lModelDrfCode
    Else
        llDrfCode = tgDrfRec(llRowNo).tDrf.lCode
    End If
    tmDpfSrchKey1.lDrfCode = llDrfCode
    tmDpfSrchKey1.iMnfDemo = 0
    ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = llDrfCode) Then
        If tgSpf.sSAudData = "H" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            slDemoStr = gLongToStrDec(tmDpf.lDemo, 3)
        Else
            slDemoStr = Trim$(Str$(tmDpf.lDemo))
        End If
        'Debug.Print "DRF Code:" & llDrfCode & " = " & slDemoStr
        tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX) = slDemoStr
    Else
        'Debug.Print "DRF Code:" & llDrfCode & " = ''"
        tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX) = ""
    End If
End Sub

Private Sub mPutImpressions(llRowNo As Long)
    Dim ilRet As Integer
    Dim llDemo As Long
    
    If smSource <> "I" Then 'Standard Airtime mode
        Exit Sub
    End If
    If llRowNo >= UBound(tmSaveShow) Then
        Exit Sub
    End If
    If Trim$(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX)) = "" Then
        Exit Sub
    End If
    If tgSpf.sSAudData = "H" Then
        llDemo = gStrDecToLong(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX), 1)
    ElseIf tgSpf.sSAudData = "N" Then
        llDemo = gStrDecToLong(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX), 2)
    ElseIf tgSpf.sSAudData = "U" Then
        llDemo = gStrDecToLong(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX), 3)
    Else
        llDemo = Val(tmSaveShow(llRowNo).sSave(DIMPRESSIONSINDEX))
    End If
    'Fix TTP 10759 - Per Teams Jason LeVine 4:21 PM - I tried TTP 10759 on this new release, and it was still happening
    If tgDrfRec(llRowNo).tDrf.lCode = 0 Then
        tmDpfSrchKey1.lDrfCode = tgAllDrf(llRowNo).tDrf.lCode
    Else
        tmDpfSrchKey1.lDrfCode = tgDrfRec(llRowNo).tDrf.lCode
    End If
    tmDpfSrchKey1.iMnfDemo = imP12PlusMnfCode 'TTP 10759 - Research List screen - Impressions book: manually added/edited impressions can disappear or change when saving
    ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tgDrfRec(llRowNo).tDrf.lCode) Then
        'Update
        tmDpf.lDemo = llDemo
        ilRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
    Else
        'Add
        tmDpf.lCode = 0
        'Fix TTP 10759 - Per Teams Jason LeVine 4:21 PM - I tried TTP 10759 on this new release, and it was still happening
        If tgDrfRec(llRowNo).tDrf.lCode = 0 Then
            tmDpf.lDrfCode = tgAllDrf(llRowNo).tDrf.lCode
            tmDpf.iDnfCode = tgAllDrf(llRowNo).tDrf.iDnfCode
        Else
            tmDpf.lDrfCode = tgDrfRec(llRowNo).tDrf.lCode
            tmDpf.iDnfCode = tgDrfRec(llRowNo).tDrf.iDnfCode
        End If
        tmDpf.iMnfDemo = imP12PlusMnfCode
        'tmDpf.iDnfCode = tgDrfRec(llRowNo).tDrf.iDnfCode
        tmDpf.lDemo = llDemo
        tmDpf.lPop = 0
        tmDpf.sUnused = ""
        ilRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
    End If
End Sub

Private Sub mRemoveImpressions(llRowNo As Long)
    Dim ilRet As Integer
    Dim slDemoStr As String
    
    If smSource <> "I" Then 'Standard Airtime mode
        Exit Sub
    End If
    Do
        tmDpfSrchKey1.lDrfCode = tgDrfRec(llRowNo).tDrf.lCode
        tmDpfSrchKey1.iMnfDemo = imP12PlusMnfCode 'TTP 10759 - Research List screen
        ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tgDrfRec(llRowNo).tDrf.lCode) Then
            ilRet = btrDelete(hmDpf)
        Else
            Exit Do
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
End Sub

Private Sub mGetP12Plus()
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim slCode As String
    
    imP12PlusMnfCode = 0
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    ilFilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilFilter(1) = INTEGERFILTER
    slFilter(1) = "0"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
    lbcDemo.Clear
    smNameCodeTag = ""
    ilRet = gIMoveListBox(Research, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    For ilIndex = LBound(tmNameCode) To UBound(tmNameCode) - 1 Step 1
        ilRet = gParseItem(tmNameCode(ilIndex).sKey, 1, "\", slName)    'Get application name
        ilRet = gParseItem(tmNameCode(ilIndex).sKey, 2, "\", slCode)    'Get application name
        If slName = "P12+" Then
            imP12PlusMnfCode = Val(slCode)
            Exit For
        End If
    Next ilIndex
    lbcDemo.Clear
End Sub

Private Sub mBuildRearchAdjustVehicles(blAfterModelCall As Boolean)
    Dim llRowNo As Long
    Dim blFd As Boolean
    Dim llVef As Long
    Dim ilVef As Integer
    Dim llUpper As Long
    
    ReDim tgResearchAdjustVehicle(0 To 0) As RESEARCHADJUSTVEHICLE
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    For llRowNo = imLBDrf To UBound(tgAllDrf) - 1 Step 1
        If tgAllDrf(llRowNo).tDrf.iVefCode > 0 Then
            blFd = False
            For llVef = 0 To UBound(tgResearchAdjustVehicle) - 1 Step 1
                If tgAllDrf(llRowNo).tDrf.iVefCode = tgResearchAdjustVehicle(llVef).iVefCode Then
                    blFd = True
                    Exit For
                End If
            Next llVef
        Else
            blFd = True
        End If
        If (Not blFd) And (tgAllDrf(llRowNo).tDrf.iVefCode > 0) Then
            llUpper = UBound(tgResearchAdjustVehicle)
            If blAfterModelCall Then
                tgResearchAdjustVehicle(llUpper).bAdjust = True 'False
            Else
                tgResearchAdjustVehicle(llUpper).bAdjust = False
            End If
            tgResearchAdjustVehicle(llUpper).iVefCode = tgAllDrf(llRowNo).tDrf.iVefCode
            ilVef = gBinarySearchVef(tgAllDrf(llRowNo).tDrf.iVefCode)
            If ilVef <> -1 Then
                tgResearchAdjustVehicle(llUpper).sVehicleName = tgMVef(ilVef).sName
                ReDim Preserve tgResearchAdjustVehicle(0 To llUpper + 1) As RESEARCHADJUSTVEHICLE
            End If
        End If
    Next llRowNo
End Sub

