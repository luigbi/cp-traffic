VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form SpecRpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Specifications"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   2145
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
   ScaleWidth      =   9360
   Begin VB.FileListBox lbcRptBmpNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   225
      Left            =   5265
      Pattern         =   "*.bmp;*.jpg;*.gif"
      TabIndex        =   66
      Top             =   120
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Height          =   285
      Left            =   6030
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.FileListBox lbcFileNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   225
      Left            =   7290
      Pattern         =   "*.bmp"
      TabIndex        =   58
      Top             =   105
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7260
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7695
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5475
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8055
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   5340
   End
   Begin VB.Timer tmcScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8325
      Top             =   5310
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
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
      Left            =   4230
      TabIndex        =   38
      Top             =   5520
      Width           =   1050
   End
   Begin VB.PictureBox plcRptName 
      Height          =   4665
      Left            =   255
      ScaleHeight     =   4605
      ScaleWidth      =   8790
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   675
      Width           =   8850
      Begin VB.TextBox edcRptFormName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1185
         MaxLength       =   12
         TabIndex        =   69
         Top             =   990
         Width           =   1725
      End
      Begin VB.HScrollBar hbcRptSample 
         Height          =   210
         LargeChange     =   8490
         Left            =   60
         TabIndex        =   68
         Top             =   3750
         Width           =   8490
      End
      Begin VB.CommandButton cmcRefresh 
         Appearance      =   0  'Flat
         Caption         =   "&Refresh"
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
         Left            =   3615
         TabIndex        =   15
         Top             =   2535
         Width           =   1050
      End
      Begin VB.ComboBox cbcRptSample 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1560
         TabIndex        =   14
         Top             =   2535
         Width           =   1845
      End
      Begin VB.CommandButton cmcErase 
         Appearance      =   0  'Flat
         Caption         =   "&Erase"
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
         Left            =   3945
         TabIndex        =   26
         Top             =   4305
         Width           =   1050
      End
      Begin VB.ComboBox cbcJobList 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   4020
         TabIndex        =   8
         Top             =   960
         Width           =   2220
      End
      Begin VB.CommandButton cmcRptUndo 
         Appearance      =   0  'Flat
         Caption         =   "&Undo"
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
         Height          =   285
         Left            =   5325
         TabIndex        =   27
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton cmcRptSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save Name"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   25
         Top             =   4305
         Width           =   1050
      End
      Begin VB.VScrollBar vbcRptSample 
         Height          =   1065
         LargeChange     =   870
         Left            =   8550
         SmallChange     =   525
         TabIndex        =   18
         Top             =   2895
         Width           =   210
      End
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   0
         Left            =   60
         ScaleHeight     =   1035
         ScaleWidth      =   8460
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2895
         Width           =   8490
         Begin VB.PictureBox pbcRptSample 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   1
            Left            =   15
            ScaleHeight     =   165
            ScaleWidth      =   6060
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   6060
         End
      End
      Begin VB.PictureBox plcRptState 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5985
         ScaleHeight     =   225
         ScaleWidth      =   2730
         TabIndex        =   22
         Top             =   4005
         Width           =   2730
         Begin VB.OptionButton rbcRptState 
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   495
            TabIndex        =   23
            Top             =   0
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton rbcRptState 
            Caption         =   "Dormant"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox plcRptMoney 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   165
         ScaleHeight     =   225
         ScaleWidth      =   4455
         TabIndex        =   19
         Top             =   4005
         Width           =   4455
         Begin VB.OptionButton rbcRptMoney 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   2100
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton rbcRptMoney 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   2775
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.TextBox edcRptDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   1140
         Left            =   1185
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1335
         Width           =   7545
      End
      Begin VB.TextBox edcRptPassValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   7470
         MaxLength       =   10
         TabIndex        =   10
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox edcRptName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   4080
         MaxLength       =   60
         TabIndex        =   5
         Top             =   615
         Width           =   4680
      End
      Begin VB.ComboBox cbcCatRpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   4620
         TabIndex        =   1
         Top             =   120
         Width           =   4125
      End
      Begin VB.PictureBox plcRptType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2535
         TabIndex        =   2
         Top             =   645
         Width           =   2535
         Begin VB.OptionButton rbcRptType 
            Caption         =   "Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   1545
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton rbcRptType 
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   480
            TabIndex        =   3
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Label lacJobList 
         Appearance      =   0  'Flat
         Caption         =   "Job/List #"
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
         Height          =   195
         Left            =   3105
         TabIndex        =   7
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   8745
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   8745
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lacRptDescription 
         Appearance      =   0  'Flat
         Caption         =   "Description"
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
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label lacRptPassValue 
         Appearance      =   0  'Flat
         Caption         =   "Pass Value"
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
         Height          =   195
         Left            =   6420
         TabIndex        =   9
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lacRptSample 
         Appearance      =   0  'Flat
         Caption         =   "Sample Picture"
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
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   2580
         Width           =   1395
      End
      Begin VB.Label lacRptExe 
         Appearance      =   0  'Flat
         Caption         =   "Form Name"
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
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label lacRptName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
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
         Height          =   195
         Left            =   3450
         TabIndex        =   40
         Top             =   645
         Width           =   570
      End
   End
   Begin VB.PictureBox plcBrowser 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   3195
      ScaleHeight     =   4620
      ScaleWidth      =   7965
      TabIndex        =   41
      Top             =   5805
      Visible         =   0   'False
      Width           =   7995
      Begin VB.VScrollBar vbcBrowser 
         Height          =   1545
         Left            =   7650
         TabIndex        =   53
         Top             =   2970
         Width           =   240
      End
      Begin VB.PictureBox pbcBrowserSample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   0
         Left            =   105
         ScaleHeight     =   1515
         ScaleWidth      =   7530
         TabIndex        =   51
         Top             =   2970
         Width           =   7560
         Begin VB.PictureBox pbcBrowserSample 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1515
            Index           =   1
            Left            =   0
            ScaleHeight     =   1515
            ScaleWidth      =   6975
            TabIndex        =   52
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.CommandButton cmcBrowserOpen 
         Appearance      =   0  'Flat
         Caption         =   "D&one"
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
         Left            =   6435
         TabIndex        =   49
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton cmcBrowserCancel 
         Appearance      =   0  'Flat
         Caption         =   "C&ancel"
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
         Left            =   6450
         TabIndex        =   50
         Top             =   915
         Width           =   1050
      End
      Begin VB.TextBox edcBrowserFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   420
         TabIndex        =   43
         Top             =   360
         Width           =   2505
      End
      Begin VB.FileListBox lbcBrowserFile 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   420
         Pattern         =   "*.bmp"
         TabIndex        =   44
         Top             =   840
         Width           =   2505
      End
      Begin VB.DriveListBox cbcBrowserDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3465
         TabIndex        =   46
         Top             =   360
         Width           =   2565
      End
      Begin VB.DirListBox lbcBrowserPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1830
         Left            =   3465
         TabIndex        =   47
         Top             =   840
         Width           =   2565
      End
      Begin VB.Label lacBrowserSample 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Sample"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3420
         TabIndex        =   48
         Top             =   2715
         Width           =   1080
      End
      Begin VB.Label lacBrowserFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   420
         TabIndex        =   42
         Top             =   120
         Width           =   2490
      End
      Begin VB.Label lacBrowserPath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Path"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3480
         TabIndex        =   45
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.PictureBox plcTree 
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   6000
      ScaleHeight     =   4635
      ScaleWidth      =   8775
      TabIndex        =   28
      Top             =   4935
      Width           =   8835
      Begin VB.PictureBox plcTreeAutoReturn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4680
         ScaleHeight     =   225
         ScaleWidth      =   4050
         TabIndex        =   62
         Top             =   225
         Width           =   4050
         Begin VB.OptionButton rbcTreeAutoReturn 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   3555
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
         Begin VB.OptionButton rbcTreeAutoReturn 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   2985
            TabIndex        =   63
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox pbcSubTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3870
         Picture         =   "SPECRPT.frx":0000
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   255
         Width           =   180
      End
      Begin VB.PictureBox pbcAddTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         Picture         =   "SPECRPT.frx":00DA
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   255
         Width           =   180
      End
      Begin VB.CommandButton cmcTreeSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save Tree"
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
         Height          =   285
         Left            =   3945
         TabIndex        =   37
         Top             =   4335
         Width           =   1050
      End
      Begin VB.TextBox edcTreeDescription 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   3240
         Width           =   8535
      End
      Begin VB.CommandButton cmcTreeDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
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
         Left            =   5835
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1050
      End
      Begin VB.PictureBox pbcDnMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3615
         Picture         =   "SPECRPT.frx":01B4
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2895
         Width           =   180
      End
      Begin VB.ListBox lbcTree 
         Appearance      =   0  'Flat
         Height          =   2175
         Index           =   1
         Left            =   3480
         TabIndex        =   34
         Top             =   450
         Width           =   5205
      End
      Begin VB.CommandButton cmcTreeAdd 
         Appearance      =   0  'Flat
         Caption         =   "&Add      "
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
         Left            =   2910
         TabIndex        =   31
         Top             =   2835
         Width           =   945
      End
      Begin VB.ListBox lbcTree 
         Appearance      =   0  'Flat
         Height          =   2175
         Index           =   0
         Left            =   150
         TabIndex        =   30
         Top             =   450
         Width           =   3105
      End
      Begin VB.Image imcPrt 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   8190
         Picture         =   "SPECRPT.frx":028E
         Top             =   2775
         Width           =   480
      End
      Begin VB.Label lacTab 
         Appearance      =   0  'Flat
         Caption         =   "Tab"
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
         Height          =   195
         Left            =   3495
         TabIndex        =   61
         Top             =   225
         Width           =   315
      End
      Begin VB.Label lacSelection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Category && Report Names"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   420
         TabIndex        =   29
         Top             =   30
         Width           =   2595
      End
      Begin VB.Label lacSelection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Report Tree Structure"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   4740
         TabIndex        =   33
         Top             =   30
         Width           =   2160
      End
      Begin VB.Label lacTreeDescription 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   150
         TabIndex        =   65
         Top             =   4260
         Width           =   3630
      End
   End
   Begin ComctlLib.TabStrip plcTabSelection 
      Height          =   5160
      Left            =   90
      TabIndex        =   67
      Top             =   270
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   9102
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Report &Names"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Report &Tree Structure"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacDBPath 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   57
      Top             =   5850
      Width           =   6060
   End
   Begin VB.Image imcTrashClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7695
      Picture         =   "SPECRPT.frx":0598
      Top             =   5415
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8070
      Picture         =   "SPECRPT.frx":08A2
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8805
      Picture         =   "SPECRPT.frx":0BAC
      Top             =   5430
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "SpecRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpecRpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Show/Fix screen code
Option Explicit
Option Compare Text
'Report Name
Dim hmRnf As Integer 'Report Name file handle
Dim tmRnf As RNF        'RNF record image
Dim tmRnfSrchKey As INTKEY0    'RNF key record image
Dim imRnfRecLen As Integer        'RNF record length
Dim tmTempRnf As RNF
'Report Tree
Dim hmRtf As Integer 'Report Tree file handle
Dim tmRtf As RTF        'RTF record image
Dim tmRtfSrchKey As INTKEY0    'RNF key record image
Dim imRtfRecLen As Integer        'RNF record length
Dim imShowHelpMsg As Integer
Dim imNoRowLbc As Integer
Dim imPasswordAsked As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imCatRptChg As Integer
Dim imTreeChg As Integer
Dim imTabSelection As Integer
Dim imChgMode As Integer
Dim imSelectedIndex As Integer
'Dim imRptExtSelectedIndex As Integer
Dim imJobListSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer
'Drag
Dim imDragIndexSrce As Integer  '
Dim imDragIndexDest As Integer  '
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imDragSrce As Integer 'Values defined below
Dim imDragDest As Integer 'Values defined below
Dim imDragScroll As Integer 'Object to be scrolled (same values as below)
Dim imScrollDir As Integer  '0=Page up; 1=Page down
Const DRAGTREE0 = 1
Const DRAGTREE1 = 2
Private Sub cbcBrowserDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcBrowserPath.Path = cbcBrowserDrive.Drive
    lbcBrowserFile.Clear
    Screen.MousePointer = vbDefault
End Sub
Private Sub cbcCatRpt_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcCatRpt, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcCatRpt.ListIndex
        If Not mReadRec(ilIndex) Then
            GoTo cbcCatRptErr
        End If
    Else
        If ilRet = 1 Then
            cbcCatRpt.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    If ilRet = 0 Then
        imSelectedIndex = cbcCatRpt.ListIndex
        mMoveRecToCtrl
        rbcRptType(0).Enabled = False
        rbcRptType(1).Enabled = False
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcRptName.Text = slStr
        End If
        rbcRptType(0).Enabled = True
        rbcRptType(1).Enabled = True
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
    imChgMode = False
    Exit Sub
cbcCatRptErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcCatRpt_Click()
    cbcCatRpt_Change
End Sub
Private Sub cbcCatRpt_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    Dim ilSvIndex As Integer
    Dim ilLoop As Integer
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'If igAdvtCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        '    If sgAdvtName = "" Then
        '        cbcSelect.ListIndex = 0
        '    Else
        '        cbcSelect.Text = sgAdvtName    'New name
        '    End If
        '    cbcSelect_Change
        '    If sgAdvtName <> "" Then
        '        mSetCommands
        '        gFindMatch sgAdvtName, 1, cbcSelect
        '        If gLastFound(cbcSelect) > 0 Then
        '            cmcDone.SetFocus
        '            Exit Sub
        '        End If
        '    End If
        '    If pbcSTab.Enabled Then
        '        pbcSTab.SetFocus
        '    Else
        '        cmcCancel.SetFocus
        '    End If
        '    Exit Sub
        'End If
    End If
    slSvText = cbcCatRpt.Text
    If cbcCatRpt.ListCount <= 1 Then
        cbcCatRpt.ListIndex = 0
        mClearCtrlFields
        'If pbcSTab.Enabled Then
        '    pbcSTab.SetFocus
        'Else
        '    cmcCancel.SetFocus
        'End If
        Exit Sub
    End If
    gCtrlGotFocus cbcCatRpt
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcCatRpt.ListIndex = 0
        cbcCatRpt_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcCatRpt
        If gLastFound(cbcCatRpt) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcCatRpt)) Or (ilSvIndex <> cbcCatRpt.ListIndex) Then
            If (slSvText <> cbcCatRpt.List(gLastFound(cbcCatRpt))) Then
                cbcCatRpt.ListIndex = gLastFound(cbcCatRpt)
                cbcCatRpt_Change    'Call change so picture area repainted
            End If
        Else
            cbcCatRpt.ListIndex = 0
            mClearCtrlFields
            cbcCatRpt_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcCatRpt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcCatRpt_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcCatRpt.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcJobList_Change()
    Dim ilRet As Integer
    Dim slStr As String
    If imChgMode = False Then
        imChgMode = True
        If cbcJobList.Text <> "" Then
            gManLookAhead cbcJobList, imBSMode, imComboBoxIndex
        End If
        imJobListSelectedIndex = cbcJobList.ListIndex
        imChgMode = False
    End If
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub cbcJobList_Click()
    cbcJobList_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcJobList_GotFocus()
    If imJobListSelectedIndex = -1 Then
        gFindMatch "[None]", 0, cbcJobList
        If gLastFound(cbcJobList) >= 0 Then
            cbcJobList.ListIndex = gLastFound(cbcJobList)
        Else
            If cbcJobList.ListCount > 0 Then
                cbcJobList.ListIndex = 0
            End If
        End If
    End If
    imComboBoxIndex = imJobListSelectedIndex
    gCtrlGotFocus cbcJobList
End Sub
Private Sub cbcJobList_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcJobList_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcJobList.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcRptSample_Change()
    Dim slFromFile As String
    Screen.MousePointer = vbHourglass
    vbcRptSample.Value = vbcRptSample.Min
    hbcRptSample.Value = hbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    slFromFile = sgRptPath & cbcRptSample.List(cbcRptSample.ListIndex)
    On Error GoTo cbcRptSampleErr:
    pbcRptSample(1).Picture = LoadPicture(slFromFile)
    'vbcRptSample.Top = 0
    vbcRptSample.Max = pbcRptSample(1).Height - pbcRptSample(0).Height
    hbcRptSample.Max = pbcRptSample(1).Width - pbcRptSample(0).Width
    vbcRptSample.Enabled = (pbcRptSample(0).Height < pbcRptSample(1).Height)
    hbcRptSample.Enabled = (pbcRptSample(0).Width < pbcRptSample(1).Width)
    If vbcRptSample.Enabled Then
        vbcRptSample.SmallChange = pbcRptSample(0).Height
        vbcRptSample.LargeChange = pbcRptSample(0).Height
    End If
    If hbcRptSample.Enabled Then
        hbcRptSample.SmallChange = pbcRptSample(0).Width
        hbcRptSample.LargeChange = pbcRptSample(0).Width
    End If
    imCatRptChg = True
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub
cbcRptSampleErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub cbcRptSample_Click()
    cbcRptSample_Change
End Sub
Private Sub cmcBrowserCancel_Click()
    cmcDone.Enabled = True
    plcBrowser.Visible = False
    mSetCommands
End Sub
Private Sub cmcBrowserOpen_Click()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    slName = Trim$(edcBrowserFile.Text)
    If Len(slName) <= 0 Then
        edcBrowserFile.SetFocus
        Exit Sub
    End If
    ilPos = InStr(slName, "*")
    If ilPos > 0 Then
        lbcBrowserFile.fileName = lbcBrowserPath.Path & "\" & slName
        edcBrowserFile.SetFocus
        Exit Sub
    End If
    ilPos = InStr(slName, "?")
    If ilPos > 0 Then
        lbcBrowserFile.fileName = lbcBrowserPath.Path & "\" & slName
        edcBrowserFile.SetFocus
        Exit Sub
    End If
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcBrowserPath.Path & "\" & slFromFile
    End If
    If InStr(slFromFile, ".") = 0 Then
        slFromFile = slFromFile & ".bmp"
    End If
    Screen.MousePointer = vbHourglass
    vbcBrowser.Value = vbcBrowser.Min
    pbcBrowserSample(1).Move 0, 0
    pbcBrowserSample(1).Picture = LoadPicture(slFromFile)
    'vbcBrowser.Top = 0
    vbcBrowser.Max = pbcBrowserSample(1).Height - pbcBrowserSample(0).Height
    vbcBrowser.Enabled = (pbcBrowserSample(0).Height < pbcBrowserSample(1).Height)
    If vbcBrowser.Enabled Then
        vbcBrowser.SmallChange = pbcBrowserSample(0).Height
        vbcBrowser.LargeChange = pbcBrowserSample(0).Height
    End If
    Screen.MousePointer = vbDefault
    cmcBrowserCancel_Click
    On Error GoTo 0
    Exit Sub
cmcBrowserOpenErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim slMess As String
    Dim ilRes As Integer
    If Not igPasswordOk Then
        mTerminate
        Exit Sub
    End If
    If imCatRptChg Then
        If cmcRptSave.Enabled Then
            If imSelectedIndex > 0 Then
                slMess = "Save Changes to " & cbcCatRpt.List(imSelectedIndex)
            Else
                slMess = "Add " & edcRptName.Text
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                Exit Sub
            End If
            If ilRes = vbYes Then
                ilRet = mCatRptSaveRec()
                If Not ilRet Then
                    edcRptName.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    If imTreeChg Then
        If cmcTreeSave.Enabled Then
            slMess = "Save Changes to Report Tree Structure"
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                Exit Sub
            End If
            If ilRes = vbYes Then
                cmcTreeSave_Click
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slMsg As String
    If Not igPasswordOk Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        'ilRet = gIICodeRefExistGivenOffset(SpecRpt, tmRnf.iCode, "Rtf.Btr", 2)
        ilRet = gIICodeRefExist(SpecRpt, tmRnf.iCode, "Rtf.Btr", "RtfRnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - Report Tree Structure references name"
            ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'ilRet = gIICodeRefExistGivenOffset(SpecRpt, tmRnf.iCode, "Srf.Btr", 6)   'chfadfCode
        ilRet = gIICodeRefExist(SpecRpt, tmRnf.iCode, "Srf.Btr", "SrfRnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Set references name"
            ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmRnf.sName, vbOkCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Do  'Loop until record updated or added
            tmRnfSrchKey.iCode = tmRnf.iCode
            imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
            ilRet = btrGetEqual(hmRnf, tmTempRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            ilRet = btrDelete(hmRnf)
            slMsg = "mCatRptSaveRec (btrDelete)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, slMsg, SpecRpt
        On Error GoTo 0
        mClearCtrlFields
        If imTerminate Then
            Screen.MousePointer = vbDefault
            mTerminate
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        mPopulate
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcRefresh_Click()
    Dim ilLoop As Integer
    Dim slStr As String
    If cbcRptSample.ListIndex > 0 Then
        slStr = cbcRptSample.List(cbcRptSample.ListIndex)
    Else
        slStr = ""
    End If
    cbcRptSample.Clear
    'lbcRptBmpNames.Clear
    'Force to repopluate
    lbcRptBmpNames.Path = Left$(sgExePath, Len(sgExePath) - 1)
    lbcRptBmpNames.Path = Left$(sgRptPath, Len(sgRptPath) - 1)
    For ilLoop = 0 To lbcRptBmpNames.ListCount - 1 Step 1
        cbcRptSample.AddItem lbcRptBmpNames.List(ilLoop)
    Next ilLoop
    cbcRptSample.AddItem "[None]", 0
    If slStr <> "" Then
        gFindMatch slStr, 0, cbcRptSample
        If gLastFound(cbcRptSample) < 0 Then
            cbcRptSample.ListIndex = -1
        Else
            cbcRptSample.ListIndex = gLastFound(cbcRptSample)
        End If
    End If
End Sub
Private Sub cmcRptSave_Click()
    Dim ilRet As Integer
    Dim slRptName As String
    If Not igPasswordOk Then
        Exit Sub
    End If
    slRptName = Trim$(edcRptName.Text)
    If rbcRptType(0).Value Then
        slRptName = slRptName & "/Category"
    Else
        slRptName = slRptName & "/Report"
    End If
    ilRet = mCatRptSaveRec()
    If Not ilRet Then
        edcRptName.SetFocus
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        mPopulate
        gFindMatch slRptName, 0, cbcCatRpt    'Determine if name exist
        If gLastFound(cbcCatRpt) > 0 Then   'Name found
            cbcCatRpt.ListIndex = gLastFound(cbcCatRpt)
        End If
        cbcCatRpt.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub cmcRptUndo_Click()
    Dim ilRet As Integer
    If Not igPasswordOk Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        ilRet = mReadRec(imSelectedIndex)
        mMoveRecToCtrl
    Else
        mClearCtrlFields
    End If
    mSetCommands
End Sub
Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slRptName As String
    Dim ilLoop As Integer
    Dim ilRnf As Integer
    Dim ilCount As Integer
    Dim ilTest As Integer
    Dim slChar As String
    Dim slName As String
    Dim slMsg As String
    Dim tlRtf As RTF
    
    If Not igPasswordOk Then
        Exit Sub
    End If
    If cmcRptSave.Enabled Then
        slRptName = Trim$(edcRptDescription.Text)
        If rbcRptType(0).Value Then
            slRptName = slRptName & "/Category"
        Else
            slRptName = slRptName & "/Report"
        End If
        ilRet = mCatRptSaveRec()
        If Not ilRet Then
            edcRptName.SetFocus
            Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            mPopulate
        End If
    End If
    If cmcTreeSave.Enabled Then
        Screen.MousePointer = vbHourglass
        'Add new rtf
        For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
            If tgRtfList(ilLoop).iStatus = 0 Then
                tgRtfList(ilLoop).tRtf.iCode = 0
                ilRet = btrInsert(hmRtf, tgRtfList(ilLoop).tRtf, imRtfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert RTF)"
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, SpecRpt
                On Error GoTo 0
                ilRet = btrGetPosition(hmRtf, tgRtfList(ilLoop).lRecPos)
                tgRtfList(ilLoop).iStatus = 1
            End If
        Next ilLoop
        'Determine levels and set Next
        For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
            ilCount = 0
            slName = lbcTree(1).List(ilLoop)
            For ilTest = 1 To Len(slName) Step 1
                slChar = Mid$(slName, ilTest, 1)
                If slChar = " " Then
                    ilCount = ilCount + 1
                Else
                    Exit For
                End If
            Next ilTest
            tgRtfList(ilLoop).tRtf.iLevel = ilCount \ 2 + 1
            If ilLoop <> UBound(tgRtfList) - 1 Then
                tgRtfList(ilLoop).tRtf.iNextRtfCode = tgRtfList(ilLoop + 1).tRtf.iCode
            Else
                tgRtfList(ilLoop).tRtf.iNextRtfCode = -1
            End If
        Next ilLoop
        'Remove Delete
        For ilLoop = 0 To UBound(tgDelRtfList) - 1 Step 1
            Do
                ilRet = btrGetDirect(hmRtf, tlRtf, imRtfRecLen, tgDelRtfList(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                tmRtfSrchKey.iCode = tlRtf.iCode
                ilRet = btrGetEqual(hmRtf, tlRtf, imRtfRecLen, tmRtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                slMsg = "mSaveRec (btrGetEqual RTF)"
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, SpecRpt
                On Error GoTo 0
                ilRet = btrDelete(hmRtf)
            Loop While ilRet = BTRV_ERR_CONFLICT
            slMsg = "mSaveRec (btrDelete RTF)"
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, SpecRpt
            On Error GoTo 0
        Next ilLoop
    End If
    'Update level
    For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
        Do
            ilRet = btrGetDirect(hmRtf, tlRtf, imRtfRecLen, tgRtfList(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            tmRtfSrchKey.iCode = tlRtf.iCode
            ilRet = btrGetEqual(hmRtf, tlRtf, imRtfRecLen, tmRtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'Update level and Rnf saved values
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If tgRtfList(ilLoop).tRtf.iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                    tgRtfList(ilLoop).tRtf.sRnfType = tgRnfList(ilRnf).tRnf.sType
                    tgRtfList(ilLoop).tRtf.sRnfState = tgRnfList(ilRnf).tRnf.sState
                    Exit For
                End If
            Next ilRnf
            ilRet = btrUpdate(hmRtf, tgRtfList(ilLoop).tRtf, imRtfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mSaveRec (btrUpdate RTF)"
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, SpecRpt
        On Error GoTo 0
    Next ilLoop
    imTreeChg = False
    gFindMatch slRptName, 0, cbcCatRpt    'Determine if name exist
    If gLastFound(cbcCatRpt) > 0 Then   'Name found
        cbcCatRpt.ListIndex = gLastFound(cbcCatRpt)
    End If
    If imTabSelection = 0 Then
        cbcCatRpt.SetFocus
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mTerminate
    Exit Sub
End Sub
Private Sub cmcTreeAdd_Click()
    Dim slName As String
    Dim slTestName As String
    Dim slType As String
    Dim slState As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilOnlyCat As Integer
    Dim ilRet As Integer
    Dim slChar As String
    If Not igPasswordOk Then
        Exit Sub
    End If
    If lbcTree(0).ListIndex < 0 Then
        Exit Sub
    End If
    ilOnlyCat = True
    For ilLoop = 0 To lbcTree(0).ListCount - 1 Step 1
        slName = lbcTree(0).List(ilLoop)
        If InStr(slName, "/Report") > 0 Then
            ilOnlyCat = False
            Exit For
        End If
    Next ilLoop
    'Find RNF
    ilFound = -1
    slName = lbcTree(0).List(lbcTree(0).ListIndex)
    For ilLoop = 0 To UBound(tgRnfList) - 1 Step 1
        'slNameCode = tgRnfList(ilLoop).sKey    'lbcMster.List(ilLoop)
        'ilRet = gParseItem(slNameCode, 1, "\", slTestName)
        'ilRet = gParseItem(slTestName, 1, "|", slType)
        'ilRet = gParseItem(slTestName, 2, "|", slTestName)
        slTestName = Trim$(tgRnfList(ilLoop).tRnf.sName) 'slTestName)
        slType = tgRnfList(ilLoop).tRnf.sType
        slState = tgRnfList(ilLoop).tRnf.sState
        If slType = "C" Then
            slTestName = slTestName & "/Category"
        Else
            slTestName = slTestName & "/Report"
        End If
        If slName = slTestName Then
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = tgRnfList(ilLoop).tRnf.iCode
            ilFound = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilFound = -1 Then
        Exit Sub
    End If
    If lbcTree(1).ListIndex < 0 Then
        'Add To End
        ilIndex = UBound(tgRtfList)
    Else
        'Add After Item Selected
        ilIndex = lbcTree(1).ListIndex + 1
    End If
    'Move all Items Down one
    ReDim Preserve tgRtfList(0 To UBound(tgRtfList) + 1) As RTFLIST
    For ilLoop = UBound(tgRtfList) - 1 To ilIndex Step -1
        tgRtfList(ilLoop + 1) = tgRtfList(ilLoop)
    Next ilLoop
    tgRtfList(ilIndex).sKey = ""
    tgRtfList(ilIndex).tRtf.iCode = 0
    tgRtfList(ilIndex).tRtf.iRnfCode = ilCode   'Val(slCode)
    tgRtfList(ilIndex).tRtf.iLevel = 0
    tgRtfList(ilIndex).tRtf.iNextRtfCode = 0
    tgRtfList(ilIndex).tRtf.sAutoReturn = "Y"
    tgRtfList(ilIndex).tRtf.lColor = 0
    tgRtfList(ilIndex).tRtf.sRnfType = slType
    tgRtfList(ilIndex).tRtf.sRnfState = slState
    tgRtfList(ilIndex).iStatus = 0
    tgRtfList(ilIndex).lRecPos = 0
    'Determine indent
    If (ilIndex > 0) And (Not ilOnlyCat) Then
        slTestName = lbcTree(1).List(ilIndex - 1)
        For ilLoop = 1 To Len(slTestName) Step 1
            slChar = Mid$(slTestName, ilLoop, 1)
            If slChar = " " Then
                slName = slChar & slName
            Else
                Exit For
            End If
        Next ilLoop
        ilPos = InStr(slTestName, "/Category")
        If ilPos > 0 Then
            slName = "  " & slName
        Else
            If slType = "C" Then
                slChar = Mid$(slTestName, 1, 2)
                If slChar = "  " Then
                    slName = right$(slName, Len(slName) - 2)
                End If
            End If
        End If
    End If
    lbcTree(1).AddItem slName, ilIndex
    imTreeChg = True
    mSetCommands
End Sub
Private Sub cmcTreeAdd_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcTreeDelete_Click()
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilBlkCount As Integer
    Dim ilTBlkCount As Integer
    Dim slChar As String
    Dim ilDone As Integer
    If Not igPasswordOk Then
        Exit Sub
    End If
    If lbcTree(1).ListIndex < 0 Then
        Exit Sub
    End If
    ilIndex = lbcTree(1).ListIndex
    'If Category- remove all and all sub categories and reports
    slName = lbcTree(1).List(ilIndex)
    If InStr(slName, "/Category") > 0 Then
        ilBlkCount = 0
        For ilLoop = 1 To Len(slName) Step 1
            slChar = Mid$(slName, ilLoop, 1)
            If slChar = " " Then
                ilBlkCount = ilBlkCount + 1
            Else
                Exit For
            End If
        Next ilLoop
        ilDone = False
        Do
            If tgRtfList(ilIndex).iStatus > 0 Then
                tgDelRtfList(UBound(tgDelRtfList)) = tgRtfList(ilIndex)
                ReDim Preserve tgDelRtfList(0 To UBound(tgDelRtfList) + 1) As RTFLIST
            End If
            For ilLoop = ilIndex To UBound(tgRtfList) - 1 Step 1
                tgRtfList(ilLoop) = tgRtfList(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgRtfList(0 To UBound(tgRtfList) - 1) As RTFLIST
            lbcTree(1).RemoveItem ilIndex
            If ilIndex >= lbcTree(1).ListCount Then
                Exit Do
            End If
            slName = lbcTree(1).List(ilIndex)
            ilTBlkCount = 0
            For ilLoop = 1 To Len(slName) Step 1
                slChar = Mid$(slName, ilLoop, 1)
                If slChar = " " Then
                    ilTBlkCount = ilTBlkCount + 1
                Else
                    Exit For
                End If
            Next ilLoop
            If ilTBlkCount <= ilBlkCount Then
                Exit Do
            End If
        Loop While Not ilDone
    Else    'Report
        If tgRtfList(ilIndex).iStatus > 0 Then
            tgDelRtfList(UBound(tgDelRtfList)) = tgRtfList(ilIndex)
            ReDim Preserve tgDelRtfList(0 To UBound(tgDelRtfList) + 1) As RTFLIST
        End If
        For ilLoop = ilIndex To UBound(tgRtfList) - 1 Step 1
            tgRtfList(ilLoop) = tgRtfList(ilLoop + 1)
        Next ilLoop
        ReDim Preserve tgRtfList(0 To UBound(tgRtfList) - 1) As RTFLIST
        lbcTree(1).RemoveItem ilIndex
    End If
    imTreeChg = True
    mSetCommands
End Sub
Private Sub cmcTreeDelete_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcTreeSave_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim ilTest As Integer
    Dim slChar As String
    Dim slName As String
    Dim ilRnf As Integer
    Dim slMsg As String
    Dim tlRtf As RTF
    'Add new rtf
    If Not igPasswordOk Then
        Exit Sub
    End If
    For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
        If tgRtfList(ilLoop).iStatus = 0 Then
            tgRtfList(ilLoop).tRtf.iCode = 0
            ilRet = btrInsert(hmRtf, tgRtfList(ilLoop).tRtf, imRtfRecLen, INDEXKEY0)
            slMsg = "mTreeSaveRec (btrInsert RTF)"
            On Error GoTo mTreeSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, SpecRpt
            On Error GoTo 0
            ilRet = btrGetPosition(hmRtf, tgRtfList(ilLoop).lRecPos)
            tgRtfList(ilLoop).iStatus = 1
        End If
    Next ilLoop
    'Determine levels and set Next
    For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
        ilCount = 0
        slName = lbcTree(1).List(ilLoop)
        For ilTest = 1 To Len(slName) Step 1
            slChar = Mid$(slName, ilTest, 1)
            If slChar = " " Then
                ilCount = ilCount + 1
            Else
                Exit For
            End If
        Next ilTest
        tgRtfList(ilLoop).tRtf.iLevel = ilCount \ 2 + 1
        If ilLoop <> UBound(tgRtfList) - 1 Then
            tgRtfList(ilLoop).tRtf.iNextRtfCode = tgRtfList(ilLoop + 1).tRtf.iCode
        Else
            tgRtfList(ilLoop).tRtf.iNextRtfCode = -1
        End If
    Next ilLoop
    'Remove Delete
    For ilLoop = 0 To UBound(tgDelRtfList) - 1 Step 1
        Do
            ilRet = btrGetDirect(hmRtf, tlRtf, imRtfRecLen, tgDelRtfList(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            tmRtfSrchKey.iCode = tlRtf.iCode
            ilRet = btrGetEqual(hmRtf, tlRtf, imRtfRecLen, tmRtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            slMsg = "mTreeSaveRec (btrGetEqual RTF)"
            On Error GoTo mTreeSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, SpecRpt
            On Error GoTo 0
            ilRet = btrDelete(hmRtf)
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mTreeSaveRec (btrDelete RTF)"
        On Error GoTo mTreeSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, SpecRpt
        On Error GoTo 0
    Next ilLoop
    'Update level
    For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
        Do
            ilRet = btrGetDirect(hmRtf, tlRtf, imRtfRecLen, tgRtfList(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            tmRtfSrchKey.iCode = tlRtf.iCode
            ilRet = btrGetEqual(hmRtf, tlRtf, imRtfRecLen, tmRtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If tgRtfList(ilLoop).tRtf.iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                    tgRtfList(ilLoop).tRtf.sRnfType = tgRnfList(ilRnf).tRnf.sType
                    tgRtfList(ilLoop).tRtf.sRnfState = tgRnfList(ilRnf).tRnf.sState
                    Exit For
                End If
            Next ilRnf
            ilRet = btrUpdate(hmRtf, tgRtfList(ilLoop).tRtf, imRtfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mTreeSaveRec (btrUpdate RTF)"
        On Error GoTo mTreeSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, SpecRpt
        On Error GoTo 0
    Next ilLoop
    imTreeChg = False
    mSetCommands
    Exit Sub
mTreeSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mTerminate
    Exit Sub
End Sub
Private Sub cmcTreeSave_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub edcBrowserFile_Change()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    If Len(edcBrowserFile.Text) <= 0 Then
        Exit Sub
    End If
    slName = Trim$(edcBrowserFile.Text)
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcBrowserPath.Path & "\" & slFromFile
    End If
    Screen.MousePointer = vbHourglass
    ilRet = 0
    On Error GoTo edcBrowserFileErr:
    vbcBrowser.Value = vbcBrowser.Min
    pbcBrowserSample(1).Move 0, 0
    pbcBrowserSample(1).Picture = LoadPicture(slFromFile)
    'vbcBrowser.Top = 0
    vbcBrowser.Max = pbcBrowserSample(1).Height - pbcBrowserSample(0).Height
    vbcBrowser.Enabled = (pbcBrowserSample(0).Height < pbcBrowserSample(1).Height)
    If vbcBrowser.Enabled Then
        vbcBrowser.SmallChange = pbcBrowserSample(0).Height
        vbcBrowser.LargeChange = pbcBrowserSample(0).Height
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
edcBrowserFileErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcRptDescription_Change()
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub edcRptDescription_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcRptFormName_Change()
    imCatRptChg = True
    mSetCommands
End Sub

Private Sub edcRptFormName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcRptName_Change()
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub edcRptName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRptPassValue_Change()
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub edcRptPassValue_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTreeDescription_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub edcTreeDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    mInit
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    'If Command$ = "Debug" Then
    '    igPasswordOk = True
    '    imPasswordAsked = True
    'Else
    '    imPasswordAsked = False
    '    'CSPWord.Show vbModal
    'End If
    CSIPass.Show vbModal
    If Not igPasswordOk Then
        'imTerminate = True
        'mTerminate
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
    Set SpecRpt = Nothing   'Remove data segment
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    btrStopAppl
    End
End Sub

Private Sub hbcRptSample_Change()
    pbcRptSample(1).Left = -hbcRptSample.Value
End Sub

Private Sub imcPrt_Click()
    Dim ilLinesPerPage As Integer
    Dim ilCurrentLineNo As Integer
    Dim ilPageNo As Integer
    Dim slHeading As String
    Dim slRecord As String
    Dim ilLoop As Integer
    Dim ilTerminate As Integer
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.FontBold = True
    slHeading = "Report Tree Structure as " & Format$(gNow(), "m/d/yy h:mm:ss AM/PM") & "         "
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    ilCurrentLineNo = 0
    ilPageNo = 0
    ilTerminate = False
    On Error GoTo mStopPrt
    'Output heading line
    GoSub mHeading
    For ilLoop = 0 To lbcTree(1).ListCount - 1 Step 1
        slRecord = lbcTree(1).List(ilLoop)
        If ilTerminate Then
            Exit For
        End If
        GoSub mLineOutput
    Next ilLoop
    Printer.EndDoc
    On Error GoTo 0
    Exit Sub
mStopPrt:
    ilTerminate = True
    Return
mHeading:  'Output file name and date
    ilPageNo = ilPageNo + 1
    Printer.Print slHeading & "Page" & Str$(ilPageNo)
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Return
mLineOutput:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        ilCurrentLineNo = 0
        GoSub mHeading
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Return
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    Select Case imDragSrce
        Case DRAGTREE0
        Case DRAGTREE1
            'lbcTree(1).RemoveItem imDragIndexSrce
            lbcTree(1).ListIndex = imDragIndexSrce
            DoEvents
            cmcTreeDelete_Click
    End Select
    mClearDrag
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case imDragSrce
        Case DRAGTREE0
        Case DRAGTREE1
            If State = vbEnter Then    'Enter drag over
                lbcTree(1).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
                imcTrash.Picture = IconTraf!imcTrashOpened.Picture
            ElseIf State = vbLeave Then
                lbcTree(1).DragIcon = IconTraf!imcIconDrag.DragIcon
                imcTrash.Picture = IconTraf!imcTrashClosed.Picture
            End If
    End Select
End Sub
Private Sub lacTab_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub lbcBrowserFile_Click()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    If lbcBrowserFile.ListIndex < 0 Then
        Exit Sub
    End If
    slName = Trim$(lbcBrowserFile.List(lbcBrowserFile.ListIndex))
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcBrowserPath.Path & "\" & slFromFile
    End If
    Screen.MousePointer = vbHourglass
    ilRet = 0
    On Error GoTo lbcBrowserFileErr:
    vbcBrowser.Value = vbcBrowser.Min
    pbcBrowserSample(1).Move 0, 0
    pbcBrowserSample(1).Picture = LoadPicture(slFromFile)
    'vbcBrowser.Top = 0
    vbcBrowser.Max = pbcBrowserSample(1).Height - pbcBrowserSample(0).Height
    vbcBrowser.Enabled = (pbcBrowserSample(0).Height < pbcBrowserSample(1).Height)
    If vbcBrowser.Enabled Then
        vbcBrowser.SmallChange = pbcBrowserSample(0).Height
        vbcBrowser.LargeChange = pbcBrowserSample(0).Height
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
lbcBrowserFileErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub lbcBrowserFile_DblClick()
    cmcBrowserOpen_Click
End Sub
Private Sub lbcBrowserPath_Change()
    edcBrowserFile.Text = "*.SFI"
    lbcBrowserFile.fileName = lbcBrowserPath.Path & "\*.SFI"
End Sub
Private Sub lbcTree_Click(Index As Integer)
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilFound As Integer
    Dim slTestName As String
    Dim slType As String
    ilIndex = lbcTree(Index).ListIndex
    slName = lbcTree(Index).List(ilIndex)
    If Index = 0 Then
        For ilLoop = 0 To UBound(tgRnfList) - 1 Step 1
            'slNameCode = tgRnfList(ilLoop).sKey    'lbcMster.List(ilLoop)
            'ilRet = gParseItem(slNameCode, 1, "\", slTestName)
            'ilRet = gParseItem(slTestName, 1, "|", slType)
            'ilRet = gParseItem(slTestName, 2, "|", slTestName)
            slTestName = Trim$(tgRnfList(ilLoop).tRnf.sName) 'slTestName)
            slType = tgRnfList(ilLoop).tRnf.sType
            If slType = "C" Then
                slTestName = slTestName & "/Category"
            Else
                slTestName = slTestName & "/Report"
            End If
            If slName = slTestName Then
                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilCode = tgRnfList(ilLoop).tRnf.iCode
                ilFound = ilLoop
                Exit For
            End If
        Next ilLoop
        If ilFound = -1 Then
            Exit Sub
        End If
        tmRnfSrchKey.iCode = ilCode 'Val(slCode)
    Else
        If InStr(slName, "/Report") > 0 Then
            If tgRtfList(ilIndex).tRtf.sAutoReturn = "Y" Then
                rbcTreeAutoReturn(1).Value = True
            Else
                rbcTreeAutoReturn(0).Value = True
            End If
        Else
            rbcTreeAutoReturn(1).Value = False
            rbcTreeAutoReturn(0).Value = False
        End If
        tmRnfSrchKey.iCode = tgRtfList(ilIndex).tRtf.iRnfCode
    End If
    imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmRnf, tmRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        'If tmRnf.iStrLen > 0 Then
        '    edcTreeDescription.Text = Trim$(Left$(tmRnf.sDescription, tmRnf.iStrLen))
        'Else
        '    edcTreeDescription.Text = ""
        'End If
        edcTreeDescription.Text = gStripChr0(tmRnf.sDescription)
    End If
    lacTreeDescription.Caption = slName
    mSetCommands
End Sub
Private Sub lbcTree_DblClick(Index As Integer)
    If Index = 0 Then
        cmcTreeAdd_Click
    End If
End Sub
Private Sub lbcTree_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim ilListIndex As Integer
    If imDragDest = -1 Then
        mClearDrag
        Exit Sub
    End If
    Select Case imDragSrce
        Case DRAGTREE0
            lbcTree(0).ListIndex = imDragIndexSrce
            ilListIndex = Y \ fgListHtSerif825 + lbcTree(1).TopIndex
            If (ilListIndex >= 0) And (ilListIndex <= lbcTree(1).ListCount - 1) Then
                lbcTree(1).ListIndex = ilListIndex
                cmcTreeAdd_Click
            Else
                If ilListIndex >= lbcTree(1).ListCount - 1 Then
                    lbcTree(1).ListIndex = -1   'Add to end of list
                    cmcTreeAdd_Click
                End If
            End If
            mClearDrag
        Case DRAGTREE1
            mClearDrag
    End Select
End Sub
Private Sub lbcTree_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    imDragDest = -1
    If imDragSrce = DRAGTREE0 Then
        If Index = 0 Then
            lbcTree(0).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else
            If State = vbLeave Then
                lbcTree(0).DragIcon = IconTraf!imcIconDrag.DragIcon
                Exit Sub
            End If
            lbcTree(0).DragIcon = IconTraf!imcIconInsert.DragIcon
            imDragDest = DRAGTREE1
        End If
    ElseIf imDragSrce = DRAGTREE1 Then
        lbcTree(1).DragIcon = IconTraf!imcIconDrag.DragIcon
    End If
End Sub
Private Sub lbcTree_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        fmDragX = X
        fmDragY = Y
        imDragButton = Button
        imDragType = 0
        imDragShift = Shift
        imDragSrce = DRAGTREE0
        imDragIndexDest = -1
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    Else
        fmDragX = X
        fmDragY = Y
        imDragButton = Button
        imDragType = 0
        imDragShift = Shift
        imDragSrce = DRAGTREE1
        imDragIndexDest = -1
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    End If
End Sub
Private Sub lbcTree_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCatRptSaveRec                  *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mCatRptSaveRec() As Integer
'
'   iRet = mCatRptSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim slName As String    'Name
    Dim ilRet As Integer
    Dim slMsg As String
    Dim ilRnfRecLen As Integer
    Dim ilChg As Integer
    Dim tlRtf As RTF
    If Not mOKName() Then
        mCatRptSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilChg = False
    Do  'Loop until record updated or added
        mMoveCtrlToRec
        ilRnfRecLen = Len(tmRnf) '- Len(tmRnf.sDescription) + Len(Trim$(tmRnf.sDescription)) ' + 2  '100 = fixed record length; 2=Length value which is part of the variable record
        If imSelectedIndex = 0 Then 'New selected
            tmRnf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmRnf, tmRnf, ilRnfRecLen, INDEXKEY0)
            slMsg = "mCatRptSaveRec (btrInsert)"
        Else 'Old record-Update
            tmRnfSrchKey.iCode = tmRnf.iCode
            imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
            ilRet = btrGetEqual(hmRnf, tmTempRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If (tmTempRnf.sState <> tmRnf.sState) Or (tmTempRnf.sType <> tmRnf.sType) Then
                ilChg = True
            End If
            ilRnfRecLen = Len(tmRnf) '- Len(tmRnf.sDescription) + Len(Trim$(tmRnf.sDescription)) ' + 2  '100 = fixed record length; 2=Length value which is part of the variable record
            ilRet = btrUpdate(hmRnf, tmRnf, ilRnfRecLen)
            slMsg = "mCatRptSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mCatRptSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, SpecRpt
    On Error GoTo 0
    If ilChg Then
        'State
        For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
            If tgRtfList(ilLoop).tRtf.iRnfCode = tmRnf.iCode Then
                Do
                    ilRet = btrGetDirect(hmRtf, tlRtf, imRtfRecLen, tgRtfList(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    tmRtfSrchKey.iCode = tlRtf.iCode
                    ilRet = btrGetEqual(hmRtf, tlRtf, imRtfRecLen, tmRtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    tgRtfList(ilLoop).tRtf.sRnfType = tmRnf.sType
                    tgRtfList(ilLoop).tRtf.sRnfState = tmRnf.sState
                    ilRet = btrUpdate(hmRtf, tgRtfList(ilLoop).tRtf, imRtfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
            End If
            slMsg = "mSaveRec (btrUpdate RTF)"
            On Error GoTo mCatRptSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, SpecRpt
            On Error GoTo 0
        Next ilLoop
    End If
    imCatRptChg = False
    mCatRptSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mCatRptSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mCatRptSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    'rbcRptType(1).Value = True
    edcRptName.Text = ""
    'cbcRptExe.ListIndex = -1
    edcRptFormName.Text = ""
    'imRptExtSelectedIndex = -1
    imJobListSelectedIndex = -1
    edcRptDescription.Text = ""
    cbcRptSample.ListIndex = -1
    edcRptPassValue.Text = ""
    rbcRptMoney(0).Value = True
    rbcRptState(0).Value = True
    vbcRptSample.Value = vbcRptSample.Min
    hbcRptSample.Value = hbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    pbcRptSample(1).Picture = LoadPicture()
    imCatRptChg = False
    imTreeChg = False
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearDrag                      *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear drag when drop on illegal*
'*                      object                         *
'*                                                     *
'*******************************************************
Private Sub mClearDrag()
    tmcScroll.Enabled = False   'incase dropped within scroll area
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imcTrash.Enabled = False
    'imcTrash.Visible = False
    imDragIndexSrce = -1
    imDragIndexDest = -1
    imDragSrce = -1
    imDragScroll = -1
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDragOver                       *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process drag over              *
'*                                                     *
'*******************************************************
Private Sub mDragOver(ilObject As Integer, State As Integer, Index As Integer)
    If imDragSrce = DRAGTREE0 Then
        Exit Sub
    End If
    imDragScroll = ilObject
    If State = vbEnter Then
        tmcScroll.Enabled = True
        imScrollDir = Index
        Exit Sub
    End If
    If State = vbLeave Then
        tmcScroll.Enabled = False
        imDragScroll = -1
        Exit Sub
    End If
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
    Dim llNoRec As Long 'Number of records
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slFromFile As String
    Dim slCommand As String
    Dim ilSpace As Integer
    Dim slStr As String
    Dim ilTestSystem As Integer
    Dim slPath As String
    Screen.MousePointer = vbHourglass
    imTerminate = False
    mParseCmmdLine
    mInitBox
    lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    imDragIndexDest = -1
    imNoRowLbc = lbcTree(0).Height \ fgListHtArial825
    imcTrash.Visible = True 'False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imcPrt.Picture = IconTraf!imcCamera.Picture
    
    hmRnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRnf, "", sgDBPath & "RNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RNF.Btr)", SpecRpt
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)
    
    hmRtf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRtf, "", sgDBPath & "RTF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RTF.Btr)", SpecRpt
    On Error GoTo 0
    imRtfRecLen = Len(tmRtf)
    
    'plcTabSelection(0).Move 285, 300
    'plcTabSelection(1).Move plcTabSelection(0).left + plcTabSelection(0).Width + 30, plcTabSelection(0).top
    
    plcBrowser.Move 720, 675
    plcRptName.Move 285, 675
    plcTree.Move plcRptName.Left, plcRptName.Top ', plcRptName.Width
'    lncLine(0).X1 = plcTabSelection(0).left - plcRptName.left + 45
'    lncEdge(0).X1 = lncLine(0).X2 '+ 15
'    lncEdge(0).X2 = lncLine(0).X2 '+ 15
'    lncEdge(0).Y1 = 0
'    lncEdge(0).Y2 = 30
'    lncEdge(1).X1 = lncLine(0).X2 + 15 '+ 30
'    lncEdge(1).X2 = lncLine(0).X2 + 15 '+ 30
'    lncEdge(1).Y1 = 0
'    lncEdge(1).Y2 = 15
'    lncLine(1).X1 = plcTabSelection(1).left - plcRptName.left + 45
'    lncEdge(2).X1 = lncLine(1).X2 + 15
'    lncEdge(2).X2 = lncLine(1).X2 + 15
'    lncEdge(2).Y1 = 0
'    lncEdge(2).Y2 = 30
'    lncEdge(3).X1 = lncLine(1).X2 + 30
'    lncEdge(3).X2 = lncLine(1).X2 + 30
'    lncEdge(3).Y1 = 0
'    lncEdge(3).Y2 = 15
    
    lbcRptBmpNames.Path = Left$(sgRptPath, Len(sgRptPath) - 1)
    For ilLoop = 0 To lbcRptBmpNames.ListCount - 1 Step 1
        cbcRptSample.AddItem lbcRptBmpNames.List(ilLoop)
    Next ilLoop
    cbcRptSample.AddItem "[None]", 0
    'lbcFileNames.Pattern = "Rpt*.Exe"
    'lbcFileNames.Path = left$(sgExePath, Len(sgExePath) - 1)
    'For ilLoop = 0 To lbcFileNames.ListCount - 1 Step 1
    '    slStr = lbcFileNames.List(ilLoop)
    '    If StrComp("SpecRpt.Exe", slStr, 1) <> 0 Then
    '        cbcRptExe.AddItem slStr
    '    End If
    'Next ilLoop
    mPopulate
    gObtainRTF hmRtf, True
    imCatRptChg = False
    imTreeChg = False
    imChgMode = False
    imFirstFocus = True
    imSelectedIndex = -1
    'imRptExtSelectedIndex = -1
    imJobListSelectedIndex = -1
    cbcJobList.AddItem "[None]"
    cbcJobList.AddItem "[List]"
    cbcJobList.AddItem "Budgets"
    cbcJobList.AddItem "Rate Cards"
    cbcJobList.AddItem "Proposals"
    cbcJobList.AddItem "Orders"
    cbcJobList.AddItem "Programming"
    cbcJobList.AddItem "Spots"
    cbcJobList.AddItem "Copy"
    cbcJobList.AddItem "Logs"
    cbcJobList.AddItem "Post Log"
    cbcJobList.AddItem "Invoices"
    cbcJobList.AddItem "Collections"
    cbcJobList.AddItem "Commissions"
    gCenterStdAlone SpecRpt
    SpecRpt.Show
    'plcTabSelection_Click 0
    If plcTabSelection.SelectedItem.Index <> 1 Then
        'plcTabSelection.SelectedItem.Index = 1
        SendKeys "%N"
    Else
        plcTabSelection_Click
    End If
    imTerminate = False
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
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    If rbcRptType(0).Value Then
        tmRnf.sType = "C"
    Else
        tmRnf.sType = "R"
    End If
    'Name
    tmRnf.sName = edcRptName.Text
    If tmRnf.sType = "R" Then
        tmRnf.sRptExe = edcRptFormName.Text 'edcRptFormName.List(cbcRptExe.ListIndex)
    Else
        tmRnf.sRptExe = ""
    End If
    'tmRnf.iStrLen = Len(edcRptDescription.Text)
    'tmRnf.sDescription = Trim$(edcRptDescription.Text) & Chr$(0) & Chr$(0) 'sgTB
    tmRnf.sDescription = Trim$(edcRptDescription.Text) & Chr$(0) '& Chr$(0) 'sgTB
    If tmRnf.sType = "C" Then
        If imJobListSelectedIndex = 1 Then
            tmRnf.iJobListNo = -1
        ElseIf imJobListSelectedIndex > 1 Then
            Select Case cbcJobList.List(imJobListSelectedIndex)
                Case "Budgets"
                    tmRnf.iJobListNo = BUDGETSJOB
                Case "Rate Cards"
                    tmRnf.iJobListNo = RATECARDSJOB
                Case "Proposals"
                    tmRnf.iJobListNo = PROPOSALSJOB
                Case "Orders"
                    tmRnf.iJobListNo = CONTRACTSJOB
                Case "Programming"
                    tmRnf.iJobListNo = PROGRAMMINGJOB
                Case "Spots"
                    tmRnf.iJobListNo = SPOTSJOB
                Case "Copy"
                    tmRnf.iJobListNo = COPYJOB
                Case "Logs"
                    tmRnf.iJobListNo = LOGSJOB
                Case "Post Log"
                    tmRnf.iJobListNo = POSTLOGSJOB
                Case "Invoices"
                    tmRnf.iJobListNo = INVOICESJOB
                Case "Collections"
                    tmRnf.iJobListNo = COLLECTIONSJOB
                Case "Commissions"
                    tmRnf.iJobListNo = SLSPCOMMSJOB
            End Select
        Else
            tmRnf.iJobListNo = -99
        End If
    Else
        tmRnf.iJobListNo = -99
    End If
    'Sample
    If tmRnf.sType = "R" Then
        tmRnf.sRptSample = cbcRptSample.List(cbcRptSample.ListIndex)
    Else
        tmRnf.sRptSample = ""
    End If
    If tmRnf.sType = "R" Then
        tmRnf.sPassValue = edcRptPassValue.Text
    Else
        tmRnf.sPassValue = ""
    End If
    If tmRnf.sType = "R" Then
        If rbcRptMoney(0).Value Then
            tmRnf.sMoneyShown = "Y"
        Else
            tmRnf.sMoneyShown = "N"
        End If
    Else
        tmRnf.sMoneyShown = ""
    End If
    If rbcRptState(1).Value Then
        tmRnf.sState = "D"
    Else
        tmRnf.sState = "A"
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim slStr As String
    If tmRnf.sType = "C" Then
        rbcRptType(0).Value = True
    Else
        rbcRptType(1).Value = True
    End If
    edcRptName.Text = Trim$(tmRnf.sName)
    If tmRnf.sType = "R" Then
        slStr = Trim$(tmRnf.sRptExe)
        'gFindMatch slStr, 0, cbcRptExe
        'If gLastFound(cbcRptExe) < 0 Then
        '    cbcRptExe.ListIndex = -1
        'Else
        '    cbcRptExe.ListIndex = gLastFound(cbcRptExe)
        '    imRptExtSelectedIndex = cbcRptExe.ListIndex
        'End If
        edcRptFormName.Text = slStr
    Else
        'cbcRptExe.ListIndex = -1
        edcRptFormName.Text = ""
    End If
    'If tmRnf.iStrLen > 0 Then
    '    edcRptDescription.Text = Trim$(Left$(tmRnf.sDescription, tmRnf.iStrLen))
    'Else
    '    edcRptDescription.Text = ""
    'End If
    edcRptDescription.Text = gStripChr0(tmRnf.sDescription)
    If tmRnf.sType = "C" Then
        Select Case tmRnf.iJobListNo
            Case -1
                slStr = "[List]"
            Case BUDGETSJOB
                slStr = "Budgets"
            Case RATECARDSJOB
                slStr = "Rate Cards"
            Case PROPOSALSJOB
                slStr = "Proposals"
            Case CONTRACTSJOB
                slStr = "Orders"
            Case PROGRAMMINGJOB
                slStr = "Programming"
            Case SPOTSJOB
                slStr = "Spots"
            Case COPYJOB
                slStr = "Copy"
            Case LOGSJOB
                slStr = "Logs"
            Case POSTLOGSJOB
                slStr = "Post Log"
            Case INVOICESJOB
                slStr = "Invoices"
            Case COLLECTIONSJOB
                slStr = "Collections"
            Case SLSPCOMMSJOB
                slStr = "Commissions"
            Case Else
                slStr = "[None]"
        End Select
        gFindMatch slStr, 0, cbcJobList
        If gLastFound(cbcJobList) < 0 Then
            cbcJobList.ListIndex = -1
        Else
            cbcJobList.ListIndex = gLastFound(cbcJobList)
            imJobListSelectedIndex = cbcJobList.ListIndex
        End If
    Else
        cbcJobList.ListIndex = -1
    End If
    If tmRnf.sType = "R" Then
        slStr = Trim$(tmRnf.sRptSample)
        gFindMatch slStr, 0, cbcRptSample
        If gLastFound(cbcRptSample) < 0 Then
            cbcRptSample.ListIndex = -1
        Else
            cbcRptSample.ListIndex = gLastFound(cbcRptSample)
        End If
    Else
        cbcRptSample.ListIndex = -1
    End If
    If tmRnf.sType = "R" Then
        edcRptPassValue.Text = Trim$(tmRnf.sPassValue)
        If tmRnf.sMoneyShown = "Y" Then
            rbcRptMoney(0).Value = True
        Else
            rbcRptMoney(1).Value = False
        End If
    Else
        edcRptPassValue.Text = ""
        rbcRptMoney(0).Value = False
        rbcRptMoney(1).Value = False
    End If
    If tmRnf.sState = "A" Then
        rbcRptState(0).Value = True
    Else
        rbcRptState(1).Value = True
    End If
    imCatRptChg = False
    imTreeChg = False
End Sub
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
    Dim slStr As String
    If edcRptName.Text <> "" Then    'Test name
        slStr = edcRptName.Text
        If rbcRptType(0).Value Then
            slStr = slStr & "/Category"
        Else
            slStr = slStr & "/Report"
        End If
        gFindMatch slStr, 0, cbcCatRpt    'Determine if name exist
        If gLastFound(cbcCatRpt) <> -1 Then   'Name found
            If gLastFound(cbcCatRpt) <> imSelectedIndex Then
                If edcRptName.Text = cbcCatRpt.List(gLastFound(cbcCatRpt)) Then
                    Beep
                    MsgBox "Name already defined, enter a different name", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
                    edcRptName.Text = Trim$(tmRnf.sName) 'Reset text
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
End Function
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
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    slCommand = Command$
    ilPos = InStr(1, slCommand, "/CS", 1)
    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
        igStdAloneMode = True 'False  'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
        ilPos = InStr(1, slCommand, "/S:Test", 1)
        If ilPos > 0 Then
            ilTestSystem = True
        End If
        ilPos = InStr(1, slCommand, "/S:Prod", 1)
        If ilPos > 0 Then
            ilTestSystem = False
        End If
        slCommand = "Traffic\Guide"
    Else
        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        'imShowHelpMsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpMsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    End If
    gInitStdAlone SpecRpt, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igSiteCallSource = Val(slStr)
    If igStdAloneMode Then
        igSiteCallSource = CALLNONE
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim slName As String
    Dim slNameCode As String
    Dim llRecPos As Long
    Dim llLen As Long
    Dim ilLoop As Integer
    Dim ilRnfRecLen As Integer
    
    cbcCatRpt.Clear
    gObtainRNF hmRnf
    llLen = 0
    For ilLoop = 0 To UBound(tgRnfList) - 1 Step 1
        'slNameCode = tgRnfList(ilLoop).sKey    'lbcMster.List(ilLoop)
        'ilRet = gParseItem(slNameCode, 1, "\", slName)
        'ilRet = gParseItem(slName, 1, "|", slType)
        'ilRet = gParseItem(slName, 2, "|", slName)
        slName = Trim$(tgRnfList(ilLoop).tRnf.sName)
        slType = tgRnfList(ilLoop).tRnf.sType
        slName = Trim$(slName)
        If slType = "C" Then
            slName = slName & "/Category"
        Else
            slName = slName & "/Report"
        End If
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        cbcCatRpt.AddItem slName  'Add ID to list box
    Next ilLoop
    cbcCatRpt.AddItem "[New]", 0
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer) As Integer
'
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    
    slNameCode = tgRnfList(ilSelectIndex - 1).sKey   'lbcTitleCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", SpecRpt
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmRnfSrchKey.iCode = CLng(slCode)
    imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmRnf, tmRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", SpecRpt
    On Error GoTo 0
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRtfPop                         *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mRTFPop()
'
'   mRtfPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llRecPos As Long
    Dim llLen As Long
    Dim ilLoop As Integer
    Dim ilLevel As Integer
    Dim ilRnf As Integer
    
    lbcTree(1).Clear
    'gObtainRTF hmRtf, True
    llLen = 0
    For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            'slNameCode = tgRnfList(ilRnf).sKey    'lbcMster.List(ilLoop)
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'If Val(slCode) = tgRtfList(ilLoop).tRtf.iRnfCode Then
            If tgRtfList(ilLoop).tRtf.iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                'ilRet = gParseItem(slNameCode, 1, "\", slName)
                'ilRet = gParseItem(slName, 1, "|", slType)
                'ilRet = gParseItem(slName, 2, "|", slName)
                slName = Trim$(tgRnfList(ilRnf).tRnf.sName) 'slName)
                slType = tgRnfList(ilRnf).tRnf.sType
                If slType = "C" Then
                    slName = slName & "/Category"
                Else
                    slName = slName & "/Report"
                End If
                For ilLevel = 1 To tgRtfList(ilLoop).tRtf.iLevel - 1 Step 1
                    slName = "  " & slName
                Next ilLevel
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    Exit For
                End If
                lbcTree(1).AddItem slName  'Add ID to list box
                Exit For
            End If
        Next ilRnf
    Next ilLoop
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSendHelpMsg                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Send message if not stand alone*
'*                                                     *
'*******************************************************
Private Sub mSendHelpMsg(slMsg As String)
    Dim ilLoop As Integer
    If (Not igStdAloneMode) And (imShowHelpMsg) Then
        edcLinkDestHelpMsg.LinkExecute slMsg
        For ilLoop = 1 To igDDEDelay Step 1
            DoEvents
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Buttons                     *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim slStr As String
    Dim slName As String
    If Not igPasswordOk Then
        cmcRptUndo.Enabled = False
        cmcErase.Enabled = False
        cmcRptSave.Enabled = False
        cmcTreeSave.Enabled = False
        Exit Sub
    End If
    If imTabSelection = 0 Then
        If imCatRptChg Then
            'plcTabSelection(0).Enabled = False
            'plcTabSelection(1).Enabled = False
            plcTabSelection.Enabled = False
            cmcRptUndo.Enabled = True
            cmcErase.Enabled = False
            If rbcRptType(0).Value Then 'Category
                slStr = Trim$(edcRptName.Text)
                If Len(slStr) = 0 Then
                    cmcRptSave.Enabled = False
                    Exit Sub
                End If
            Else
                slStr = Trim$(edcRptName.Text)
                If Len(slStr) = 0 Then
                    cmcRptSave.Enabled = False
                    Exit Sub
                End If
                'If cbcRptExe.ListIndex <= -1 Then
                '    cmcRptSave.Enabled = False
                '    Exit Sub
                'End If
                If Trim$(edcRptFormName.Text) = "" Then
                    cmcRptSave.Enabled = False
                    Exit Sub
                End If
            End If
            cmcRptSave.Enabled = True
        Else
            cmcRptUndo.Enabled = False
            If imSelectedIndex > 0 Then
                cmcErase.Enabled = True
            Else
                cmcErase.Enabled = False
            End If
            cmcRptSave.Enabled = False
            'plcTabSelection(0).Enabled = True
            'plcTabSelection(1).Enabled = True
            plcTabSelection.Enabled = True
        End If
    Else
        If imTreeChg Then
            'plcTabSelection(0).Enabled = False
            'plcTabSelection(1).Enabled = False
            plcTabSelection.Enabled = False
            cmcTreeSave.Enabled = True
        Else
            'plcTabSelection(0).Enabled = True
            'plcTabSelection(1).Enabled = True
            plcTabSelection.Enabled = True
            cmcTreeSave.Enabled = False
        End If
        If lbcTree(1).ListIndex >= 0 Then
            slName = lbcTree(1).List(lbcTree(1).ListIndex)
            If InStr(slName, "/Report") > 0 Then
                rbcTreeAutoReturn(0).Enabled = True
                rbcTreeAutoReturn(1).Enabled = True
                pbcSubTab.Enabled = False
                pbcAddTab.Enabled = False
            Else
                rbcTreeAutoReturn(0).Enabled = False
                rbcTreeAutoReturn(1).Enabled = False
                pbcSubTab.Enabled = True
                pbcAddTab.Enabled = True
            End If
        End If
    End If
    'If ((Not imCatRptChg) Or (cmcRptSave.Enabled)) And ((Not imTreeChg) Or (cmcTreeSave.Enabled)) Then
    '    If imCatRptChg Or imTreeChg Then
    '        cmcSave.Enabled = True
    '    Else
    '        cmcSave.Enabled = False
    '    End If
    'Else
    '    cmcSave.Enabled = False
    'End If
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
    Dim ilRet As Integer
    Erase tgRnfList
    Erase tgRtfList
    Erase tgDelRtfList
    ilRet = btrClose(hmRtf)
    btrDestroy hmRtf
    ilRet = btrClose(hmRnf)
    btrDestroy hmRnf
    igParentRestarted = False
    If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & Trim$(Str$(igSiteCallSource))
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute Trim$(Str$(igSiteCallSource))
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    End If
    Screen.MousePointer = vbDefault
    Unload IconTraf
    igManUnload = YES
    Unload SpecRpt
    Set SpecRpt = Nothing   'Remove data segment
    igManUnload = NO
End Sub
Private Sub pbcAddTab_Click()
    Dim slName As String
    Dim ilBlkCount As Integer
    Dim ilTBlkCount As Integer
    Dim slChar As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    
    If lbcTree(1).ListIndex < 0 Then
        Exit Sub
    End If
    ilIndex = lbcTree(1).ListIndex
    slName = lbcTree(1).List(ilIndex)
    If InStr(slName, "/Category") > 0 Then
        ilBlkCount = 0
        For ilLoop = 1 To Len(slName) Step 1
            slChar = Mid$(slName, ilLoop, 1)
            If slChar = " " Then
                ilBlkCount = ilBlkCount + 1
            Else
                Exit For
            End If
        Next ilLoop
        slName = "  " & slName
        lbcTree(1).List(ilIndex) = slName
        ilIndex = ilIndex + 1
        Do While ilIndex < lbcTree(1).ListCount
            slName = lbcTree(1).List(ilIndex)
            ilTBlkCount = 0
            For ilLoop = 1 To Len(slName) Step 1
                slChar = Mid$(slName, ilLoop, 1)
                If slChar = " " Then
                    ilTBlkCount = ilTBlkCount + 1
                Else
                    Exit For
                End If
            Next ilLoop
            If ilTBlkCount <= ilBlkCount Then
                Exit Do
            End If
            slName = "  " & slName
            lbcTree(1).List(ilIndex) = slName
            ilIndex = ilIndex + 1
        Loop
    End If
End Sub
Private Sub pbcAddTab_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub pbcDnMove_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub pbcSubTab_Click()
    Dim slName As String
    Dim ilBlkCount As Integer
    Dim ilTBlkCount As Integer
    Dim slChar As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    
    If lbcTree(1).ListIndex < 0 Then
        Exit Sub
    End If
    ilIndex = lbcTree(1).ListIndex
    slName = lbcTree(1).List(ilIndex)
    If InStr(slName, "/Category") > 0 Then
        ilBlkCount = 0
        For ilLoop = 1 To Len(slName) Step 1
            slChar = Mid$(slName, ilLoop, 1)
            If slChar = " " Then
                ilBlkCount = ilBlkCount + 1
            Else
                Exit For
            End If
        Next ilLoop
        If ilBlkCount <= 1 Then
            Exit Sub
        End If
        slName = right$(slName, Len(slName) - 2)
        lbcTree(1).List(ilIndex) = slName
        ilIndex = ilIndex + 1
        Do While ilIndex < lbcTree(1).ListCount
            slName = lbcTree(1).List(ilIndex)
            ilTBlkCount = 0
            For ilLoop = 1 To Len(slName) Step 1
                slChar = Mid$(slName, ilLoop, 1)
                If slChar = " " Then
                    ilTBlkCount = ilTBlkCount + 1
                Else
                    Exit For
                End If
            Next ilLoop
            If ilTBlkCount <= ilBlkCount Then
                Exit Do
            End If
            slName = right$(slName, Len(slName) - 2)
            lbcTree(1).List(ilIndex) = slName
            ilIndex = ilIndex + 1
        Loop
    End If
End Sub
Private Sub pbcSubTab_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcTabSelection_Click()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slType As String
    Dim slState As String
    Dim slName As String
    Dim slNameCode As String
    Dim ilRnf As Integer
    Dim llLen As Long
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    'plcTabSelection(Index).ZOrder vbBringToFront
    
    ilIndex = plcTabSelection.SelectedItem.Index - 1
    If ilIndex = 0 Then   'Category/Report
        plcTree.Visible = False
        plcRptName.Visible = True
        plcRptName.ZOrder vbBringToFront
    ElseIf ilIndex = 1 Then   'Tree
        Screen.MousePointer = vbHourglass
        plcRptName.Visible = False
        plcTree.Visible = True
        plcTree.ZOrder vbBringToFront
        lbcTree(0).Clear
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            'slNameCode = tgRnfList(ilRnf).sKey    'lbcMster.List(ilLoop)
            'ilRet = gParseItem(slNameCode, 3, "|", slState)
            slState = tgRnfList(ilRnf).tRnf.sState
            If slState <> "D" Then
                'ilRet = gParseItem(slNameCode, 1, "\", slName)
                'ilRet = gParseItem(slName, 1, "|", slType)
                'ilRet = gParseItem(slName, 2, "|", slName)
                slName = Trim$(tgRnfList(ilRnf).tRnf.sName)    'slName)
                slType = tgRnfList(ilRnf).tRnf.sType
                If slType = "C" Then
                    slName = slName & "/Category"
                Else
                    slName = slName & "/Report"
                End If
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    Exit For
                End If
                lbcTree(0).AddItem slName  'Add ID to list box
            End If
        Next ilRnf
        mRTFPop
        Screen.MousePointer = vbDefault
    End If
    imTabSelection = ilIndex
End Sub
Private Sub plcTabSelection_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcTree_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcTreeAutoReturn_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub rbcRptMoney_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcRptMoney(Index).Value
    'End of coded added
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub rbcRptState_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcRptState(Index).Value
    'End of coded added
    imCatRptChg = True
    mSetCommands
End Sub
Private Sub rbcRptType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcRptType(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then   'Category
            'cbcRptExe.Enabled = False
            edcRptFormName.Enabled = False
            cbcRptSample.Enabled = False
            edcRptPassValue.Enabled = False
            rbcRptMoney(0).Enabled = False
            rbcRptMoney(1).Enabled = False
            pbcBrowserSample(1).Picture = LoadPicture()
            cbcJobList.Enabled = True
        Else    'Report
            'cbcRptExe.Enabled = True
            edcRptFormName.Enabled = True
            cbcRptSample.Enabled = True
            edcRptPassValue.Enabled = True
            rbcRptMoney(0).Enabled = True
            rbcRptMoney(1).Enabled = True
            cbcJobList.Enabled = False
        End If
    End If
End Sub
Private Sub rbcTreeAutoReturn_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcTreeAutoReturn(Index).Value
    'End of coded added
    Dim slName As String
    If lbcTree(1).ListIndex < 0 Then
        Exit Sub
    End If
    If Value Then
        slName = lbcTree(Index).List(lbcTree(1).ListIndex)
        If InStr(slName, "/Report") > 0 Then
            If Index = 0 Then
                tgRtfList(lbcTree(1).ListIndex).tRtf.sAutoReturn = "N"
            Else
                tgRtfList(lbcTree(1).ListIndex).tRtf.sAutoReturn = "Y"
            End If
        End If
    End If
End Sub
Private Sub rbcTreeAutoReturn_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub tmcDrag_Timer()
    Dim ilListIndex As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilLen As Integer
    tmcDrag.Enabled = False
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            If imDragButton <> 1 Then
                Exit Sub
            End If
            Select Case imDragSrce
                Case DRAGTREE0
                    ilListIndex = fmDragY \ fgListHtSerif825 + lbcTree(0).TopIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcTree(0).ListCount - 1) Then
                        lbcTree(0).ListIndex = ilListIndex
                        lbcTree(0).DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcTree(0).Drag vbBeginDrag
                        imcTrash.Enabled = False
                    Else
                        lbcTree(0).ListIndex = -1
                    End If
                Case DRAGTREE1
                    ilListIndex = fmDragY \ fgListHtArial825 + lbcTree(1).TopIndex
                    If (ilListIndex <= lbcTree(1).ListCount - 2) Then
                        lbcTree(1).ListIndex = ilListIndex
                        lbcTree(1).DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcTree(1).Drag vbBeginDrag
                        imcTrash.Visible = True
                        imcTrash.Enabled = True
                    Else
                        lbcTree(1).ListIndex = -1
                    End If
            End Select
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub tmcScroll_Timer()
    If imDragSrce = -1 Then
        tmcScroll.Enabled = False
        Exit Sub
    End If
    If imDragScroll = DRAGTREE1 Then
        If imScrollDir = 0 Then 'Page up
           If lbcTree(1).TopIndex > 0 Then
              lbcTree(1).TopIndex = lbcTree(1).TopIndex - 1
           End If
        Else    'Page down
           If lbcTree(1).TopIndex < lbcTree(1).ListCount - imNoRowLbc Then
              lbcTree(1).TopIndex = lbcTree(1).TopIndex + 1
           End If
        End If
    End If
End Sub
Private Sub vbcBrowser_Change()
    pbcBrowserSample(1).Top = -vbcBrowser.Value
End Sub
Private Sub vbcRptSample_Change()
    pbcRptSample(1).Top = -vbcRptSample.Value
End Sub
Private Sub plcRptState_Paint()
    plcRptState.CurrentX = 0
    plcRptState.CurrentY = 0
    plcRptState.Print "State"
End Sub
Private Sub plcRptMoney_Paint()
    plcRptMoney.CurrentX = 0
    plcRptMoney.CurrentY = 0
    plcRptMoney.Print "Money Shown on Report"
End Sub
Private Sub plcRptType_Paint()
    plcRptType.CurrentX = 0
    plcRptType.CurrentY = 0
    plcRptType.Print "Type"
End Sub
Private Sub plcTreeAutoReturn_Paint()
    plcTreeAutoReturn.CurrentX = 0
    plcTreeAutoReturn.CurrentY = 0
    plcTreeAutoReturn.Print "Allow User to Select Another Report"
End Sub
