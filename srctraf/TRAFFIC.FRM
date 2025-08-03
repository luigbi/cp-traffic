VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.MDIForm Traffic 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "CSI Traffic"
   ClientHeight    =   5385
   ClientLeft      =   255
   ClientTop       =   2520
   ClientWidth     =   11910
   Icon            =   "TRAFFIC.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "HelpMsg"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbcList 
      Align           =   1  'Align Top
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
      Height          =   1980
      Left            =   0
      ScaleHeight     =   1950
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   330
      Visible         =   0   'False
      Width           =   11910
      Begin VB.TextBox edcMsgFromAffiliate 
         Height          =   345
         Left            =   10590
         TabIndex        =   35
         Top             =   915
         Width           =   735
      End
      Begin VB.Timer tmcMonitor 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   4065
         Top             =   -15
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   555
         Left            =   9645
         TabIndex        =   32
         Top             =   270
         Visible         =   0   'False
         Width           =   420
         ExtentX         =   741
         ExtentY         =   979
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   795
         Left            =   9570
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1402
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"TRAFFIC.frx":08CA
      End
      Begin VB.Timer tmcTerminate 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4755
         Top             =   135
      End
      Begin VB.Timer tmcInitAlert 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5715
         Top             =   180
      End
      Begin VB.PictureBox pbc256List 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   8715
         Picture         =   "TRAFFIC.frx":0944
         ScaleHeight     =   480
         ScaleWidth      =   465
         TabIndex        =   29
         Top             =   1125
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox pbc256Job 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   8775
         Picture         =   "TRAFFIC.frx":DC0C
         ScaleHeight     =   405
         ScaleWidth      =   480
         TabIndex        =   28
         Top             =   630
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox pbc256SignOn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   8820
         Picture         =   "TRAFFIC.frx":18438
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   27
         Top             =   15
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Timer tmcFlashAlert 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   7215
         Top             =   30
      End
      Begin VB.FileListBox lbcZeroBTR 
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
         Height          =   225
         Left            =   5820
         MultiSelect     =   2  'Extended
         Pattern         =   "*.btr"
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1305
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.PictureBox pbcTestFont 
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
         Height          =   165
         Left            =   7485
         ScaleHeight     =   135
         ScaleWidth      =   1335
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1365
      End
      Begin VB.FileListBox lbcFile 
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
         Height          =   225
         Left            =   6465
         Pattern         =   "*.exe"
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1425
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox edcPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   6240
         PasswordChar    =   "*"
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   2115
      End
      Begin MSMask.MaskEdBox mkcTempFaxNo 
         Height          =   210
         Left            =   3510
         TabIndex        =   7
         Top             =   1635
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   16776960
         ForeColor       =   0
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(AAA) AAA-AAAA"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mkcTempPhoneNo 
         Height          =   210
         Left            =   1515
         TabIndex        =   6
         Top             =   1725
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
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
      Begin VB.PictureBox pbcSymb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   675
         Left            =   1215
         ScaleHeight     =   675
         ScaleWidth      =   690
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   690
      End
      Begin VB.Timer tmcCheckAlert 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   7770
         Top             =   405
      End
      Begin VB.TextBox edcLinkSrceHelpMsg 
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
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSMask.MaskEdBox mkcPhone 
         Height          =   210
         Left            =   1545
         TabIndex        =   2
         Top             =   1515
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
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
      Begin VB.PictureBox pbcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   675
         Left            =   2745
         ScaleHeight     =   675
         ScaleWidth      =   2985
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   2985
         Begin VB.Timer tmcWebConnectIssue 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   2520
            Top             =   0
         End
         Begin VB.PictureBox pbcArial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1485
            ScaleHeight     =   225
            ScaleWidth      =   1005
            TabIndex        =   11
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Timer tmcClock 
         Interval        =   60000
         Left            =   7740
         Top             =   -60
      End
      Begin MSComDlg.CommonDialog cdcSetup 
         Left            =   6480
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin ComctlLib.ImageList imcTabColor 
         Left            =   8190
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   91
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   20
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":1CB5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":1DCEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":1EE80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":20012
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":211A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":22336
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":234C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2465A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":257EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2697E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":27B10
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":28CA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":29E34
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2AFC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2C158
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2D2EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2E47C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":2F60E
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":307A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "TRAFFIC.frx":31932
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pbcMsgArea 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11910
      TabIndex        =   3
      Top             =   0
      Width           =   11910
      Begin VB.CommandButton cmcAffiliate 
         Caption         =   "AFFILIATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9360
         TabIndex        =   34
         Top             =   15
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton cmcEMail 
         Caption         =   "E-Mail History"
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
         Left            =   10905
         TabIndex        =   33
         ToolTipText     =   "Display E-Mail History"
         Top             =   15
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton cmcFind 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8700
         TabIndex        =   30
         ToolTipText     =   "Analyze Order"
         Top             =   15
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "COMM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   8055
         TabIndex        =   24
         ToolTipText     =   "Define Salesperson Commissions"
         Top             =   15
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "COLLECT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   7245
         TabIndex        =   23
         ToolTipText     =   "Post Payments"
         Top             =   15
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "INVOICE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   6480
         TabIndex        =   22
         ToolTipText     =   "Generate Invoices"
         Top             =   15
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   5940
         TabIndex        =   21
         ToolTipText     =   "Post Exceptions from Log"
         Top             =   15
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "FEED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   5400
         TabIndex        =   20
         ToolTipText     =   "Produce Station Feed"
         Top             =   15
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "LOG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   4950
         TabIndex        =   19
         ToolTipText     =   "Generate Logs"
         Top             =   15
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "COPY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   4380
         TabIndex        =   18
         ToolTipText     =   "Define Copy Rotations"
         Top             =   15
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "SPOT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   3840
         TabIndex        =   17
         ToolTipText     =   "Manipulate Spots"
         Top             =   15
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "PROGRAM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2925
         TabIndex        =   16
         ToolTipText     =   "Define Programming"
         Top             =   15
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2250
         TabIndex        =   15
         ToolTipText     =   "Enter Orders"
         Top             =   15
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "PROPOSAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "Enter Proposals"
         Top             =   15
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   765
         TabIndex        =   13
         ToolTipText     =   "Define Rate Card"
         Top             =   15
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmcTask 
         Caption         =   "BUDGET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Define Budgets"
         Top             =   15
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmcReport 
         Caption         =   "&Reports"
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
         Left            =   10110
         TabIndex        =   25
         ToolTipText     =   "Generate Reports"
         Top             =   15
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Task"
      Begin VB.Menu mnuTaskJob 
         Caption         =   "&Jobs"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuTaskList 
         Caption         =   "&Lists"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuTaskSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTqskBackup 
         Caption         =   "&Backup"
         Begin VB.Menu mnuBackupTrafficAndAffiliate 
            Caption         =   "Complete &System Backup"
         End
         Begin VB.Menu mnuTrafficDataToTest 
            Caption         =   "&Copy Data to Test"
         End
      End
      Begin VB.Menu mnuResetWorkFiles 
         Caption         =   "&Reset Work Files"
      End
      Begin VB.Menu mnuRemoteUser 
         Caption         =   "&Remote User"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTaskSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuImportAuto 
            Caption         =   "&Automation"
         End
         Begin VB.Menu mnuImportMegaphone 
            Caption         =   "&Megaphone Impression"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuImportResearch 
            Caption         =   "&Research"
            Begin VB.Menu mnuImportMark 
               Caption         =   "&Act 1"
            End
            Begin VB.Menu mnuImportIMS 
               Caption         =   "&CSV"
            End
            Begin VB.Menu mnuImportImpressions 
               Caption         =   "&Impressions"
            End
            Begin VB.Menu mnuImportRadar 
               Caption         =   "&RADAR"
            End
            Begin VB.Menu mnuImportSatellite 
               Caption         =   "&Satellite"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuImportVehicle 
               Caption         =   "&Vehicle"
            End
         End
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuExportAccount 
            Caption         =   "Accoun&ting"
         End
         Begin VB.Menu mnuExportAirWave 
            Caption         =   "Air&Wave"
         End
         Begin VB.Menu mnuExportAdServerBillDisc 
            Caption         =   "Ad Server Billing Discrepancy"
         End
         Begin VB.Menu mnuISCIXReference 
            Caption         =   "Audio ISCI Title"
         End
         Begin VB.Menu mnuExportMP2 
            Caption         =   "Audio MP&2"
         End
         Begin VB.Menu mnuExportAuto 
            Caption         =   "&Automation"
         End
         Begin VB.Menu mnuBarter 
            Caption         =   "&Barter Payment"
         End
         Begin VB.Menu mnuExportCarts 
            Caption         =   "Carts"
         End
         Begin VB.Menu mnuExportCash 
            Caption         =   "&Cash Receipts"
         End
         Begin VB.Menu mnuCNC 
            Caption         =   "&Clearance n Compensation"
            Begin VB.Menu mnuCNCAP 
               Caption         =   "&Advertiser Products"
            End
            Begin VB.Menu mnuCNCNI 
               Caption         =   "&Network Inventory"
            End
            Begin VB.Menu mnuCNCSA 
               Caption         =   "&Selling to Airing Links"
            End
            Begin VB.Menu mnuCNCSS 
               Caption         =   "S&pot Summary"
            End
         End
         Begin VB.Menu mnuExpCmmlChg 
            Caption         =   "C&ommercial Change"
         End
         Begin VB.Menu mnuExpCntrLine 
            Caption         =   "Contract &Line"
         End
         Begin VB.Menu mnuExportBkCpy 
            Caption         =   "Copy &Bulk Feed"
         End
         Begin VB.Menu mnuExportCorporate 
            Caption         =   "&Corporate Export"
         End
         Begin VB.Menu mnuExportCRE 
            Caption         =   "C&ustom Revenue"
         End
         Begin VB.Menu mnuExportDallas 
            Caption         =   "Dallas &Feed"
         End
         Begin VB.Menu mnuEfficio 
            Caption         =   "E&fficio"
            Enabled         =   0   'False
            Begin VB.Menu mnuEfficioProj 
               Caption         =   "&Projections"
            End
            Begin VB.Menu mnuEfficioRev 
               Caption         =   "&Revenue"
            End
         End
         Begin VB.Menu mnuExportEnco 
            Caption         =   "Enco"
         End
         Begin VB.Menu mnuExportEngr 
            Caption         =   "&Engineering Feed"
            Begin VB.Menu mnuExportNYASP 
               Caption         =   "&ASP (with ISCI)"
            End
            Begin VB.Menu mnuExportNYESPN 
               Caption         =   "E&SPN (with ISCI)"
            End
         End
         Begin VB.Menu mnuGetPaid 
            Caption         =   "&GetPaid"
         End
         Begin VB.Menu mnuGreatPlainsGL 
            Caption         =   "Great&Plains G/L"
         End
         Begin VB.Menu mnuExportInvReg 
            Caption         =   "&Invoice Register"
         End
         Begin VB.Menu mnuExportMatrix 
            Caption         =   "&Matrix"
         End
         Begin VB.Menu mnuExportMK 
            Caption         =   "Miller &Kaplan"
         End
         Begin VB.Menu mnuExportPhoenix 
            Caption         =   "Phoeni&x Log"
         End
         Begin VB.Menu mnuExportRAB 
            Caption         =   "&RAB"
         End
         Begin VB.Menu mnuExportRevenue 
            Caption         =   "&Revenue"
         End
         Begin VB.Menu mnuExportStnFeed 
            Caption         =   "Station Feed"
         End
         Begin VB.Menu mnuExportTableau 
            Caption         =   "&Tableau"
         End
         Begin VB.Menu mnuExportWOInvoice 
            Caption         =   "&WO Invoice"
         End
      End
      Begin VB.Menu mnuTaskSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaskExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAccessory 
      Caption         =   "&Accessories"
      Begin VB.Menu mnuAccessoryCalendar 
         Caption         =   "&Calendar                        F4"
      End
      Begin VB.Menu mnuAccessoryCalculator 
         Caption         =   "&Calculator                      F5"
      End
      Begin VB.Menu mnuAccessoryEditor 
         Caption         =   "Edi&tor"
      End
      Begin VB.Menu mnuAccessoryDate 
         Caption         =   "Counterpoint &Date"
      End
      Begin VB.Menu mnuAccessorySpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessorySetCredit 
         Caption         =   "Set &Credit"
      End
      Begin VB.Menu mnuAccessoryShowCredits 
         Caption         =   "&Show Credit Alerts"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuAccessorySpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryUtilities 
         Caption         =   "&Utilities"
      End
      Begin VB.Menu mnuAccessorySpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryMessages 
         Caption         =   "&Messages"
      End
      Begin VB.Menu mnuAccessorySpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryBlocks 
         Caption         =   "&View Blocks"
      End
      Begin VB.Menu mnuAccessorySpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryUserStatus 
         Caption         =   "&User Status"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGoToWeb 
         Caption         =   "&Counterpoint Web Site"
         Begin VB.Menu mnuDocumentation 
            Caption         =   "&Documentation"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuWebHome 
            Caption         =   "&Home Page"
         End
      End
      Begin VB.Menu mnuHelpSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutHelp 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuAlert 
      Caption         =   ""
      Begin VB.Menu mnuAlertView 
         Caption         =   "View"
      End
   End
   Begin VB.Menu mnuBlank1 
      Caption         =   "  "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "css"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuCSSName 
         Caption         =   "Contract Spot Scheduler"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCSSStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ssb"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuSSBName 
         Caption         =   "Station Spot Builder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSSBStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "aeq"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuAEQName 
         Caption         =   "Affiliate Export Queue"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAEQStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "asi"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuASIName 
         Caption         =   "Affidavit Spot Import"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuASIStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "amb"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuAMBName 
         Caption         =   "Affiliate Measurement Builder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAMBStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "arq"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuARQName 
         Caption         =   "Affiliate Report Queue"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuARQStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "asg"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuASGName 
         Caption         =   "Avail Summary Generation"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuASGStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "sc"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuSCName 
         Caption         =   "Set Credit"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSCStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ce"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu mnuCEName 
         Caption         =   "Corporate Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "sfe"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu mnuSFEName 
         Caption         =   "Sales Force Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSFEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "me"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnuMEName 
         Caption         =   "Matrix Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "epe"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu mnuEPEName 
         Caption         =   "Efficio Projection Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEPEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ere"
      Index           =   12
      Visible         =   0   'False
      Begin VB.Menu mnuEREName 
         Caption         =   "Efficio Revenue Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEREStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "gpe"
      Index           =   13
      Visible         =   0   'False
      Begin VB.Menu mnuGPEName 
         Caption         =   "Get Paid Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGPEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "bd"
      Index           =   14
      Visible         =   0   'False
      Begin VB.Menu mnuBDName 
         Caption         =   "Backup Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBDStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "te"
      Index           =   15
      Visible         =   0   'False
      Begin VB.Menu mnuTeName 
         Caption         =   "Tableau Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTeStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "wvm"
      Enabled         =   0   'False
      Index           =   16
      Visible         =   0   'False
      Begin VB.Menu mnuWviName 
         Caption         =   "Web Vendor Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWviStatus 
         Caption         =   "See Affiliate-Alerts-Web Vendors"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "Pb"
      Index           =   17
      Visible         =   0   'False
      Begin VB.Menu mnuPbName 
         Caption         =   "Programmatic Buy"
      End
      Begin VB.Menu mnuPbStatus 
         Caption         =   "Not Working, Get Help!"
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "cai"
      Enabled         =   0   'False
      Index           =   18
      Visible         =   0   'False
      Begin VB.Menu mnucaiName 
         Caption         =   "Compel Auto Import"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucaiStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "re"
      Index           =   19
      Visible         =   0   'False
      Begin VB.Menu mnureName 
         Caption         =   "RAB Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnurestatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "cre"
      Index           =   20
      Visible         =   0   'False
      Begin VB.Menu mnucreName 
         Caption         =   "Custom Revenue Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucrestatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDate 
      Caption         =   "  "
   End
End
Attribute VB_Name = "Traffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of TRAFFIC.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  hmUlf                         tmPhf                         hmPhf                     *
'*  imPhfRecLen                                                                           *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mnuBackupAffiliate_Click      mnuBackupTraffic_Click        mnuExportMP3_Click        *
'*  mnuImportOrder_Click                                                                  *
'******************************************************************************************

' Copyright 1993 Counterpoint Software® All rights reserved.
' Proprietary Software, Do not copy
' This is a test
' File Name: Traffic.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Traffic input screen code
'Dan M 9/17/09 pbEmail object replaced
'Email sent
'Private objMyEmail As Email
''Email sender used to send out the emails
'Private WithEvents objEmailSender As EmailSender
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit
Option Compare Text

Private dmShellRet As Double


Const clMaxName As Long = 255
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const GW_HWNDFIRST = 0
Const WM_SETTEXT = &HC

Private Declare Function GetWindow Lib "user32" (ByVal HWnd _
    As Long, ByVal wCmd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Dim tmAuf As AUF
Dim imCountdown As Integer
Dim lmCountdownEnteredTime As Long
Dim imPrevAlertBlockChk As Integer

Dim hmAlert As Long
Dim lmMenuID As Long
Dim hmRedAlertBitmap As Long
Dim hmWhiteAlertBitmap As Long

Dim hmMonitor As Long
'Dim hmMonitorBitmap(0 To 13) As Long
'Dim imMonitor As Integer
Dim tmTmf As TMF
Dim imTmfRecLen As Integer
Dim tmTmfSrchKey1 As TMFKEY1
Dim smZone As String
Dim lmDate1970 As Long

Dim imShowVersionNo As Integer
Dim smWallpaper As String

Dim imTerminate As Integer  'True= terminating form; False= OK
Dim smCommand As String
Dim hmUrf As Integer
Dim tmUrfSrchKey As INTKEY0
Dim tmUrf As URF
Dim imUrfRecLen As Integer
Dim tmUlfSrchKey As LONGKEY0
Dim tmUlfSrchKey1 As ULFKEY1
Dim tmUlf As ULF
Dim imUlfRecLen As Integer
Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
'Dim hmSite As Integer
Dim imSiteRecLen As Integer
Dim hmSsf As Integer
Dim tmPif As PIF
Dim hmPif As Integer
Dim imPifRecLen As Integer
Dim tmVefSrchKey As INTKEY0
Dim hmVef As Integer
Dim imVefRecLen As Integer
Dim hmCef As Integer    'Comment file handle
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length
Dim tmChf As CHF
Dim hmCHF As Integer
Dim imCHFRecLen As Integer
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim tmRvf As RVF
Dim hmRvf As Integer
Dim imRvfRecLen As Integer
Dim tmPnfBuyer() As TYPESORT
Dim tmPnf As PNF
Dim hmPnf As Integer
Dim imPnfRecLen As Integer
Dim hmArf As Integer            'Name/Address file handle
Dim tmArf As ARF                'ARF record image
Dim tmArfSrchKey As INTKEY0            'ARF record image
Dim imArfRecLen As Integer        'ARF record length
Dim hmUaf As Integer
Dim tmUaf As UAF        'GSF record image
Dim tmUafSrchKey0 As LONGKEY0    'GSF key record image
Dim tmUafSrchKey2 As UAFKEY2    'GSF key record image
Dim imUafRecLen As Integer        'GSF record length
Dim tmGhf As GHF
Dim hmGhf As Integer
Dim imGhfRecLen As Integer
Dim tmGsf As GSF
Dim hmGsf As Integer
Dim imGsfRecLen As Integer
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim tmVff As VFF
Dim hmVff As Integer
Dim imVffRecLen As Integer
Dim tmVffSrchKey1 As INTKEY0    'GSF key record image
Dim hmPvf As Integer 'Package file handle
Dim tmPvf As PVF
Dim tmPvfSrchKey As LONGKEY0
Dim imPvfRecLen As Integer

Dim tmCifInfo() As CIFINFO
Dim tmCrfDates() As CRFDATES
Dim lmChgCount As Long
Dim lmTotalCount As Long

Dim tmCuf As CUF
Dim hmCuf As Integer
Dim imCufRecLen As Integer
Dim tmCrf As CRF
Dim hmCrf As Integer
Dim imCrfRecLen As Integer
Dim tmMcf As MCF
Dim hmMcf As Integer
Dim imMcfRecLen As Integer
Dim hmCif As Integer
Dim hmCnf As Integer

Dim imLastHourGGChecked As Integer
'8793
Private emMonitoringAllowed As MonitorAllowed
Private smMonitorID As String
Private Enum MonitorAllowed
    MonitorOff = 0
    MonitorNormal = 1
    MonitorNoEmail = 2
End Enum

Private rst_Ust As ADODB.Recordset



Private Function mCopyPictureImage(SourceImage As PictureBox) As Long
    Dim bm As BITMAP
    Dim newbm&
    Dim tdc&, oldbm&
    Dim di&

    ' First get the information about the image bitmap
    di = GetObjectAPI(SourceImage.Image, Len(bm), bm)
    bm.bmBits = 0
    ' Create a new bitmap with the same structure and size
    ' of the image bitmap
    newbm = CreateBitmapIndirect(bm)

    ' Create a temporary memory device context to use
    tdc = CreateCompatibleDC(SourceImage.hdc)
    ' Select in the newly created bitmap
    oldbm = SelectObject(tdc, newbm)

    ' Now copy the bitmap from the persistant bitmap in
    ' picture 2 (note that picture2 has AutoRedraw set TRUE
    di = BitBlt(tdc, 0, 0, bm.bmWidth, bm.bmHeight, SourceImage.hdc, 0, 0, SRCCOPY)
    ' Select out the bitmap and delete the memory DC
    oldbm = SelectObject(tdc, oldbm)
    di = DeleteDC(tdc)

    ' And return the new bitmap
    mCopyPictureImage = newbm
End Function

Private Sub mClearReportFiles()
    On Error GoTo ErrHand
    Dim ilRet As Integer
    Dim slTrafficINIPathName As String
    Dim smCSIServerINIFile As String
    Dim slStatus As String

    slTrafficINIPathName = sgDBPath & "Traffic.ini"
    smCSIServerINIFile = sgExePath & "\CSI_Server.ini"

    ilRet = csiClearReportFiles(sgDBPath, slTrafficINIPathName, 0)
    Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastClearReportStatus", slStatus)

    If ilRet = 0 Then
        MsgBox "Work files were reset successfully." & vbCrLf & vbCrLf & slStatus
    Else
        If ilRet = 8001 Then
            MsgBox "CSI_Server is not running. The work files cannot be reset."
            Exit Sub
        End If
        MsgBox "Unable to reset work files at this time." & vbCrLf & vbCrLf & slStatus
    End If
    Exit Sub

ErrHand:
    MsgBox "An error occurred in mClearReportFiles"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckFonts                     *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check Fonts                     *
'*                                                     *
'*******************************************************
Private Function mCheckFonts() As Integer
'    Dim slFontName As String
'    Dim flFontSize As Single
'    Dim llColor As Long
'    Dim ilRet As Integer
'    ilRet = 0
'    On Error GoTo mWrongFonts:
'    slFontName = pbcTestFont.FontName
'    flFontSize = pbcTestFont.FontSize
'    llColor = pbcTestFont.ForeColor
'    pbcTestFont.ForeColor = BLUE
'    pbcTestFont.FontBold = False
'    pbcTestFont.FontSize = 7
'    pbcTestFont.FontName = "Arial"
'    If ilRet <> 0 Then
'        pbcTestFont.FontSize = flFontSize
'        pbcTestFont.FontName = slFontName
'        pbcTestFont.FontSize = flFontSize
'        pbcTestFont.ForeColor = llColor
'        pbcTestFont.FontBold = True
'        Screen.MousePointer = vbDefault
'        MsgBox "'Arial' font missing from the System", vbOkOnly + vbCritical, "Error"
'        imTerminate = True
'        mCheckFonts = False
'        Exit Function
'    End If
'    pbcTestFont.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
'    pbcTestFont.FontSize = flFontSize
'    pbcTestFont.FontName = slFontName
'    pbcTestFont.FontSize = flFontSize
'    pbcTestFont.ForeColor = llColor
'    pbcTestFont.FontBold = True
'
'    pbcTestFont.FontName = "Monotype Sorts"
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "'Monotype Sorts' font missing from the System", vbOkOnly + vbCritical, "Error"
'        imTerminate = True
'        mCheckFonts = False
'        Exit Function
'    End If
'
'    pbcTestFont.FontName = "Arial Narrow"
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "'Arial Narrow' font missing from the System", vbOkOnly + vbCritical, "Error"
'        imTerminate = True
'        mCheckFonts = False
'        Exit Function
'    End If
'
'    pbcTestFont.FontName = "Times New Roman"
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "'Times New Roman' font missing from the System", vbOkOnly + vbCritical, "Error"
'        imTerminate = True
'        mCheckFonts = False
'        Exit Function
'    End If
'    ' Dan M 2/3/10 Century not used
''    pbcTestFont.FontName = "Century Schoolbook"
''    If ilRet <> 0 Then
''        Screen.MousePointer = vbDefault
''        MsgBox "'Century Schoolbook' font missing from the System", vbOkOnly + vbCritical, "Error"
''        imTerminate = True
''        mCheckFonts = False
''        Exit Function
''    End If
'    pbcTestFont.FontName = "Arial"
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "'Arial' font missing from the System", vbOkOnly + vbCritical, "Error"
'        imTerminate = True
'        mCheckFonts = False
'        Exit Function
'    End If
'    On Error GoTo 0
'    mCheckFonts = True
'    Exit Function
'mWrongFonts:
'    On Error GoTo 0
'    ilRet = 1
'    Resume Next
    Dim slFontName As String
    Dim flFontSize As Single
    Dim llColor As Long
    Dim ilRet As Integer
    Dim slMessage As String
    Dim blRet As Boolean
    
    ilRet = 0
    blRet = True
    On Error GoTo mWrongFonts:
    slMessage = ""
    slFontName = pbcTestFont.FontName
    flFontSize = pbcTestFont.FontSize
    llColor = pbcTestFont.ForeColor
    pbcTestFont.ForeColor = BLUE
    pbcTestFont.FontBold = False
    pbcTestFont.FontSize = 7
    pbcTestFont.FontName = "Arial"
    If ilRet <> 0 Then
        slMessage = "'Arial'"
        blRet = False
        ilRet = 0
    End If
    pbcTestFont.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    pbcTestFont.FontSize = flFontSize
    pbcTestFont.FontName = slFontName
    pbcTestFont.FontSize = flFontSize
    pbcTestFont.ForeColor = llColor
    pbcTestFont.FontBold = True
    pbcTestFont.FontName = "Monotype Sorts"
    If ilRet <> 0 Then
        slMessage = slMessage & " 'Monotype Sorts' "
        ilRet = 0
        blRet = False
    End If
    pbcTestFont.FontName = "Arial Narrow"
    If ilRet <> 0 Then
        slMessage = slMessage & " 'Arial Narrow' "
        ilRet = 0
        blRet = False
    End If

    pbcTestFont.FontName = "Times New Roman"
    If ilRet <> 0 Then
        slMessage = slMessage & " 'Times New Roman' "
        ilRet = 0
        blRet = False
    End If
    pbcTestFont.FontName = "Arial"
    If ilRet <> 0 Then
        slMessage = slMessage & " 'Arial' "
        ilRet = 0
        blRet = False
    End If
    ''7349
    'pbcTestFont.FontName = "Lucida Sans"
    'If ilRet <> 0 Then
    '    slMessage = slMessage & " 'Lucida Sans' "
    '    ilRet = 0
    '    blRet = False
    'End If
    On Error GoTo 0
    If blRet = False Then
        If Len(slMessage) > 0 Then
            slMessage = "Font(s) missing from the system: " & slMessage
            'gLogMsg slMessage, "TrafficErrors.Txt", False
            gMsgBox slMessage & ". Contact IT or Counterpoint to get help to install the missing font(s)", vbInformation + vbOKOnly, "Missing Fonts"
        End If
    End If
    mCheckFonts = blRet
    Exit Function
mWrongFonts:
    'On Error GoTo 0
    ilRet = 1
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckPW                        *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check User Password             *
'*                                                     *
'*******************************************************
Public Function mCheckPW() As Integer
    Dim ilSetPassword As Integer
    Dim llPWDate As Long
    Dim slNowDate As String
    Dim llNowDate As Long
    Dim slNowTime As String
    Dim ilRet As Integer
    Dim ilUrf As Integer
    Dim slStr As String
    Dim ilNewPWRequired As Integer

    'Check user password date- if over allowed time ask for new password
    ilSetPassword = False
    gUnpackDateLong tgUrf(0).iPasswordDate(0), tgUrf(0).iPasswordDate(1), llPWDate
    slNowDate = Format$(gNow(), "m/d/yy")
    llNowDate = gDateValue(slNowDate)
    slNowTime = Format$(gNow(), "h:m:s AM/PM")
    If tgUrf(0).iCode > 2 Then
        If igSportsSystem <> 3 Then
            ilNewPWRequired = False
            If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
                slStr = Trim$(tgUrf(0).sPassword)
                If Not gStrongPassword(slStr) Then
                    ilNewPWRequired = True
                End If
            End If
            If (llPWDate + tgSpf.iGNoDaysPass < llNowDate) Or (Trim$(tgUrf(0).sPassword) = "") Or (ilNewPWRequired) Then
                tgPWUrf = tgUrf(0)
                CSINewPW.Show vbModal
                If (igExitTraffic) Then
                    mCheckPW = False
                    'Unload Traffic
                    'Set Traffic = Nothing   'Remove data segment
                    Exit Function
                End If
                edcPassword.Text = sgPWResult
                ilSetPassword = True
            End If
        End If
    End If
    'Dan M allow csiNewPw as needed for guide
    If tgUrf(0).iCode = 2 And Not bgInternalGuide Then
        ilNewPWRequired = False
        If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
            slStr = Trim$(tgUrf(0).sPassword)
            If Not gStrongPassword(slStr) Then
                ilNewPWRequired = True
            End If
        End If
        If (Trim$(tgUrf(0).sPassword) = "") Then
            tgPWUrf = tgUrf(0)
            CSINewPW.lacCurrentPassword.Visible = False
            CSINewPW.edcCurrentPassword.Visible = False
            CSINewPW.Show vbModal
            If (igExitTraffic) Then
                mCheckPW = False
                Exit Function
            End If
            edcPassword.Text = sgPWResult
            ilSetPassword = True
        End If
    End If
    If tgUrf(0).iCode > 1 Then
        ilRet = csiHandleValue(0, 6)
        If ilRet = 1 Then
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                mCheckPW = True
                Exit Function
            End If
            mCheckPW = False
            Exit Function
        End If
        hmUrf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                MsgBox "Dual System Not Linked, Update Prohibited", vbOKOnly + vbCritical, "Traffic"
                mCheckPW = True
                Exit Function
            End If
            mCheckPW = False
            'Unload Traffic
            'Set Traffic = Nothing   'Remove data segment
            Exit Function
        End If
        imUrfRecLen = Len(tmUrf)  'btrRecordLength(hlUrf)  'Get and save record length
        'ilRet = btrGetFirst(hmUrf, tmUrf, imUrfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        'If ilRet <> BTRV_ERR_NONE Then
        '    If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
        '        MsgBox "Dual System Not Linked, Update Prohibited", vbOkOnly + vbCritical, "Traffic"
        '        mCheckPW = True
        '        Exit Function
        '    End If
        '    mCheckPW = False
        '    'Unload Traffic
        '    'Set Traffic = Nothing   'Remove data segment
        '    Exit Function
        'End If
        'Do While ilRet = BTRV_ERR_NONE
        '    gUrfDecrypt tmUrf
        '    If (Trim$(tmUrf.sName) = Trim$(sgUserName)) And (tmUrf.sDelete <> "Y") Then
        '        If ilSetPassword Then
        '            gPackDate slNowDate, tmUrf.iPasswordDate(0), tmUrf.iPasswordDate(1)
        '            tmUrf.sPassword = edcPassword.Text
        '        End If
        '        gPackDate slNowDate, tmUrf.iSignonDate(0), tmUrf.iSignonDate(1)
        '        gPackTime slNowTime, tmUrf.iSignonTime(0), tmUrf.iSignonTime(1)
        '        gUrfEncrypt tmUrf
        '        ilRet = btrUpdate(hmUrf, tmUrf, imUrfRecLen)
        '        If ilRet <> BTRV_ERR_NONE Then
        '            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
        '                MsgBox "Dual System Not Linked, Update Prohibited", vbOkOnly + vbCritical, "Traffic"
        '                mCheckPW = True
        '                Exit Function
        '            End If
        '            mCheckPW = False
        '            'Unload Traffic
        '            'Set Traffic = Nothing   'Remove data segment
        '            Exit Function
        '        End If
        '    End If
        '    ilRet = btrGetNext(hmUrf, tmUrf, imUrfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        'Loop
        For ilUrf = LBound(tgUrf) To UBound(tgUrf) Step 1
            Do
                tmUrfSrchKey.iCode = tgUrf(ilUrf).iCode
                ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmUrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                        MsgBox "Dual System Not Linked, Update Prohibited", vbOKOnly + vbCritical, "Traffic"
                        mCheckPW = True
                        Exit Function
                    End If
                   mCheckPW = False
                   Exit Function
                End If
                gUrfDecrypt tmUrf
                If ilSetPassword Then
                
'D.S. 09-09-15 place for API call of new Password  make call in EDS Subs

                    gPackDate slNowDate, tmUrf.iPasswordDate(0), tmUrf.iPasswordDate(1)
                    tmUrf.sOldPassword(2) = tmUrf.sOldPassword(1)
                    tmUrf.sOldPassword(1) = tmUrf.sOldPassword(0)
                    tmUrf.sOldPassword(0) = tmUrf.sPassword
                    tmUrf.sPassword = edcPassword.Text
                End If
                'gPackDate slNowDate, tmUrf.iSignOnDate(0), tmUrf.iSignOnDate(1)
                'gPackTime slNowTime, tmUrf.iSignOnTime(0), tmUrf.iSignOnTime(1)
                gUrfEncrypt tmUrf
                ilRet = btrUpdate(hmUrf, tmUrf, imUrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet = 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                    MsgBox "Dual System Not Linked, Update Prohibited, Error" & str$(ilRet), vbOKOnly + vbCritical, "Traffic"
                    mCheckPW = True
                    Exit Function
                Else
                    MsgBox "Update URF Prohibited, Error" & str$(ilRet), vbOKOnly + vbCritical, "Traffic"
                End If
                mCheckPW = False
                Exit Function
            End If
        Next ilUrf
        btrDestroy hmUrf
    End If
    mCheckPW = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckVersion                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check Database version          *
'*                                                     *
'*******************************************************
Private Function mCheckVersion() As Integer
    Dim ilRet As Integer
    Dim hlVersion As Integer
    Dim slVersion As String
    Dim ilPos As Integer
    Dim slChar As String
    Dim slLine As String

    slVersion = ""
    ilPos = InStr(sgCSVersion, " ") + 1
    slChar = Mid$(sgCSVersion, ilPos, 1)
    Do While slChar <> " "
        slVersion = slVersion & slChar
        ilPos = ilPos + 1
        If ilPos > Len(sgCSVersion) Then
            Exit Do
        End If
        slChar = Mid$(sgCSVersion, ilPos, 1)
    Loop
    ilRet = 0
    On Error GoTo mWrongVersion:
    'hlVersion = FreeFile
    'Open sgDBPath & "Version.Csi" For Input Access Read As hlVersion
    ilRet = gFileOpen(sgDBPath & "Version.Csi", "Input Access Read", hlVersion)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Version.Csi missing from " & sgDBPath, vbOKOnly + vbCritical, "Error"
        imTerminate = True
        'Unload Traffic
        'Set Traffic = Nothing   'Remove data segment
        mCheckVersion = False
        Exit Function
    End If
    On Error GoTo mWrongVersion:
    Line Input #hlVersion, slLine
    On Error GoTo 0
    Close hlVersion
    If (Asc(slLine) = 26) Or (ilRet <> 0) Or (StrComp(slVersion, slLine, 1) <> 0) Then    'Ctrl Z
        Screen.MousePointer = vbDefault
        MsgBox "Programs and Database versions don't match", vbOKOnly + vbCritical, "Error"
        imTerminate = True
        'Unload Traffic
        'Set Traffic = Nothing   'Remove data segment
        mCheckVersion = False
        Exit Function
    End If
    mCheckVersion = True
    Exit Function
mWrongVersion:
    On Error GoTo 0
    ilRet = 1
    Resume Next
End Function


Private Sub cmcAffiliate_Click()
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim blStart As Boolean
    Dim ilRet As Integer
    
    Dim hlWndAffiliate As Long
    Dim hlWndChild As Long
    Dim slTestName As String * 255
    Dim llNumChars As Long
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserNameToPassToAffiliate & "\" & sgUserPasswordToPassToAffiliate & "\" & "FROMTRAFFIC"
    Else
        slCommandStr = "Traffic^Prod\" & sgUserNameToPassToAffiliate & "\" & sgUserPasswordToPassToAffiliate & "\" & "FROMTRAFFIC"
    End If
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    On Error GoTo LoadErr
    ilRet = 0
    'AppActivate "CSI Affiliate"
    'If ilRet = 1 Then
    '    dmShellRet = Shell(sgExePath & "Affiliat2.Exe " & slCommandStr, vbNormalFocus)
    'End If
    
    hlWndAffiliate = FindWindow(vbNullString, "AffiliateTrafficCom")
    If hlWndAffiliate = 0 Or ilRet = 1 Then
        ilRet = 0
        AppActivate "CSI Affiliate"
        If ilRet = 1 Then
            dmShellRet = Shell(sgExePath & "Affiliat2.Exe " & slCommandStr, vbNormalFocus)
        End If
    Else
        hlWndChild = GetWindow(hlWndAffiliate, GW_CHILD)
        Do While hlWndChild <> 0
            llNumChars = GetClassName(hlWndChild, slTestName, (clMaxName))
            If InStr(1, slTestName, "TextBox") > 0 Then
                SendMessageByString& hlWndChild, WM_SETTEXT, 0, "Set Window size"
                ilRet = 0
                AppActivate "CSI Affiliate"
                If ilRet = 1 Then
                    dmShellRet = Shell(sgExePath & "Affiliat2.Exe " & slCommandStr, vbNormalFocus)
                End If
                Exit Do
            End If
            hlWndChild = GetWindow(hlWndChild, GW_HWNDNEXT)
        Loop
    End If
    
    'End If
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next

End Sub

Private Sub cmcEMail_Click()
    
    'D.S. 07-13-15
    Dim slBrowserURL As String
    
    If gLoadOption("EDS-API", "EDS_RootWebURL", slBrowserURL) Then
    Else
        MsgBox "[EDS-API] - EDS_RootWebURL definition not found in the Traffic.ini file"
        Exit Sub
    End If
    'Open the default browser
    ShellExecute 0, vbNullString, slBrowserURL, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub cmcFind_Click()
    CntrFind.Show vbModeless    'vbModeless  'This is required to activate form if previousely activated
    CntrFind.SetFocus
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String
    Dim ilRptType As Integer
    Dim ilPos As Integer

    ilRptType = 0
    ''If igGGFlag = 0 Then
    ''    Exit Sub
    ''End If
    'ilPos = InStr(1, sgCommand, "/Reports", vbTextCompare)
    'If (igGGFlag = 0) And (ilPos <= 0) Then
    '    Exit Sub
    'End If
    If (igGGFlag = 0) And (igRptGGFlag = 0) Then
        Exit Sub
    End If
    igRptCallType = -2  'CONTRACTSJOB
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
    End If
    sgCommandStr = slStr
    sgSpfStamp = "~"    'Force read
    gSpfRead
    'gSetUsingTraffic
    On Error Resume Next
    RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated

End Sub

Private Sub cmcTask_Click(Index As Integer)
    Dim slStr As String
    Dim ilRptType As Integer
    Dim ilAskMkt As Integer
    Dim ilLoop As Integer

    ilRptType = 0
    If igGGFlag = 0 Then
        Exit Sub
    End If
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    If Not igVehicleDefined Then
        MsgBox "Please Define Vehicle", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    '*** Market Code
    If (tgSpf.sMktBase = "Y") And ((Index = 2) Or (Index = 3) Or (Index = 5) Or (Index = 7)) Then
        ilAskMkt = True
        If Index = 3 Then
            If igJobShowing(CONTRACTSJOB) Then
                ilAskMkt = False
            End If
        End If
        If Index = 2 Then
            If igJobShowing(PROPOSALSJOB) Then
                ilAskMkt = False
            End If
        End If
        If Index = 5 Then
            If igJobShowing(SPOTSJOB) Then
                ilAskMkt = False
            End If
        End If
        If Index = 7 Then
            If igJobShowing(LOGSJOB) Then
                ilAskMkt = False
            End If
        End If
        If ilAskMkt Then
            LogMkt.Show vbModal
            DoEvents
            If UBound(igMktCode) <= 0 Then
                'plcBasic10.Visible = False
                'plcBasic10.Visible = True
                'imTaskIndex = -1
                Exit Sub
            End If
            If (Index = 2) Or (Index = 3) Then
                ReDim igCntrMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igCntrMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            ElseIf (Index = 5) Then
                ReDim igSpotMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igSpotMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            Else
                ReDim igLogMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igLogMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            End If
        End If
    Else
        'ReDim igMktCode(0 To 0) As Integer
        'ReDim igCntrMktCode(0 To 0) As Integer
        'ReDim igLogMktCode(0 To 0) As Integer
    End If
    '*** Market Code
    Select Case Index
        Case 0  'Budget
            If tgSpf.sGUsePropSys = "Y" Then
                If igWinStatus(BUDGETSJOB) = 0 Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            On Error Resume Next
            If Not gRecExistForFile("Sof.btr") Then
                MsgBox "Please Define Sales Offices", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
            If Not gRecExistForFile("Slf.btr") Then
                MsgBox "Please Define Salespeople", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
            'If Not igJobShowing(BUDGETSJOB) Then
            '    Budget.Show vbModeless  'This is required to activate form if previousely activated
            'Else
            '    Budget.SetFocus
            'End If
            gShellToBugdet Traffic
            
        Case 1  'Rate Card
            If igWinStatus(RATECARDSJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            'If Not igJobShowing(RATECARDSJOB) Then
            '    RateCard.Show vbModeless  'This is required to activate form if previousely activated
            'Else
            '    RateCard.SetFocus
            'End If
            gShellToRateCard
            
        Case 2  'Proposal
            If tgSpf.sGUsePropSys = "Y" Then
                If igWinStatus(PROPOSALSJOB) = 0 Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            On Error Resume Next
            If Not gRecExistForFile("Rcf.btr") Then
                MsgBox "Please Define Rate Card", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
            If Not igJobShowing(CONTRACTSJOB) Then
                igAlertCntrStatus = -1
                sgCntrScreen = "Proposals"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Contract.SetFocus
            End If
            
        Case 3  'Orders
            If igWinStatus(CONTRACTSJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If Not gRecExistForFile("Rcf.btr") Then
                MsgBox "Please Define Rate Card", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
            If Not igJobShowing(CONTRACTSJOB) Then
                igAlertCntrStatus = -1
                sgCntrScreen = "Orders"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Contract.SetFocus
            End If
            
        Case 4  'Programming
            If igWinStatus(PROGRAMMINGJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If Not igJobShowing(PROGRAMMINGJOB) Then
                Program.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Program.SetFocus
            End If
            
        Case 5  'Spots
            If igWinStatus(SPOTSJOB) = 0 Then
                Exit Sub
            End If
            If tgSpf.sSystemType = "R" Then
                sgSpotCallType = "T"
            Else
                sgSpotCallType = "A"
            End If
            On Error Resume Next
            If Not igJobShowing(SPOTSJOB) Then
                Spots.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Spots.SetFocus
            End If
            
        Case 6  'Copy
            If igWinStatus(COPYJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If tgSpf.sUseCartNo <> "N" Then
                If Not gRecExistForFile("Mcf.btr") Then
                    MsgBox "Please Define Media", vbOKOnly + vbInformation, "Counterpoint"
                    Exit Sub
                End If
            End If
            If Not igJobShowing(COPYJOB) Then
                Copy.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Copy.SetFocus
            End If
            
        Case 7  'Log
            If igWinStatus(LOGSJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If Not igJobShowing(LOGSJOB) Then
                Logs.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Logs.SetFocus
            End If
            
'        Case 8  'Station Feed
'            If tgSpf.sGUseAffFeed = "Y" Then
'                If igWinStatus(STATIONFEEDJOB) = 0 Then
'                    Exit Sub
'                End If
'            Else
'                Exit Sub
'            End If
'            On Error Resume Next
'            If Not igJobShowing(STATIONFEEDJOB) Then
'                ExpStnFd.Show vbModeless  'This is required to activate form if previousely activated
'            Else
'                ExpStnFd.SetFocus
'            End If

        Case 9  'Post Log
            If igWinStatus(POSTLOGSJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            'If Not igJobShowing(POSTLOGSJOB) Then
            '    PostLog.Show vbModeless  'This is required to activate form if previousely activated
            'Else
            '    PostLog.SetFocus
            'End If
            If igJobShowing(POSTLOGSJOB) Then
                If igPostType = 0 Then
                    PostLog.SetFocus
                ElseIf igPostType = 1 Then
                    PostRep.SetFocus
                ElseIf igPostType = 2 Then
                    PostRep.SetFocus
                ElseIf igPostType = 3 Then
                    PostRep.SetFocus
                ElseIf igPostType = 4 Then
                    PostRep.SetFocus
                End If
            Else
                gAnyClustersDef
                gAnyRepDef
                'Determine which Post Log screen
                If (sgMktClusterDef) = "Y" Or (sgRepDef = "Y") Then
                    PLogMkt.Show vbModal
                    DoEvents
                    If igPostType < 0 Then
                        'plcBasic10.Visible = False
                        'plcBasic10.Visible = True
                        'imTaskIndex = -1
                        Exit Sub
                    End If
                Else
                    igPostType = 0
                End If
                If igPostType = 0 Then
                    PostLog.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 1 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 2 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 3 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 4 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 5 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                End If
            End If
            
        Case 10 'Invoices
            If igWinStatus(INVOICESJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            'If Not igJobShowing(INVOICESJOB) Then
            '    Invoice.Show vbModeless  'This is required to activate form if previousely activated
            'Else
            '    Invoice.SetFocus
            'End If
            gShellToInvoice 'Traffic
            
        Case 11 'Collections
            If igWinStatus(COLLECTIONSJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If Not igJobShowing(COLLECTIONSJOB) Then
                Collect.Show vbModeless  'This is required to activate form if previousely activated
            Else
                Collect.SetFocus
            End If
            
        Case 12 'Commission
            If igWinStatus(SLSPCOMMSJOB) = 0 Then
                Exit Sub
            End If
            igRptCallType = SLSPCOMMSJOB    'Using contractjob instead spotsjob
            igRptType = 0
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType)) & "\" & Trim$(str$(tgUrf(0).iCode))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType)) & "\" & Trim$(str$(tgUrf(0).iCode))
            End If
            sgCommandStr = slStr
            On Error Resume Next
            If tgSpf.sSubCompany = "Y" Then
                If Not igJobShowing(SLSPCOMMSJOB) Then
                    SlspComm.Show vbModeless  'This is required to activate form if previousely activated
                Else
                    SlspComm.SetFocus
                End If
            Else
                RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated
            End If
            
        Case 13 'Reports
            igRptCallType = -2  'CONTRACTSJOB
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
            End If
            sgCommandStr = slStr
            sgSpfStamp = "~"    'Force read
            gSpfRead
            'gSetUsingTraffic
            On Error Resume Next
            RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated
    End Select
End Sub

Private Sub edcMsgFromAffiliate_Change()
    MsgBox "Window State " & Me.WindowState
End Sub

Private Sub MDIForm_Activate()
    'Code added to test if could cycle to active module- i.e. traffic->order->advertiser, then desktop, then select traffic
    'If Traffic.Enabled = False Then
    '    SendKeys "%{TAB}"
    'Else
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
    'End If
End Sub
Private Sub MDIForm_LinkExecute(CmdStr As String, Cancel As Integer)
    Dim slCommand As String
    slCommand = CmdStr
    'If slCommand = "BT" Then
    '    'imUpdateAllowed = True
    '    gShowBranner True
    'ElseIf slCommand = "BF" Then
    '    'imUpdateAllowed = False
    '    gShowBranner False
    'ElseIf Asc(slCommand) = Asc("@") Then
    '    'Remove @ symbol
    '    sgDoneMsg = right$(slCommand, Len(slCommand) - 1)
    '    igChildDone = True
    'Else
        'plcHelp.Caption = " " & slCommand
        'plcHelp.Cls
        'plcHelp.CurrentX = 0
        'plcHelp.CurrentY = 0
        'plcHelp.Print " " & slCommand
    'End If
'    igWaitForDDE = False
    Cancel = 0
End Sub
Private Sub MDIForm_Load()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slCommand As String
    Dim slDate As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilDelay As Integer
    Dim slSetDate As String
    Dim ilShell As Integer
    Dim ilErr As Integer
    Dim slMessage As String
    Dim slStartIn As String
    Dim slWinDir As String
    Dim llWinLen As Long
    Dim slWindowsDir As String * MAX_PATH
    Dim dlShellRet As Double
    
    igTrafficFirstActivate = True 'TTP 10937 - Contract Dashboard: new screen to show contract header information
    
    'Debug
    iDebugIdx = 0
    dmShellRet = -1
    imLastHourGGChecked = Hour(Now)

    imTerminate = False
    igIsSignedOn = False 'D.S. 3/10/03
    igRUStartExeCopy = False
    imPrevAlertBlockChk = 0
    bgAllowCopyGrid = False
    'Only allow one instance
    If App.PrevInstance Then
        MsgBox "Only one copy of Traffic can be run at a time, sorry", vbOKOnly + vbInformation, "Counterpoint"
        End
    End If
    sgCommand = Command$
    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    imShowVersionNo = igShowVersionNo
    gSetBkgdMode
    'Align 1=Top; 2=Bottom
    If pbcMsgArea.Align = 2 Then
        igTaskLoc = 1
    Else
        igTaskLoc = 0
    End If
    Traffic.WindowState = vbMaximized
    Traffic.Show 0

    'Get current setting
    lgCurrHRes = GetDeviceCaps(pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(pbcList.hdc, BITSPIXEL)

    'Determine if on Citrix
    llWinLen = GetWindowsDirectory(slWindowsDir, MAX_PATH)
    slWinDir = Left$(slWindowsDir, llWinLen)
    If (InStr(1, slWinDir, "c:\win", vbTextCompare) <> 0) Then
        igOnCitrix = False
    Else
        igOnCitrix = True
    End If
    slSetDate = ""
    csiInitName
    'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
    'ilRet = csiSetName("SYSDate", "")
    ilRet = gCsiSetName("")
    pbcMsgArea.Move pbcMsgArea.Left, pbcMsgArea.Top, Traffic.Width
    'cmcAuxWnd.Move pbcMsgArea.Left + pbcMsgArea.Width - cmcAuxWnd.Width
    'plcTime.Move cmcAuxWnd.Left - plcTime.Width
    'plcDate.Move plcTime.Left - plcDate.Width
    'plcHelp.Move plcHelp.Left, plcHelp.Top, plcDate.Left - cmcBasicWnd.Left - cmcBasicWnd.Width + 15
    'plcDate.Move pbcMsgArea.Left + pbcMsgArea.Width - plcDate.Width
    'cmcTask(13).Move cmcTask(13).Left, cmcTask(13).Top, plcDate.Left - cmcTask(12).Left - cmcTask(12).Width + 15
    'cmcReport.Move pbcMsgArea.Left + pbcMsgArea.Width - cmcReport.Width
    cmcReport.Move cmcTask(12).Left + cmcTask(12).Width + 15, cmcReport.Top, pbcMsgArea.Width - cmcTask(12).Left - cmcTask(12).Width + 15
    sgSystemDate = gAdjYear(Format$(gNow(), "m/d/yy"))  'fix date
    igResetSystemDate = False
    ilRet = csiSetAlloc("NAMES", 0, 2)
    If StrComp(sgCommand, "Debug1", 1) <> 0 Then
        ilPos = InStr(1, sgCommand, "/X:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, sgCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(sgCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    sgCommand = Trim$(Left$(sgCommand, ilPos - 1))
                Else
                    sgCommand = ""
                End If
            Else
                slStr = Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    sgCommand = Left$(sgCommand, ilPos - 1) & Mid$(sgCommand, ilSpace)
                Else
                    sgCommand = Mid$(sgCommand, ilSpace)
                End If
            End If
            On Error GoTo TrafficFormErr
            'ilShell = Shell(slStr)
            dlShellRet = Shell(slStr)
            If ilErr Then
                Exit Sub
            End If
            While GetModuleUsage(ilShell) > 0
                ilRet = DoEvents()
            Wend
            On Error GoTo 0
        End If
        igRemoteUser = False
        ilPos = InStr(1, sgCommand, "/RemoteUser", 1)
        If ilPos > 0 Then
            igRemoteUser = True
        End If
        igBypassFileSize = False
        ilPos = InStr(1, sgCommand, "/NoSize", 1)
        If ilPos > 0 Then
            igBypassFileSize = True
        End If
        igDemo = False
        ilPos = InStr(1, sgCommand, "/Demo", 1)
        If ilPos > 0 Then
            igDemo = True
        End If
        igChkExe = True
        ilPos = InStr(1, sgCommand, "/NoDateCheck", 1)
        If ilPos > 0 Then
            igChkExe = False
        End If
        bgAllowCopyGrid = False
        ilPos = InStr(1, sgCommand, "/Grid", 1)
        If ilPos > 0 Then
            bgAllowCopyGrid = True
        End If
        ilPos = InStr(1, sgCommand, "/D:", 1)
        If ilPos > 0 Then
            'ilPos = InStr(slCommand, ":")
            ilSpace = InStr(ilPos, sgCommand, " ")
            If ilSpace = 0 Then
                slDate = Trim$(Mid$(sgCommand, ilPos + 3))
            Else
                slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
            End If
            If gValidDate(slDate) Then
                slDate = gAdjYear(slDate)
                'Date$ = slDate
                'igResetSystemDate = True
                slSetDate = slDate
                'Moved to after mTrafficStartUp
                'ilRet = csiSetName("SYSDate", slDate)
                'tmcClock_Timer
            End If
        End If
        gInitGlobalVar   'Initialize global variables
        'D.S. 5/31/13 added optional parameter to gCheckDDFDates below
        If Not gCheckDDFDates(False) Then
            'ttp 4905- pervasive not open.  pause and try again.
            Sleep 2000
            If Not gCheckDDFDates() Then
                Unload Traffic
                Set Traffic = Nothing   'Remove data segment
                Exit Sub
            End If
        End If
        smWallpaper = sgWallpaper
        'Test if exe need updating
        If igChkExe Then
            If Len(sgDLExePath) > 0 Then
                lbcFile.Path = Left$(sgExePath, Len(sgExePath) - 1)
                'lbcFile.Pattern = "*.Exe"
                lbcFile.Pattern = "*.*"
                For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                    On Error GoTo ExeTestErr
                    ilRet = True
                    If FileDateTime(sgExePath & lbcFile.List(ilLoop)) < FileDateTime(sgDLExePath & lbcFile.List(ilLoop)) Then
                        If ilRet Then
                            ilRet = MsgBox("Your Software must be Updated to the Latest Version before You may use the System, Proceed with Update", vbYesNo + vbExclamation, "Traffic")
                            If ilRet = vbYes Then
                                On Error GoTo ExeTestErr
                                ilRet = True
                                If FileDateTime(sgExePath & "ExeCopy.Exe") < FileDateTime(sgDLExePath & "ExeCopy.Exe") Then
                                    If ilRet Then
                                        FileCopy sgDLExePath & "ExeCopy.Exe", sgExePath & "ExeCopy.Exe"
                                        For ilDelay = 0 To 20 Step 1
                                            DoEvents
                                        Next ilDelay
                                    End If
                                End If
                                'ilShell = Shell(sgExePath & "ExeCopy.Exe /CS", 1)
                                dlShellRet = Shell(sgExePath & "ExeCopy.Exe /CS", 1)
                            End If
                            Unload Traffic
                            Set Traffic = Nothing   'Remove data segment
                            Exit Sub
                        End If
                    End If
                Next ilLoop
                ''lbcFile.Path = Left$(sgDLExePath, Len(sgDLExePath) - 1)
                'lbcFile.Path = Left$(sgExePath, Len(sgExePath) - 1)
                'lbcFile.Pattern = "*.Dll"
                'For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                '    On Error GoTo ExeTestErr
                '    ilRet = True
                '    If FileDateTime(sgExePath & lbcFile.List(ilLoop)) < FileDateTime(sgDLExePath & lbcFile.List(ilLoop)) Then
                '        If ilRet Then
                '            ilRet = MsgBox("Your Software must be Updated to the Latest Version before You may use the System, Proceed with Update", vbYesNo + vbExclamation, "Traffic")
                '            If ilRet = vbYes Then
                '                ilShell = Shell(sgExePath & "ExeCopy.Exe /CS", 1)
                '            End If
                '            Unload Traffic
                '            Set Traffic = Nothing   'Remove data segment
                '            Exit Sub
                '        End If
                '    End If
                'Next ilLoop
            End If
            If Len(sgDLRptPath) > 0 Then
                lbcFile.Path = Left$(sgRptPath, Len(sgRptPath) - 1)
                'For ilPass = 0 To 1 Step 1
                '    If ilPass = 0 Then
                        'lbcFile.Pattern = "*.Rpt"
                        lbcFile.Pattern = "*.*"
                '    Else
                '        lbcFile.Pattern = "*.Lst"
                '    End If
                    For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                        On Error GoTo ExeTestErr
                        ilRet = True
                        If FileDateTime(sgRptPath & lbcFile.List(ilLoop)) < FileDateTime(sgDLRptPath & lbcFile.List(ilLoop)) Then
                            If ilRet Then
                                ilRet = MsgBox("Your Software must be Updated to the Latest Version before You may use the System, Proceed with Update", vbYesNo + vbExclamation, "Traffic")
                                If ilRet = vbYes Then
                                    'ilShell = Shell(sgExePath & "ExeCopy.Exe /CS", 1)
                                    dlShellRet = Shell(sgExePath & "ExeCopy.Exe /CS", 1)
                                End If
                                Unload Traffic
                                Set Traffic = Nothing   'Remove data segment
                                Exit Sub
                            End If
                        End If
                    Next ilLoop
                'Next ilPass
            End If
        End If
        gSpfRead
        If Trim$(slSetDate) = "" Then
            If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
                slSetDate = "12/14/2015"
                slDate = slSetDate
            End If
        End If
        If Trim$(slSetDate) <> "" Then
            'TrafficStartUp make csiSetAlloc cal that removes all values
            'Dan M problems with reports.exe and csiSetName
            'ilRet = csiSetName("SYSDate", slDate)
            ilRet = gCsiSetName(slDate)
        End If
        Signon.Show vbModal
        If (igExitTraffic) Then
            Unload Traffic
            Set Traffic = Nothing   'Remove data segment
            Exit Sub
        End If
        'Move it prior to Sign-on so that tgSpf.sUsingTraffic is set
        'gSpfRead
        If Not mCheckPW() Then
            Unload Traffic
            Set Traffic = Nothing   'Remove data segment
            Exit Sub
        End If


        hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hgAuf, "", sgDBPath & "AUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open Alert File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        hgUlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hgUlf, "", sgDBPath & "ULF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open User Log File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        hgSite = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hgSite, "", sgDBPath & "Site.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open Affiliate Site File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        hgTmf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hgTmf, "", sgDBPath & "Tmf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open Task Monitor File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        imTmfRecLen = Len(tmTmf)
        
        ilRet = gAlertCheckBlock(tmAuf)
        If ilRet = 1 Then   'Block user
            'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Or ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
            If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
                slMessage = "System block mode has been overridden"
                MsgBox slMessage, vbCritical, "Warning"
            Else
                slMessage = "System is shut down until further notice"
                MsgBox slMessage, vbCritical, "Shut down"
                Unload Traffic
                Set Traffic = Nothing   'Remove data segment
                Exit Sub
            End If
        ElseIf ilRet = 2 Then   'Block Initiated by this user
        End If
        ilRet = mSignOnOff(True)
        'If igRemoteUser Then
        '    slParameter = "\RemoteUser"
        'Else
        '    slParameter = "\"
        'End If
        'If igBypaFileListBoxSize Then
        '    slParameter = slParameter & "\NoSize"
        'Else
        '    slParameter = slParameter & "\"
        'End If
        'If igDemo Then
        '    slParameter = slParameter & "\Demo\"
        'Else
        '    slParameter = slParameter & "\\"
        'End If
        'If igTestSystem Then
        '    sgCommand = sgUserName & "\Test" & slParameter & sgCommand
        'Else
        '    'Retain sgCommand as it was so it can be used when user reenters signon
        '    sgCommand = sgUserName & "\Prod" & slParameter & sgCommand
        'End If
    End If
    slCommand = sgCommand
    Screen.MousePointer = vbHourglass
    'slCommand = Command$
    smCommand = ""
    If StrComp(slCommand, "Debug1", 1) = 0 Then
        sgUserName = "Guide" '"Counterpoint"
'        igTestSystem = False    'True
        igRemoteUser = False    'True
        igBypassFileSize = False
        igDemo = False
        'mnuSignon.Enabled = False
        gInitGlobalVar   'Initialize global variables
        smWallpaper = sgWallpaper
        If Not gCheckDDFDates() Then
            Unload Traffic
            Set Traffic = Nothing   'Remove data segment
            Exit Sub
        End If
    Else
        'UserName\Prod or Test\RemoteUser or blank\NoSize or blank\Demo or Blank\Command Line when Calling CSISignOn
        'ilParse = gParseItem(slCommand, 1, "\", sgUserName)
        'If ilParse <> CP_MSG_NONE Then
        '    Screen.MousePointer = vbDefault
        '    MsgBox "Start Traffic properly", vbOkOnly + vbInformation, "Counterpoint"
        '    End
        'End If
        'ilParse = gParseItem(slCommand, 2, "\", slStr)
        'If ilParse <> CP_MSG_NONE Then
        '    Screen.MousePointer = vbDefault
        '    MsgBox "Start Traffic properly", vbOkOnly + vbInformation, "Counterpoint"
        '    End
        'End If
        'If StrComp(slStr, "Prod", 1) = 0 Then
        '    igTestSystem = False
        'ElseIf StrComp(slStr, "Test", 1) = 0 Then
        '    igTestSystem = True
        'Else
        '    Screen.MousePointer = vbDefault
        '    MsgBox "Traffic not started probably, sorry", vbOkOnly + vbInformation, "Counterpoint"
        '    End
        'End If
        'ilParse = gParseItem(slCommand, 3, "\", slStr)
        'If StrComp(slStr, "RemoteUser", 1) = 0 Then
        '    igRemoteUser = True
        'Else
        '    igRemoteUser = False
        'End If
        'ilParse = gParseItem(slCommand, 4, "\", slStr)
        'If StrComp(slStr, "NoSize", 1) = 0 Then
        '    igBypaFileListBoxSize = True
        'Else
        '    igBypaFileListBoxSize = False
        'End If
        'ilParse = gParseItem(slCommand, 5, "\", slStr)
        'If StrComp(slStr, "Demo", 1) = 0 Then
        '    igDemo = True
        'Else
        '    igDemo = False
        'End If
        'ilParse = gParseItem(slCommand, 6, "\", slStr)
        'slCommand = slStr
        'smCommand = slCommand
    End If
    mkcPhone.Text = "(___) ___-____ Ext(____)"  'Required for verios 4 of V.B.
    'gInitGlobalVar   'Initialize global variables
    'Test if Database and Exe are on same version
    If Not mCheckVersion() Then
        Unload Traffic
        Set Traffic = Nothing   'Remove data segment
        Exit Sub
    End If
    'Dan M 3/30/2011 don't stop because fonts don't check out.  Also, each 'false' in mCheckFonts returns a message, so don't need to send one here.
    mCheckFonts
'    If Not mCheckFonts() Then
'        Unload Traffic
'        Set Traffic = Nothing   'Remove data segment
'        Exit Sub
'    End If
    If Not mTrafficStartUp() Then
        Unload Traffic
        Set Traffic = Nothing
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    'Traffic.WindowState = vbMaximized
    'Traffic.Show 0
    tmcClock_Timer
    'App.HelpFile = sgHelpPath & "traffic.hlp"
    On Error Resume Next
    'pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.Bmp")
    'On Error GoTo mRetryPicture:
    'ilRet = 0
    If gFileExist(sgLogoPath & "CustLogo.jpg") = FILEEXISTS Then
        pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.jpg")
    ElseIf gFileExist(sgLogoPath & "CustLogo.gif") = FILEEXISTS Then
        pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.gif")
    ElseIf gFileExist(sgLogoPath & "CustLogo.Bmp") = FILEEXISTS Then
        pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.Bmp")
    End If
    'pbcSymb.Picture = LoadPicture(sgLogoPath & "CustSymb.Bmp")
    'On Error GoTo mRetryPicture:
    'ilRet = 0
    If gFileExist(sgLogoPath & "CustSymb.jpg") = FILEEXISTS Then
        pbcSymb.Picture = LoadPicture(sgLogoPath & "CustSymb.jpg")
    ElseIf gFileExist(sgLogoPath & "CustSymb.gif") = FILEEXISTS Then
        pbcSymb.Picture = LoadPicture(sgLogoPath & "CustSymb.gif")
    ElseIf gFileExist(sgLogoPath & "CustSymb.Bmp") = FILEEXISTS Then
        pbcSymb.Picture = LoadPicture(sgLogoPath & "CustSymb.Bmp")
    End If
    On Error GoTo 0
    'Force csi_math.dll load to avoid timing problem with set credit
    slStr = gAddStr("1", "1")
'    Master.Picture = LoadPicture("C:\Afterdrk\bitmaps\earth.BMP")
    'MousePointer = vbHourGlass  'Wait
'    InitAvailEvnts
'    InitCnt
'    SetCompetitive
'    InitSpotEvnts
    'Start Alerts
'    If tgSpf.iGAlertInterval <> 0 Then
'        If (tgUrf(0).sCompAlert <> "N") Or (tgUrf(0).sReprintLogAlert <> "N") Then
'            If tgSpf.iGAlertInterval > 0 Then
'                igAlertTimer = tgSpf.iGAlertInterval
'            End If
'            tmcCheckAlert_Timer
'            igAlertFlash = -1
'        End If
'    End If
    If (igSportsSystem <> 2) And (igSportsSystem <> 3) Then
        ilRet = gAlertForceCheck()
    End If
    Screen.MousePointer = vbDefault  'Wait
    'Signon.Show vbModal
    'If igExitTraffic Then
    '    Unload Traffic
    '    Set Traffic = Nothing   'Remove data segment
    '    Exit Sub
    'End If
    'Dan M 8/30/10 test for updates in case user is skipping 'traffic.exe'
    If mUpdatesAvailable() Then
        'gMsgBox "An update is available.  Please run Traffic.exe to install updates", 0, "Updates Available"
        If Trim$(tgUrf(0).sRept) <> "" Then
            '5676
             gLogMsg "An update is available for " & Trim$(tgUrf(0).sRept) & ".  Please run Traffic.exe to install updates", "UpdateErrors.Txt", False
            'gMsgBox "An update is available for " & Trim$(tgUrf(0).sRept) & ".  Please run Traffic.exe to install updates", 0, "Updates Available"
        Else
             gLogMsg "An update is available for " & Trim$(sgUserName) & ".  Please run Traffic.exe to install updates", "UpdateErrors.Txt", False
            'gMsgBox "An update is available for " & Trim$(sgUserName) & ".  Please run Traffic.exe to install updates", 0, "Updates Available"
        End If
    End If
    'Dan M 8/31/10 if netReporter doesn't exist, let user know
'    If Not gCallNetReporter(CsiReportCall.StartReports) Then
'        'if running in test folder, csiNetReporterAlternate is missing.
'        gMsgBox "A component is missing that is needed for reports to run properly.  Please contact Counterpoint.", 0, "CsiNetReporter Missing"
'    End If
    'Dan M 3/23/10 start report module
    'gCallNetReporter CsiReportCall.StartReports
    sgFnfFileDate = FileDateTime(sgDBPath & "FNF.btr")
    ilRet = gBuildFeedArray()
    ilRet = gBuildWebUser
    '7496
    mnuExportMP2.Caption = "Audio " & UCase(Mid(sgAudioExtension, 2))
    '11010
    mnuImportMegaphone.Enabled = gIsVendor("Megaphone")
    '7967 test if allow web vendor
    '10473 issue with sockets and making web sql calls.
    bgVendorToWebAllowed = False
'    Me.MousePointer = vbHourglass
'    gVendorToWebAllowed tmcWebConnectIssue
'    Me.MousePointer = vbDefault
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture1:
    pbcLogo.Picture = LoadPicture()
    Resume Next
mNoPicture2:
    pbcSymb.Picture = LoadPicture()
    Resume Next
TrafficFormErr:
    ilErr = True
    ilRet = MsgBox("Shell Execution Failed:" & slStr, 0, "Counterpoint Software")
    Resume Next
ExeTestErr:
    ilRet = False
    Resume Next
End Sub

Private Sub MDIForm_Resize()
    'MsgBox "got a resize command " & Traffic.WindowState
    If lgShellAndWaitID <> 0 And Traffic.WindowState <> vbMinimized Then
        Traffic.WindowState = vbMinimized
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Dim ilShell As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim dlShellRet As Double
    
    igStopCancel = False
    mUnloadForms True
    If igStopCancel Then    'User stop the unload
        Cancel = 1  'Stop unload
        Exit Sub
    End If
    'If igResetSystemDate Then
    '    Date$ = sgSystemDate
    'End If
    ilRet = mSignOnOff(False)
    
    gAddOrUpdateLrf "U", igManualPostVefCode, sgManualPostDate

    tmcCheckAlert.Enabled = False
    tmcClock.Enabled = False
    If hmAlert <> 0 Then
        Call DeleteObject(hmAlert)
    End If
    gEraseGlobalVar True
    If (igRUStartExeCopy) And ((Len(sgDLExePath) > 0) Or (Len(sgDLRptPath) > 0)) Then
        gObtainRUIniValues
        If igRUDialUp Then
            On Error GoTo ExeOkErr
            ilRet = True
            If FileDateTime(sgExePath & "RUExpt.Exe") < FileDateTime(sgRUImportPath & "RUExpt.Exe") Then
                If ilRet Then
                    FileCopy sgRUImportPath & "RUExpt.Exe", sgExePath & "RUExpt.Exe"
                    For ilLoop = 0 To 20 Step 1
                        DoEvents
                    Next ilLoop
                    If ilRet Then
                        Kill sgRUImportPath & "RUExpt.Exe"
                    End If
                End If
            End If
            ilRet = True
            If FileDateTime(sgExePath & "Traffic.Exe") < FileDateTime(sgRUImportPath & "Traffic.Exe") Then
                If ilRet Then
                    'ilShell = Shell(sgExePath & "ExeCopy.Exe /CS-RemoteUser", 1)
                    dlShellRet = Shell(sgExePath & "ExeCopy.Exe /CS-RemoteUser", 1)
                End If
            End If
        Else
            ilFound = False
            If Len(sgDLExePath) > 0 Then
                lbcFile.Path = Left$(sgExePath, Len(sgExePath) - 1)
                lbcFile.Pattern = "*.*"
                For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                    On Error GoTo ExeOkErr
                    ilRet = True
                    If FileDateTime(sgExePath & lbcFile.List(ilLoop)) < FileDateTime(sgDLExePath & lbcFile.List(ilLoop)) Then
                        If ilRet Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilLoop
            End If
            If (Len(sgDLRptPath) > 0) And (Not ilFound) Then
                lbcFile.Path = Left$(sgRptPath, Len(sgRptPath) - 1)
                For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                    On Error GoTo ExeOkErr
                    ilRet = True
                    If FileDateTime(sgRptPath & lbcFile.List(ilLoop)) < FileDateTime(sgDLRptPath & lbcFile.List(ilLoop)) Then
                        If ilRet Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilLoop
            End If
            If ilFound Then
                On Error GoTo ExeOkErr
                ilRet = True
                If FileDateTime(sgExePath & "ExeCopy.Exe") < FileDateTime(sgDLExePath & "ExeCopy.Exe") Then
                    If ilRet Then
                        FileCopy sgDLExePath & "ExeCopy.Exe", sgExePath & "ExeCopy.Exe"
                        For ilLoop = 0 To 20 Step 1
                            DoEvents
                        Next ilLoop
                    End If
                End If
                'ilShell = Shell(sgExePath & "ExeCopy.Exe /CS-RemoteUser", 1)
                dlShellRet = Shell(sgExePath & "ExeCopy.Exe /CS-RemoteUser", 1)
            End If
        End If
    End If



    'Reset used instead of Close to cause # Clients on network to be decrement
    'Orlando meeting- ignore help messages
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
'    btrDestroy hgDB
    gCloseVpf
    On Error Resume Next
    ilRet = btrClose(hgAuf)
    btrDestroy hgAuf
    ilRet = btrClose(hgUlf)
    btrDestroy hgUlf
    ilRet = btrClose(hgChf)
    btrDestroy hgChf
    ilRet = btrClose(hgVsf)
    btrDestroy hgVsf
    ilRet = btrClose(hgSite)
    btrDestroy hgSite
    ilRet = btrClose(hgTmf)
    btrDestroy hgTmf
    btrStopAppl
    'D.S. 07/20/15
    cnn.Close
    rst.Close
    '3/23/10 Dan M made global
'    gCallNetReporter CsiReportCall.FinishReports
    '11/20/09 close csiNetReporter
   ' mQuitNetReporterAsNeeded
    'Shell sgExePath & "csinetreporter.exe /Q"
    '8297
  '  Set ogPdf = Nothing
    Set Traffic = Nothing   'Remove data segment
    End
ExeOkErr:
    ilRet = False
    Resume Next
End Sub
'Dan 3/23/10 made global gcallnetreporter
'Private Sub mQuitNetReporterAsNeeded()
'    Dim slCommandLine As String
'On Error Resume Next
'    If bgReportModuleRunning Then
'        gCallNetReporter
''        slCommandLine = gBuildAlternateAsNeeded
''        slCommandLine = slCommandLine & " /Q"
''        Shell slCommandLine
'    End If
'End Sub


Private Sub mnuAboutHelp_Click()
    Dim iAns As Integer
    Dim ilErr As Integer
    ''sgCSVersion set within InitSubs.Bas-gInitGlobalVar
    ''iAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
    'ilShell = 0
    'ilErr = False
    'On Error GoTo mnuAboutHelpErr
    'If igTestSystem Then
    '    ilShell = Shell(sgExePath & "CSIAbout.Exe Traffic^Test\" & sgUserName, 1)
    'Else
    '    ilShell = Shell(sgExePath & "CSIAbout.Exe Traffic^Prod\" & sgUserName, 1)
    'End If
    'If ilErr Then
    '    Exit Sub
    'End If
    'While GetModuleUsage(ilShell) > 0
    '    ilRet = DoEvents()
    'Wend
    On Error Resume Next
    CSIAbout.Show vbModal
    Exit Sub

    ilErr = True
    iAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
    Resume Next
End Sub


Private Sub mnuAccessoryAcquisitionCheck_Click()
    'AcqCheck.Show vbModal
End Sub

Private Sub mnuAccessoryBlocks_Click()
    BlockVw.Show vbModal
End Sub

Private Sub mnuAccessoryCalculator_Click()
    gFunctionKeyBranch KEYF5
End Sub

Private Sub mnuAccessoryCalendar_Click()
    'Calendar.Show vbModal
    gFunctionKeyBranch KEYF4
End Sub

Private Sub mnuAccessoryDate_Click()
    Dim slDate As String
    Dim slTime As String
    Dim slMsg As String
    Dim slInputDate As String
    Dim ilRet As Integer
    Dim llNowDate As Long
    slDate = Format$(gNow(), "m/d/yy")
    slTime = Format$(gNow(), "h:mm:ss AM/PM")
    slMsg = "Counterpoint Date: " & slDate & "; Counterpoint Time: " & slTime
    'Dan M 1/11/10  allow user to change date.
    'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Or (tgUrf(0).sChangeCSIDate = "I") Then
        'slInputDate = InputBox$(slMsg, "Date/Time", slDate)
        sgGenMsg = slMsg
        sgCMCTitle(0) = "Done"
        sgCMCTitle(1) = "Cancel"
        sgCMCTitle(2) = ""
        sgCMCTitle(3) = ""
        igDefCMC = 0
        igEditBox = 1
        sgEditValue = slDate
        GenMsg.Show vbModal
        If igAnsCMC = 0 Then
            slInputDate = sgEditValue
            If Len(slInputDate) <> 0 Then
                If gValidDate(slInputDate) Then
                    If gDateValue(slDate) <> gDateValue(slInputDate) Then
                        'Add year if missing
                        slInputDate = gAdjYear(slInputDate)
                        'Date$ = slInputDate
                        'igResetSystemDate = True
                        'Dan M 9/20/10 problems in Reports.exe with csiSetName
                        'ilRet = csiSetName("SYSDate", slInputDate)
                        ilRet = gCsiSetName(slInputDate)
                        tmcClock_Timer
                    Else
                        '7/24/08:Retain date if not changed
                        'ilRet = csiSetName("SYSDate", "")
                    End If
                End If
            Else
                '7/24/08:  If date removed, reset to todays date
                'Dan M 9/20/10 problems in Reports.exe with csiSetName
                'ilRet = csiSetName("SYSDate", "")
                ilRet = gCsiSetName("")
                tmcClock_Timer
            End If
        End If
        llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
        sgDateRangeStartDate = Format$(llNowDate - 365, "m/d/yy")
        sgDateRangeEndDate = Format$(llNowDate - 1, "m/d/yy")
    Else
        MsgBox slMsg & "; User: " & tgUrf(0).sName, vbOKOnly + vbInformation, "Date/Time"
    End If
End Sub

Private Sub mnuAccessoryEditor_Click()
    Dim ilShell As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim ilPos As Integer
    Dim ilPFE As Integer

    sgGenMsg = "Select Editor or Cancel"
    sgCMCTitle(0) = "&NotePad"
    sgCMCTitle(1) = "&WordPad"
    ilPFE = False
    ilRet = 0
    On Error GoTo mnuAccessoryEditErr:
    ilPos = InStr(1, sgExePath, "\exe", vbTextCompare)
    If ilPos >= 0 Then
        slDateTime = FileDateTime(Left(sgExePath, ilPos) & "\setup\pfe\pfe32.exe")
        If ilRet = 0 Then
            ilPFE = True
        End If
    End If
    If ilPFE Then
        sgCMCTitle(2) = "&PFE"
        sgCMCTitle(3) = "&Cancel"
    Else
        sgCMCTitle(2) = "&Cancel"
        sgCMCTitle(3) = ""
    End If
    igDefCMC = 0
    igEditBox = 0
    sgEditValue = ""
    On Error Resume Next
    GenMsg.Show vbModal
    If ilPFE Then
        If igAnsCMC = 3 Then
            Exit Sub
        End If
    Else
        If igAnsCMC = 2 Then
            Exit Sub
        End If
    End If
    ilShell = 0
    On Error GoTo mnuEditorFileErr
    'ilShell = Shell("NotePad.Exe", 3)
    'While GetModuleUsage(ilShell) > 0
    '    ilRet = DoEvents()
    'Wend
    If igAnsCMC = 0 Then
        gShellAndWait Traffic, "NotePad.Exe", vbMaximizedFocus, True    'vbTrue
    ElseIf igAnsCMC = 1 Then
        gShellAndWait Traffic, "WordPad.Exe", vbMaximizedFocus, True    'vbTrue
    ElseIf igAnsCMC = 2 Then
        gShellAndWait Traffic, Left(sgExePath, ilPos) & "\setup\pfe\pfe32.exe", vbMaximizedFocus, True  'vbTrue
    End If
    Exit Sub
mnuEditorFileErr:
    If igAnsCMC = 0 Then
        MsgBox "The Editor Notepad.Exe not in Search Path, Please copy into system or system32 folder", vbOKOnly + vbInformation, "Error"
    ElseIf igAnsCMC = 1 Then
        MsgBox "The Editor Wordpad.Exe not in Search Path, Please copy into system or system32 folder", vbOKOnly + vbInformation, "Error"
    ElseIf igAnsCMC = 2 Then
        MsgBox "The Editor PFE32.Exe is Missing", vbOKOnly + vbInformation, "Error"
    End If
    Resume Next
mnuAccessoryEditErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub mnuAccessoryMessages_Click()
    Messages.Show vbModal

End Sub

Private Sub mnuAccessorySetCredit_Click()
    Dim ilRet As Integer
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim slDate As String
    Dim dlShellRet As Double

    ilRet = MsgBox("This should be run nightly on the server.  If run here, the system may be slowed down for several hours.  Ok to Proceed?", vbDefaultButton2 + vbInformation + vbOKCancel, "Set Credit")
    If ilRet = vbOK Then
        slCommandStr = "SetCrdit.Exe /IniLoc:" & CurDir$
        slDate = gCSIGetName()
        If slDate <> "" Then
            'use slDate when writing to file later
            slDate = " /D:" & slDate
            'slCommandStr = slCommandStr & " /D:" & slDate
            slCommandStr = slCommandStr & slDate
        End If
        slCommandStr = slCommandStr & " /ForceRun"
        'ilShell = Shell(sgExePath & slCommandStr, vbMinimizedNoFocus)
        dlShellRet = Shell(sgExePath & slCommandStr, vbMinimizedNoFocus)
    End If
End Sub

Private Sub mnuAccessoryShowCredits_Click()
    On Error Resume Next
    'Modeless required so that Collect can be call
    If igShowCredit = False Then
        ShoCrdit.Show vbModeless    'vbModal
    Else
        ShoCrdit.SetFocus
    End If
End Sub

Private Sub mnuAccessorySportCheck_Click()
    'On Error Resume Next
    'SportChk.Show vbModal
End Sub

Private Sub mnuAccessoryUserStatus_Click()
    UsersLog.Show vbModal
End Sub


Private Sub mnuAccessoryUtilitiesCFFCheck_Click()
    'Dim slStr As String
    'If igTestSystem Then
    '    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'Else
    '    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'End If
    'sgCommandStr = slStr
    'On Error Resume Next
    'CffCheck.Show vbModal
    'gChDrDir
End Sub

Private Sub mnuAccessoryUtilitiesContractSwapCheck_Click()
    'ContractSwapCheck.Show vbModal
End Sub

Private Sub mnuAccessoryUtilitiesInvCheck_Click()
    'Dim slStr As String
    'If igTestSystem Then
    '    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'Else
    '    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'End If
    'sgCommandStr = slStr
    'On Error Resume Next
    'InvCheck.Show vbModal
    'gChDrDir
End Sub

Private Sub mnuAccessoryUtilitiesSetInvDate_Click()
    'SetCopyDates.Show vbModal
End Sub

Private Sub mnuAccessoryUtilitiesSMFCheck_Click()
    'Dim slStr As String
    'If igTestSystem Then
    '    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'Else
    '    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'End If
    'sgCommandStr = slStr
    'On Error Resume Next
    'SMFCheck.Show vbModal
'    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)
'        ChDir sgCurDir
'    End If
    'gChDrDir

End Sub

Private Sub mnuAccessoryUtilities_Click()
    Dim ilShell As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    

    'dlShellRet = Shell(sgExePath & "ExportProj.Exe /IniLoc:" & CurDir$ & " /UserInput", vbMinimizedNoFocus)
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    slCommandStr = slCommandStr & "/IniLoc:" & CurDir$ & " /UserInput"
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait Traffic, sgExePath & "Utilities.exe " & slCommandStr, vbNormalFocus, False    'vbFalse
    
    Traffic.WindowState = vbMaximized


End Sub

Private Sub mnuAlert_Click()
    If igSportsSystem = 2 Then
        Exit Sub
    End If
    If tmcInitAlert.Enabled Then
        Exit Sub
    End If
    'igAlertSpotStatus = 0
    'igAlertCntrStatus = 0
    AlertVw.Show vbModal
    If (igAlertCntrStatus >= 1) And (igAlertCntrStatus <= 5) Then
        DoEvents
        Sleep 50
        If Not igJobShowing(CONTRACTSJOB) Then
            If igAlertCntrStatus = 1 Then
                sgCntrScreen = "Proposals"  'Change Complete
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igAlertCntrStatus = 2 Then   'Change Rev Complete
                sgCntrScreen = "Proposals"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igAlertCntrStatus = 3 Then   'View mode
                sgCntrScreen = "Proposals"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igAlertCntrStatus = 4 Then   'Change Not Schedule
                sgCntrScreen = "Orders"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igAlertCntrStatus = 5 Then   'View mode
                sgCntrScreen = "Orders"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            End If
        Else
            igAlertCntrStatus = -1
        End If
    Else
        'If (igAlertSpotStatus = 2) And (Trim$(sgAlertSpotVehicle) <> "") And (Trim$(sgAlertSpotMoDate) <> "") Then
        If (igAlertSpotStatus = 1) And (Trim$(sgAlertSpotVehicle) <> "") And (Trim$(sgAlertSpotMoDate) <> "") Then
            DoEvents
            Sleep 50
            'igAlertSpotStatus = 1
            Spots.Show vbModeless
        End If
    End If
End Sub







Private Sub mnuBackupTrafficAndAffiliate_Click()

    On Error GoTo BuErr
    BUZip.Show vbModal
    ' BUZip.mZipMain False, 2   '2 = Traffic and Affiliate
    Exit Sub

BuErr:
    MsgBox "An Error has occured in Traffic.frm - mnuBackupTrafficAndAffiliate_Click. Backup may not have completed."

End Sub

Private Sub mnuBarter_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    ExpGPBarter.Show vbModal
    mExport "Barter Payment"
     
End Sub

Private Sub mnuCNCAP_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpCnCAP.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Clearance n Compensation /Sub: AP"
     
End Sub

Private Sub mnuCNCNI_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpCnCNI.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Clearance n Compensation /Sub: NI"
     

End Sub

Private Sub mnuCNCSA_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpCnCSA.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Clearance n Compensation /Sub: SA"
     

End Sub

Private Sub mnuCNCSS_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpCnCSS.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Clearance n Compensation /Sub: SS"
     

End Sub
'Dan M removed 6/16/10
'Private Sub mnuContentHelp_Click()
'    cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    cdcSetup.HelpCommand = cdlHelpIndex
'    cdcSetup.Action = 6
'End Sub

Private Sub mnuDate_Click()
    Calendar.Show vbModal
End Sub



Private Sub mnuEfficioProj_Click()
    mExport "Efficio Projections /Sub: Efficio Projections"
End Sub

Private Sub mnuEfficioRev_Click()
    mExport "Efficio Revenue /Sub: Efficio Revenue"
End Sub

Private Sub mnuExpCmmlChg_Click()
    Dim slStr As String
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    If Not igVehicleDefined Then
        MsgBox "Please Define Vehicle", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ExpCmChg.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExpCmChg.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    
    mExport "Commercial Change"
     

End Sub

Private Sub mnuExpCntrLine_Click()
    mExport "CntrLine"
     
End Sub

Private Sub mnuExportAccount_Click()
'    Dim slStr As String
''    Dim ilLoop As Integer
''    Dim ilRet As Integer
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ExpACC.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExpACC.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mExport "Accounting"
     
End Sub

Private Sub mnuExportAdServerBillDisc_Click()
    mExport "AdServerBillDisc"
    
End Sub

Private Sub mnuExportAirWave_Click()
    mExport "AirWave"

End Sub

Private Sub mnuExportAuto_Click()
'    Dim slStr As String
''    Dim ilLoop As Integer
''    Dim ilRet As Integer
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ExptGen.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExptGen.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mExport "Automation"
     
End Sub

Private Sub mnuExportBkCpy_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpBkCpy.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Copy Bulk Feed"
     

End Sub

Private Sub mnuExportCarts_Click()
    mExport "Carts"
End Sub

Private Sub mnuExportCash_Click()
    mExport "Cash Receipts"
End Sub

Private Sub mnuExportCorporate_Click()
    Dim ilShell As Integer
    Dim dlShellRet As Double

    dlShellRet = Shell(sgExePath & "ExportProj.Exe /IniLoc:" & CurDir$ & " /UserInput", vbMinimizedNoFocus)

    gChDrDir
End Sub

Private Sub mnuExportCRE_Click()
    mExport "CustomRevenueExport"
End Sub

Private Sub mnuExportDallas_Click()
'    Dim slStr As String
''    Dim ilLoop As Integer
''    Dim ilRet As Integer
''    If Not igSiteSet Then
''        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
''        Exit Sub
''    End If
''    If Not igVehicleDefined Then
''        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
''        Exit Sub
''    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ExpDall.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpDall.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Dallas Feed"
     

End Sub

Private Sub mnuExportEnco_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExptEnco.Show vbModal
'    gChDrDir

    mExport "Enco"
     
End Sub

Private Sub mnuExportInvReg_Click()

    mExport "Invoice Register"
     
End Sub

Private Sub mnuExportMatrix_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExpMatrix.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Matrix"
     

End Sub



Private Sub mnuExportMK_Click()
    mExport "Miller Kaplan"
End Sub

Private Sub mnuExportMP2_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExptMP2.Show vbModal
'    gChDrDir
   ' mExport "Audio MP2"
   '7496
    mExport "Audio " & UCase(Mid(sgAudioExtension, 2))
End Sub

Private Sub mnuExportNYASP_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ASP"
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ASP"
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpNY.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If

    mExport "Engineering Feed /Sub: ASP"
     

End Sub



Private Sub mnuExportNYESPN_Click()
    mExport "Engineering Feed /Sub: ESPN"
End Sub

Private Sub mnuExportPhoenix_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'    On Error Resume Next
'    sgCommandStr = slStr
'    ExpPhnx.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir

    mExport "Phoenix Log"
     

End Sub

Private Sub mnuExportRAB_Click()
    mExport "RAB"
End Sub

Private Sub mnuExportRevenue_Click()
'    Dim slStr As String
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    ExpRevenue.Show vbModal
'    gChDrDir
    
    mExport "Revenue"
     

End Sub

Private Sub mnuExportStnFeed_Click()
    mExport "Station Feed"
End Sub

Private Sub mnuExportTableau_Click()
    mExport "Tableau"
End Sub

Private Sub mnuExportWOInvoice_Click()
    mExport "WO Invoice"
End Sub

Private Sub mnuGetPaid_Click()
    Dim ilRet As Integer
    Dim ilShell As Integer
    Dim dlShellRet As Double
    
    ilRet = MsgBox("This feature can be setup to run automatically.  Click Ok to Proceed", vbDefaultButton2 + vbInformation + vbOKCancel, "GetPaid Export")
    If ilRet = vbOK Then
        'ilShell = Shell(sgExePath & "GetPaid.Exe /IniLoc:" & CurDir$, vbMinimizedNoFocus)
        dlShellRet = Shell(sgExePath & "GetPaid.Exe /IniLoc:" & CurDir$, vbMinimizedNoFocus)
   End If
End Sub

Private Sub mnuGreatPlainsGL_Click()
'    Dim slStr As String
'
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    ExpGP.Show vbModal

    mExport "Great Plains G/L"
     
End Sub

Private Sub mnuHelpGoToWeb_Click()
    'Dan M 6/16/10 go to website
    'Shell "cmd /c start http://www.counterpoint.net/clientsLogin.php?error=4", vbHide
    'Dan M 11/19/11 this connects to menu "counterpoint web site", which contains sub menus: see below:
'    ShellExecute 0&, vbNullString, "www.counterpoint.net/clientsLogin.php?error=4", vbNullString, vbNullString, vbNormalFocus

End Sub
Private Sub mnuDocumentation_Click()
  '  WebConnect.Show vbModal
    '10482
        
    'URL_Documentation added 2024-01-09 JJB to allow for a dynamic location for documentation
    'ShellExecute 0&, vbNullString, "www.counterpoint.net/clients/DocumentationNew.php?CsiNo=4307&CsiCo=" & Trim(tgSpf.sGClient), vbNullString, vbNullString, vbNormalFocus
    If sgURL_Documentation = "" Then
        MsgBox "The URL for the documentation is not currently configured on this system.  Please contact Counterpoint support.", vbInformation, "Counterpoint Documentation Not Found"
    Else
        ShellExecute 0&, vbNullString, sgURL_Documentation & "?CsiNo=4307&CsiCo=" & Trim(tgSpf.sGClient), vbNullString, vbNullString, vbNormalFocus
    End If
   
End Sub

Private Sub mnuImportImpressions_Click()
    mImportResearch "Impressions"

End Sub

Private Sub mnuImportMegaphone_Click()
    '11010
    Dim slImportPath As String
    
    slImportPath = sgExePath & "MegaPhoneImport.exe"
    If gFileExist(slImportPath) = FILEEXISTS Then
        Screen.MousePointer = vbDefault
        Traffic.WindowState = vbMinimized
        gShellAndWait Me, slImportPath, vbNormalFocus, True
        Traffic.WindowState = vbMaximized
    Else
        gMsgBox "Megaphone Import does not exist on your system.  Contact Counterpoint for installion instructions.", vbOKOnly, "Megaphone Import"
    End If
End Sub

Private Sub mnuImportVehicle_Click()
'    Dim slStr As String
'    If Not igSiteSet Then
'        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not igVehicleDefined Then
'        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not gRecExistForFile("Rdf.btr") Then
'        MsgBox "Please Define Dayparts", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ImptRad.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ImptVeh.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mImportResearch "Vehicle"

End Sub

Private Sub mnuSendToEDS_Click()

    ''07-13-15
    'If bgEDSIsActive Then
    '    EDSLink.Show vbModal
    'Else
    '    MsgBox "The E-mail Distribution System is not activated."
    'End If

End Sub

Private Sub mnuViewSql_Click()
    ''6240
    'Dim slPath As String
    'Dim myConnection As CCsiSystemConnection
    'Dim slDSN As String
   '
    'slPath = sgExePath & "ViewSql.exe"
    ''8886
   '' If Dir(slPath) <> "" Then
   'If gFileExist(slPath) = FILEEXISTS Then
   '     Set myConnection = New CCsiSystemConnection
   '     slDSN = myConnection.GlobalDSNName
   '     If Len(slDSN) = 0 Then
   '         slDSN = "BLANK"
   '     End If
   '     Shell slPath & " " & slDSN & " , " & sgDBPath, vbNormalFocus
   ' Else
   '     gMsgBox "ViewSql does not exist on your system.  Contact Counterpoint for installion instructions.", vbOKOnly, "ViewSql"
   ' End If
End Sub

Private Sub mnuMK_Click()
    mExport "Miller Kaplan"
End Sub

Private Sub mnuWebHome_Click()
    'ShellExecute 0&, vbNullString, "www.counterpoint.net", vbNullString, vbNullString, vbNormalFocus
    WebConnect.bmHomePage = True
    WebConnect.Show vbModal
End Sub
Private Sub mnuImportAuto_Click()
    Dim slStr As String
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'Traffic.edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    'lgShellRet = Shell(sgExePath & "ImptGen.Exe " & slStr, 1)
    'Traffic.Enabled = False
    ''Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    ImptGen.Show vbModal
'    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)
'        ChDir sgCurDir
'    End If
    gChDrDir


    'slStr = sgDoneMsg
    'Traffic.Enabled = True
    'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
End Sub
Private Sub mnuImportIMS_Click()
'    Dim slStr As String
    Dim slStr As String


''    Dim slStr As String
''    Dim ilLoop As Integer
''    Dim ilRet As Integer
''    If Not igSiteSet Then
''        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
''        Exit Sub
''    End If
''    If Not igVehicleDefined Then
''        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
''        Exit Sub
''    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ImptIMS.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    ImptIMS.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
        mImportResearch "CSV"

End Sub
Private Sub mnuImportMark_Click()
'    Dim slStr As String
''    Dim ilLoop As Integer
''    Dim ilRet As Integer
'    If Not igSiteSet Then
'        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not igVehicleDefined Then
'        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ImptMark.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ImptMark.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mImportResearch "ACT 1"

End Sub

Private Sub mnuImportRadar_Click()
'    Dim slStr As String
'    If Not igSiteSet Then
'        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not igVehicleDefined Then
'        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not gRecExistForFile("Rdf.btr") Then
'        MsgBox "Please Define Dayparts", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ImptRad.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ImptRad.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mImportResearch "RADAR"

End Sub



Private Sub mnuImportSatellite_Click()
'    Dim slStr As String
'    If Not igSiteSet Then
'        MsgBox "Please Define Site Options", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not igVehicleDefined Then
'        MsgBox "Please Define Vehicle", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
'    If Not gRecExistForFile("Rdf.btr") Then
'        MsgBox "Please Define Dayparts", vbOkOnly + vbInformation, "Counterpoint"
'        Exit Sub
'    End If
''    'Screen.MousePointer = vbHourGlass  'Wait
''    igChildDone = False
''    Traffic.edcLinkSrceHelpMsg.Text = ""
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
''    lgShellRet = Shell(sgExePath & "ImptRad.Exe " & slStr, 1)
''    Traffic.Enabled = False
''    'Screen.MousePointer = vbDefault  'Wait
''    Do While Not igChildDone
''        DoEvents
''    Loop
''    slStr = sgDoneMsg
''    Traffic.Enabled = True
''    Traffic.edcLinkSrceHelpMsg.Text = "Ok"
''    For ilLoop = 0 To 10
''        DoEvents
''    Next ilLoop
'    sgCommandStr = slStr
'    On Error Resume Next
'    ImptSat.Show vbModal
''    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''    If InStr(1, sgCurDir, ":") > 0 Then
''        ChDrive Left$(sgCurDir, 2)
''        ChDir sgCurDir
''    End If
'    gChDrDir
    mImportResearch "Satellite"

End Sub
Private Sub mnuInvoiceExport_Click()
'   removed 5-11-17
'    Dim slStr As String
'    If igTestSystem Then
'        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    Else
'        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
'    End If
'
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExpInv.Show vbModal
'
'    gChDrDir

'    mExport "Invoice"
     

End Sub

Private Sub mnuISCIXReference_Click()
    'ExpISCIXRef.Show (vbModal)
    mExport "Audio ISCI Title"
     
End Sub

Private Sub mnuRemoteUser_Click()
    Dim slStr As String
    'igChildDone = False
    'Traffic.edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName '& "\" & Trim$(Str$(CALLNONE))
    Else
        slStr = "Traffic^Prod\" & sgUserName '& "\" & Trim$(Str$(CALLNONE))
    End If
    'lgShellRet = Shell(sgExePath & "RUExpt.Exe " & slStr, 1)
    'Traffic.Enabled = False
    ''Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    'RUExpt.Show vbModal
    slStr = sgDoneMsg
    'Traffic.Enabled = True
    'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    If (InStr(1, slStr, "Done-ExeCopy", 1) > 0) And (tgUrf(0).iRemoteUserID > 0) Then
        igRUStartExeCopy = True
    End If
    'Get new Block Date/Time (1/6/2000)
    sgSpfStamp = "~"    'Force read
    gSpfRead
    'gSetUsingTraffic
End Sub

Private Sub mnuResetWorkFiles_Click()
    Call mClearReportFiles
End Sub
'Dan removed 6/16/10
'Private Sub mnuSearchHelp_Click()
'    cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    cdcSetup.HelpKey = ""
'    cdcSetup.HelpCommand = cdlHelpPartialKey
'    cdcSetup.Action = 6
'End Sub

Private Sub mnuTaskExit_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************




    Unload Traffic
    If Not igStopCancel Then    'User stop the unload
        'Set Traffic = Nothing   'Remove data segment
    End If
End Sub

Private Sub mnuTaskJob_Click()
    If igSportsSystem <> 0 Then
        If igBasic10Showing Then
            Sports.SetFocus
        Else
            Sports.Show vbModeless
        End If
        pbcMsgArea.Visible = False
        mDisableMenu
    Else
        If igBasic10Showing Then
            Basic10.SetFocus
        Else
            Basic10.Show vbModeless
        End If
    End If
End Sub

Private Sub mnuTaskList_Click()
    If igAuxWndShowing Then
        AuxWnd.SetFocus
    Else
        AuxWnd.Show vbModeless
    End If
End Sub

Private Sub mnuTrafficDataToTest_Click()
    BUDataToTest.Show vbModal
End Sub
'Dan removed 6/16/10
'Private Sub mnuUsingHelp_Click()
''    cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    cdcSetup.HelpCommand = cdlHelpHelpOnHelp
'    cdcSetup.Action = 6
'End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulateListBoxes              *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate all list boxes on      *
'*                     traffic form.  This list boxes  *
'*                     used to populate list boxes on  *
'*                     forms.                          *
'*                                                     *
'*******************************************************
Private Sub mPopulateListBoxes()
'
'   mPopulateListBoxes
'   Where:
'
'    On Error GoTo PopListErr
'    ilRet = gIPopListBox(traffic, traffic.lbcVehicle, "Vef.btr", 2, 20, "", 0)   'Vehicle
'    On Error GoTo 0
    Dim ilRet As Integer


    ilRet = gObtainRcfRifRdf()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainAdvt()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainAgency()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainVef()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainComp()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainAvail()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainSAF()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("R", sgRevSetStamp, tgRevSet())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("O", sgShareBudgetStamp, tgShareBudget())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("X", sgExclMnfStamp, tgExclMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("B", sgBusCatMnfStamp, tgBusCatMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("P", sgPotMnfStamp, tgPotMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("D", sgDemoMnfStamp, tgDemoMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainMnfForType("F", sgSocEcoMnfStamp, tgSocEcoMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    
    '11/6/14: Preload stations
    ilRet = gObtainStations()
    
    'ilRet = gObtainMnfForType("H3", sgMktMnfStamp, tgMktMnf())
    'If ilRet = False Then
    '    imTerminate = True
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    'sgMktClusterDef = "N"
    'If tgSpf.sMktBase = "Y" Then
    '    For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
    '        If Trim$(tgMktMnf(ilLoop).sRPU) = "Y" Then
    '            sgMktClusterDef = "Y"
    '            Exit For
    '        End If
    '    Next ilLoop
    'Else
    '    sgMktClusterDef = "N"
    'End If

    ilRet = mGetDBID()

    gAnyClustersDef
    gAnyRepDef
    
    gObtainMissedReasonCode
    
    mCheckAndPopEDIServiceNames
    
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTrafficStartUp                 *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Start up                        *
'*                                                     *
'*******************************************************
Private Function mTrafficStartUp() As Integer
    Dim slVehType As String
    Dim ilRet As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim llNoRec As Long
    Dim ilSetTest As Integer    '0=Test; 1=Set
    Dim llNow As Long
    Dim slStr As String
    Dim slNow As String
    Dim ilLoop As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llDate As Long
    Dim slAppName As String
    Dim ilPos As Integer
    Dim llNowDate As Long
    Dim blSendGGWarning As Boolean
    Dim ilVff As Integer
    '5676
    Dim slRootPath As String
    Dim slPhotoPath As String
    Dim slBitmapPath As String
    Dim slLogoPath As String
    Dim slTemp As String
    Dim slDSN As String
    '8200
    Dim slLocation As String
    '2/19/18: Add timeout setting
    Dim slTimeOut As String
    Dim ilTimeOut As Integer
    '4/15/20: Add affiliate as Option
    Dim slSQLQuery As String
    Dim blShowAffiliate As Boolean
    'Set objEmailSender = New EmailSender    'Initialize Sender and Email
    'Set objMyEmail = New Email

    'Dim objReg As KeyCodeRegistration
    'Set objReg = New KeyCodeRegistration
    'Call objReg.EnterKeyCode("QwBvAHUAbgB0AGUAcgBwAG8AaQBuAHQAIABTAG8AZgB0AHcAYQByAGUAAABEAEIARQBBADAANQAxADUAOQA3AEQAQQA3ADUAQQAzADMANwA5ADQAQwBFADgARgBFADYANwA4ADAAMAAwAEIANgBEAEYARgBFAEEAMQAxADgAOQAzAEUAMAAyADEARABCAEEAQwA0ADkARgBCADMAQgBEADEAMAA3ADUANABCAEEANwBBADAANQBEAEIARQAzADQAMQA2ADUAMgA5AEIAQQA5ADkARAA0ADYAMgBEADUARABBADIAMgBCADMARQAwADAAQwA0ADAARAAxAEUAOABCAEMAOQBBAEIAQwBBAA==")

    'ilRet = 0
    'On Error GoTo mTrafficStartUpErr:
    'slDateTime1 = FileDateTime(sgExePath & "Traffic.Exe")
    'If ilRet = 0 Then
    '    sgBrannerMsg = sgBrannerMsg & " (Made:" & Format$(slDateTime1, "m/d/yy") & ") "
    'End If
    'Traffic.Caption = sgBrannerMsg
        
    'D.S. 07/20/15 Startup Relational/SQL Engine
    '8200
    slLocation = "Locations"
    If igTestSystem Then
        slLocation = "TestLocations"
    End If
    If Not gLoadOptionTrafficThenAffiliate(slLocation, "Name", sgDatabaseName) Then
        gMsgBox "Traffic.Ini [Locations] 'Name' key is missing. Traffic will now shutdown.", vbCritical
        mTrafficStartUp = False
        Exit Function
    End If
    Set cnn = New ADODB.Connection
    slDSN = sgDatabaseName
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & slDSN
    '7908
    'On Error GoTo mTrafficStartUpErr:
    On Error GoTo ERRNOLOCATIONNAME
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & slDSN
    End If
    '2/19/18: Add timeout setting
    ilTimeOut = 600
    If gLoadOptionTrafficThenAffiliate(slLocation, "TimeOut", slTimeOut) Then
        ilTimeOut = Val(slTimeOut)
    End If
    cnn.CommandTimeout = ilTimeOut
    '7908
    On Error GoTo mTrafficStartUpErr:
    Set rst = New ADODB.Recordset

    blSendGGWarning = False
    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"
    ilRet = 0
    'On Error GoTo mTrafficStartUpErr:
    'slDateTime1 = FileDateTime(sgExePath & "Traffic.Exe")
    ilRet = gFileExist(sgExePath & slAppName)
    If ilRet = 0 Then
        slDateTime1 = gFileDateTime(sgExePath & slAppName)
        sgCSVersion = sgCSVersion & " created " & Format$(slDateTime1, "m/d/yy") & " at " & Format$(slDateTime1, "h:mm:ssAM/PM")
    End If
    'cmcBasicWnd.Visible = True
    'plcHelp.Visible = True
    'plcDate.Visible = True
    'plcTime.Visible = True
    'cmcAuxWnd.Visible = True
    'If want pop, then move after erase Global call
    'mPopulateListBoxes   'Populate list boxes which are used to retain the names
    'If imTerminate Then
    '    mTrafficStartUp = False
    '    'Unload Traffic
    '    'Set Traffic = Nothing   'Remove data segment
    '    Exit Function
    'End If
    igNoMinForms = 2    'Traffic; MstPict;
    MstPict.Show vbModeless
    'Mstpict.Move Mstpict.Left, Mstpict.Top, pbcMsgArea.Width, pbcMsgArea.Top
    If igTaskLoc = 0 Then
        MstPict.Move 0, 0, pbcMsgArea.Width, Screen.height
    Else
        MstPict.Move 0, 0, pbcMsgArea.Width, pbcMsgArea.Top
    End If
    If Not igBypassFileSize Then
        gTestRecLengths
    End If
    gHlfRead "JOB", tgJobHelp()
    gHlfRead "LST", tgListHelp()


    gSpfRead
    If igExitTraffic Then
        Unload Traffic
        Set Traffic = Nothing   'Remove data segment
        Exit Function
    End If
    'gSetUsingTraffic

    Screen.MousePointer = vbHourglass  'Wait

    hmRedAlertBitmap = mCopyPictureImage(IconTraf!pbcRedAlert)
    hmWhiteAlertBitmap = mCopyPictureImage(IconTraf!pbcWhiteAlert)

'    For imMonitor = 0 To 13 Step 1
'        hmMonitorBitmap(imMonitor) = mCopyPictureImage(IconTraf!pbcMonitor_Green(imMonitor))
'        mShowMonitor hmMonitor(imMonitor), hmMonitorBitmap(imMonitor), imMonitor
'    Next imMonitor
'    IconTraf!pbcMonitor.BackColor = DARKGREEN
'    IconTraf!pbcMonitor.ForeColor = vbWhite
'    IconTraf!pbcMonitor.CurrentX = 15
'    IconTraf!pbcMonitor.CurrentY = 15
'    IconTraf!pbcMonitor.Print "CSS"
'    imMonitor = 0
'    hmMonitorBitmap(imMonitor) = mCopyPictureImage(IconTraf!pbcMonitor)
'    mShowMonitor hmMonitor(imMonitor), hmMonitorBitmap(imMonitor), imMonitor
    mInitMonitorMenu
    
'    If tgSpf.iGAlertInterval <> 0 Then
        'hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        'ilRet = btrOpen(hgAuf, "", sgDBPath & "AUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    MsgBox "Unable to Open Alert File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        'End If
        'hgUlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        'ilRet = btrOpen(hgUlf, "", sgDBPath & "ULF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    MsgBox "Unable to Open User Log File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        'End If
        hgChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hgChf, "", sgDBPath & "CHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open Contract File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        hgVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hgVsf, "", sgDBPath & "VSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open Contract Vehicle File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
        End If
        gOpenVpf
'    End If

'    If (tgUrf(0).sCompAlert = "Y") Or (tgUrf(0).sReprintLogAlert = "Y") Then
'        If tgSpf.iGAlertInterval > 0 Then
'            igAlertTimer = tgSpf.iGAlertInterval
'        Else
'            igAlertTimer = 5
'        End If
'        tmcCheckAlert.Enabled = True
'        igAlertFlash = -1
'    End If


    'Field used as Number of spaces in windowed envelope
'    If tgSpf.sExport = "N" Then
'        mnuFileExport.Enabled = False
'    End If
'    If tgSpf.sImport = "N" Then
'        mnuFileImport.Enabled = False
'    End If
    'If tgSpf.sUsingMatrix = "N" Then           '4-24-03
    If ((Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT) And ((Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL) Then
        mnuExportMatrix.Enabled = False
    End If
    '7-9-15 Tableau added, same as Matrix format
    If ((Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL) Then
        mnuExportTableau.Enabled = False
    End If

    If (Asc(tgSpf.sUsingFeatures) And REVENUEEXPORT) <> REVENUEEXPORT Then
        mnuExportCorporate.Enabled = False
    End If

    If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) <> GREATPLAINS Then
        mnuGreatPlainsGL.Enabled = False
        mnuBarter.Enabled = False           '8-24-10
    End If

    If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) <> GETPAIDEXPORT) Then
        mnuGetPaid.Enabled = False
    End If

    '5-11-17 disable feature
'    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) <> INVEXPORTPARAMETERS) Then
'        mnuInvoiceExport.Enabled = False
'    End If

    If ((Asc(tgSpf.sUsingFeatures7) And EXPORTREVENUE) <> EXPORTREVENUE) Then
        mnuExportRevenue.Enabled = False
    End If

    If (Asc(tgSaf(0).sFeatures1) And EFFICIOEXPORT) = EFFICIOEXPORT Then        '6-6-14
        mnuEfficio.Enabled = True
    End If
    
    If (Asc(tgSaf(0).sFeatures6) And RABCALENDAR) = RABCALENDAR Or (Asc(tgSaf(0).sFeatures7) And RABSTD) = RABSTD Or (Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) = RABCALSPOTS Then      '6-19-20 add std and cal spots
        mnuExportRAB.Enabled = True
    Else
        mnuExportRAB.Enabled = False
    End If
    
    '3/19/21 for TTP 9992
    If (Asc(tgSaf(0).sFeatures7) And CUSTOMEXPORT) = CUSTOMEXPORT Then
        mnuExportCRE.Enabled = True
    Else
        mnuExportCRE.Enabled = False
    End If
    
    '3/19/21 - Ad Server Billing Discrepancy Report menu
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        mnuExportAdServerBillDisc.Enabled = True
    Else
        mnuExportAdServerBillDisc.Enabled = False
    End If

    If tgSpf.sDemoEstAllowed <> "Y" Then
        mnuImportSatellite.Enabled = False
    Else
        mnuImportMark.Enabled = False
        mnuImportIMS.Enabled = False
        mnuImportRadar.Enabled = False
    End If
    If tgSpf.sGUsePropSys <> "Y" Then
        mnuImportResearch.Enabled = False
    End If
    'TTP 10205 - 6/21/21 - JW - WO Invoice Export - using SPFX Extended Site Features : spfxInvExpFeature
    'If (tgSpfx.iInvExpFeature And 1) > 0 Then
    If (tgSpfx.iInvExpFeature And INVEXP_AUDACYWO) > 0 Then
        mnuExportWOInvoice.Enabled = True
    Else
        mnuExportWOInvoice.Enabled = False
    End If
    
    'If (tgSpfx.iInvExpFeature And 2) > 0 Then
    If (tgSpfx.iInvExpFeature And INVEXP_AUDACYLINE) > 0 Then
        mnuExpCntrLine.Enabled = True
    Else
        mnuExpCntrLine.Enabled = False
    End If
    
    

    '6/16/09:  Removed menu item Import Contracts because out of memory
    'If (tgSpf.sImptCntr <> "Y") And (tgSpf.sImptCntr <> "N") Then           '11-9-04
    '    mnuImportOrder.Enabled = False
    'End If

    Screen.MousePointer = vbHourglass  'Wait
    gEraseGlobalVar False 'Clear Tag's: retain sgUrfStamp and Help files

    Screen.MousePointer = vbHourglass  'Wait
    'Moved prior to setting the date (Form_Load)
    'ilRet = csiSetAlloc("NAMES", 0, 2)
    'lgDateCreditShown = 0   'Reshow credit alert
    igShowCreditShown = False
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainUrf()
    gUrfRead Traffic, sgUserName, True, tgUrf(), False  'Obtain user records (and set sgCPName and sgSUName)
    sgUserName = Trim$(tgUrf(0).sName)
    ilRet = csiSetName("CPName", sgCPName)
    ilRet = csiSetName("SUName", sgSUName)
    If igRUDirectMode = 1 Then
        tgUrf(0).iRemoteUserID = 0
        'Disallow Contract Projections in Direct Mode (1/4/2000)
        igWinStatus(PROPOSALSJOB) = 1
    ElseIf (igRUDirectMode = 2) And (tgUrf(0).iRemoteUserID > 0) Then
        igWinStatus(PROPOSALSJOB) = 1
    End If
    gObtainVehicleName tgUrf(0).iDefVeh, sgUserDefVehicleName, slVehType
    Screen.MousePointer = vbHourglass  'Wait
    gObtainSalespersonName tgUrf(0).iSlfCode, sgUserSalespersonName
    ilRet = csiSetName("SLSPName", sgUserSalespersonName)
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gVpfRead()    'Obtain preference table
    Screen.MousePointer = vbHourglass  'Wait
    'gObtainVehComboList    'Not used by contract
    'gObtainVirtVehList     'Not used by contract
    'merged all software requires that this call be made as it was in start
    gObtainVirtVehList

    gObtainBooksForEstByUSA
    ' Dan M not doing anything
   ' gInitSuperUser tgUrf(0)
    Screen.MousePointer = vbHourglass  'Wait
    If tgUrf(0).sClnMoYr = "M" Then
        igCalByYear = False
    Else
        igCalByYear = True
    End If
    Select Case tgUrf(0).sClnType
        Case "S"
            igCalMonthType = 0
        Case "R"
            igCalMonthType = 1
        Case "U"
            igCalMonthType = 2
        Case "D"
            igCalMonthType = 3
        Case "C"
            igCalMonthType = 4
    End Select
    If tgUrf(0).sClnLayout = "A" Then
        igCalAcross = True
    Else
        igCalAcross = False
    End If
    'mUpdateCredit
    igFromSignon = True
    If igDemo Then
        igShowPicture = True    'remove this line to get Jog screen up first instead of picture
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        mnuRemoteUser.Enabled = False
    Else
        mnuRemoteUser.Enabled = True
        mnuFileExport.Enabled = False
        mnuFileImport.Enabled = False
    End If

    fgCalLeft = tgUrf(0).iClnLeft
    fgCalTop = tgUrf(0).iClnTop
    fgCalcLeft = tgUrf(0).iClcLeft
    fgCalcTop = tgUrf(0).iClcTop
    igUpdateAllowed = True
    
    'If ((Asc(tgSpf.sUsingFeatures10) And CONTRACTVERIFY) <> CONTRACTVERIFY) Or (tgUrf(0).sCntrVerify <> "I") Then
    '    mnuTaskSpace4.Visible = False
    '    mnuCntrVerify.Visible = False
    'End If
    ''If (Not igShowPicture) Or (igRemoteUser) Or (igTestSystem) Then
    ''    Basic10.Show vbModeless
    ''End If
    ''Test for Log.Exe file
    ''
    ''In signOn special password: "8373191n"
    ''where n is 0 to 9
    ''
    'igGGFlag = -1
    'ilSetTest = 0
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        igGGFlag = 1
        ilSetTest = 1
    Else
        igGGFlag = -1
        ilSetTest = 0
    End If
    ilRet = 0
    If igGGBDSignOn Then
        ilSetTest = 1
    Else
        On Error GoTo mTrafficStartUpErr:
        slDateTime1 = FileDateTime(sgExePath & "Log.Exe")
        If ilRet = 0 Then
            ilSetTest = 1
            'Start
            Kill sgExePath & "Log.Exe"
        Else
            ilRet = 0
            On Error GoTo mTrafficStartUpErr:
            slDateTime1 = FileDateTime(sgExePath & "Spot.Exe")
            If ilRet = 0 Then
                ilSetTest = 1
                'Start
                Kill sgExePath & "Spot.Exe"
            End If
        End If
    End If
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        imSafRecLen = Len(tmSaf)  'btrRecordLength(hlUrf)  'Get and save record length
        If imSafRecLen = btrRecordLength(hmSaf) Then
            llNoRec = btrRecords(hmSaf) 'Obtain number of records
            If llNoRec = 0 Then
                tmSaf.sSchRNG = "T"
                tmSaf.sSchMaL = "H"
                tmSaf.sSchMdL = "D"
                tmSaf.sSchMiL = "Q"
                tmSaf.sSchCycle = "1"
                tmSaf.sSchMove = "N"
                tmSaf.sSchCompact = "N"
                tmSaf.sSchPreempt = "E"
                tmSaf.sSchHour = "B"
                tmSaf.sSchDay = "B"
                tmSaf.sSchQH = "B"
                tmSaf.sName = ""
                tmSaf.iVefCode = 0
                tmSaf.lLowPrice = 0
                tmSaf.lLevelToPrice(0) = 0
                tmSaf.lLevelToPrice(1) = 0
                tmSaf.lLevelToPrice(2) = 0
                tmSaf.lLevelToPrice(3) = 0
                tmSaf.lLevelToPrice(4) = 0
                tmSaf.lLevelToPrice(5) = 0
                tmSaf.lLevelToPrice(6) = 0
                tmSaf.lLevelToPrice(7) = 0
                tmSaf.lLevelToPrice(8) = 0
                tmSaf.lLevelToPrice(9) = 0
                tmSaf.lLevelToPrice(10) = 0
                'tmSaf.lLevelToPrice(11) = 0
                tmSaf.lHighPrice = 0
                tmSaf.iPreferredPct = 0
                tmSaf.iWk1stSoloIndex = 0
                tmSaf.sInvISCIForm = "R"
                tmSaf.lGGEMailSentDate = 0
                tmSaf.iRetainTrafProj = 24
                tmSaf.iRetainCntr = 60
                tmSaf.iRetainPayRevHist = 60
                tmSaf.iNoDaysRetainUAF = 5
                gPackDate "1/1/1990", tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1)
                gPackDate "1/1/1990", tmSaf.iEarliestAffSpot(0), tmSaf.iEarliestAffSpot(1)
                gPackDate "1/1/1990", tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1)
                gPackDate "1/1/1970", tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1)
                tmSaf.sFinalLogDisplay = "N"
                '6-9-14
                tmSaf.sProdProtMan = "N"
                tmSaf.sAvailGreenBar = "N"
                tmSaf.sIPumpZone = ""
                gPackDate "1/1/1990", tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
                '6-9-14
                tmSaf.sInvoiceSort = "P"
                tmSaf.sGenAutoFileWOSpt = "N"
                tmSaf.sFeatures1 = Chr(0)
                tmSaf.sFeatures2 = Chr(0)
                tmSaf.sEventTitle1 = ""
                tmSaf.sEventTitle2 = ""
                tmSaf.sEventSubtotal1 = ""
                tmSaf.sEventSubtotal2 = ""
                tmSaf.sCreditLimitMsg = "W"
                tmSaf.sXMidSpotsBill = "O"
                tmSaf.sHideDemoOnBR = "N"
                tmSaf.sClientSentToEDS = "N"
                tmSaf.sXDSHeadEndZone = "E"
                tmSaf.sSyncCopyInRot = "N"
                tmSaf.sFeatures3 = Chr(0)
                tmSaf.lStatementComment = 0
                tmSaf.sAudByPackage = "N"
                tmSaf.sFeatures4 = Chr(0)
                tmSaf.sFeatures5 = Chr(0)
                tmSaf.sFeatures6 = Chr(0)
                tmSaf.lCitationComment = 0
                tmSaf.sExcludeAudioTypeR = "N"       ' Exclude Autio Type Recorded
                                                         ' Commercial (RC: R). (Y/N, Test
                                                         ' for Y. Treat Blank same as N)
                tmSaf.sExcludeAudioTypeL = "N"         ' Exclude Autio Type Live
                                                         ' Commercial (LC: L). (Y/N, Test
                                                         ' for Y. Treat Blank same as N)
                tmSaf.sExcludeAudioTypeM = "N"         ' Exclude Autio Type Live Promo
                                                         ' (LP: M). (Y/N, Test for Y. Treat
                                                         ' Blank same as N)
                tmSaf.sExcludeAudioTypeS = "N"         ' Exclude Autio Type Recorded Promo
                                                         ' (RP: S). (Y/N, Test for Y. Treat
                                                         ' Blank same as N)
                tmSaf.sExcludeAudioTypeP = "N"         ' Exclude Autio Type Pre-Recorded
                                                         ' Live Commercial (PC: P). (Y/N,
                                                         ' Test for Y. Treat Blank same as
                                                         ' N)
                tmSaf.sExcludeAudioTypeQ = "N"        ' Exclude Autio Type Pre-Recorded
                tmSaf.sInvExpDelimiter = ","            'Invoice export delimiter
                tmSaf.sFeatures7 = Chr(0)
                tmSaf.sUnused = ""
                tmSaf.iCode = 0  'Autoincrement
                ilRet = btrInsert(hmSaf, tmSaf, imSafRecLen, INDEXKEY0)
            Else
                ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            End If
            If ilRet = BTRV_ERR_NONE Then
                bgHideHiddenLines = False
                If (Asc(tmSaf.sFeatures2) And HIDEHIDDENLINES) = HIDEHIDDENLINES Then
                    bgHideHiddenLines = True
                End If
                'D.S. 3/21/18
                If (Asc(tmSaf.sFeatures5) And PROGRAMMATICALLOWED) <> PROGRAMMATICALLOWED Then
                    mnuMonitor(17).Visible = False
                End If
                If ilSetTest = 0 Then
                    '5/2/13: Force GG on as if missing
                    ilSetTest = 2
                    ilField1 = Asc(tmSaf.sName)
                    slStr = Mid$(tmSaf.sName, 2, 5)
                    llDate = Val(slStr)
                    llNow = gDateValue(Format$(gNow(), "m/d/yy"))
                    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
                    If (ilField1 = 0) And (ilField2 = 1) Then
                        ilSetTest = 0
                        If llDate <= llNow Then
                            ilField2 = 0
                            Mid$(tmSaf.sName, 11, 1) = Chr(0)
                            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                        Else
                            igGGFlag = llDate - llNow   'llNow - llDate
                            If (tmSaf.lGGEMailSentDate <> llNow) And (igGGFlag <= 10) Then
                                tmSaf.lGGEMailSentDate = llNow
                                ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                                'Dan M 9/19/09 no smtp info, can't send yet
                                blSendGGWarning = True
                                'mSendGGWarningMsg
                            End If
                        End If
                    End If
                    If (ilField1 = 32) And (ilField2 = 32) Then
                        'ilSetTest = 0
                        'slStr = Chr(1)
                        ''slStr = slStr & "99999" & "0000"
                        'slStr = slStr & "9"
                        'For ilLoop = 0 To 3 Step 1
                        '    slStr = slStr & Trim$(Str$(Int((9 - 1 + 1) * Rnd + 1)))
                        'Next ilLoop
                        'slStr = slStr & "0000"
                        'slStr = slStr & Chr(1)
                        'slStr = slStr & "9"
                        'For ilLoop = 0 To 3 Step 1
                        '    slStr = slStr & Trim$(Str$(Int((9 - 1 + 1) * Rnd + 1)))
                        'Next ilLoop
                        'tmSaf.sName = slStr
                        'ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                    End If
                    If (ilField1 = 0) And (ilField2 = 0) Then
                        ilSetTest = 0
                        For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
                            igWinStatus(ilLoop) = 0
                        Next ilLoop
                        tgSpf.sAExcl = "N"
                        tgSpf.sAEDIC = "N"
                        tgSpf.sAEDII = "N"
                        tgSpf.sUseCartNo = "N"
                        tgSpf.sGUsePropSys = "N"
                        tgSpf.sGUseAffFeed = "N"
                        igGGFlag = 0
                    End If
                '5/2/13: Force GG on as if missing exe or d command specified at sign-on
                'Else
                End If
                '5/2/13: Don't set GG code for Demo database
                If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Or (igDemo) Then
                    ilSetTest = 0
                End If
                If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                    ilSetTest = 0
                End If
                If bgInternalGuide Then
                    ilField1 = Asc(tmSaf.sName)
                    slStr = Mid$(tmSaf.sName, 2, 5)
                    llDate = Val(slStr)
                    llNow = gDateValue(Format$(Now, "m/d/yy"))
                    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
                    If (ilField1 = 0) And (ilField2 = 1) Then
                        If llDate <= llNow Then
                            ilField2 = 0
                        End If
                    End If
                    If (ilField1 = 0) And (ilField2 = 0) Then
                        MsgBox "Database is client restricted, CSI Ok!", vbInformation, "Notification"
                    End If
                End If
                If ilSetTest <> 0 Then
                    slStr = Chr(0)
                    If igGGBDSignOn Then
                        llNow = gDateValue(Format$(gNow(), "m/d/yy")) + igGGDBSignOnDays
                    ElseIf ilSetTest = 1 Then
                        llNow = gDateValue(Format$(gNow(), "m/d/yy")) + Int(31 * Rnd + 60)
                    Else
                        '5/2/13: Force GG on.  Use Jan 1 of next year
                        llNow = gDateValue(gObtainYearEndDate(1, Format$(gNow(), "m/d/yy"))) + 1
                    End If
                    slNow = Trim$(str$(llNow))
                    Do While Len(slNow) < 5
                        slNow = "0" & slNow
                    Loop
                    slStr = slStr & slNow & "0000"
                    slStr = slStr & Chr(1)
                    llNow = gDateValue(Format$(gNow(), "m/d/yy"))
                    slNow = Trim$(str$(llNow))
                    Do While Len(slNow) < 5
                        slNow = "0" & slNow
                    Loop
                    tmSaf.sName = slStr & slNow
                    ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                End If
            End If
        Else
            MsgBox "Call Counterpoint as DDFReorg needs to be run", vbOKOnly + vbCritical, "Error"
            mTrafficStartUp = False
            btrDestroy hmSaf
            Exit Function
        End If
        gSetRptGGFlag tmSaf
        btrDestroy hmSaf
    Else
        MsgBox "Call Counterpoint as DDFReorg probably needs to be run, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
        mTrafficStartUp = False
        btrDestroy hmSaf
        Exit Function
    End If

    'hmSite = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmSite, "", sgDBPath & "Site.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet = BTRV_ERR_NONE Then
        imSiteRecLen = Len(tgSite)  'btrRecordLength(hlUrf)  'Get and save record length
        If imSiteRecLen = btrRecordLength(hgSite) Then
            llNoRec = btrRecords(hgSite) 'Obtain number of records
            If llNoRec = 0 Then
                'SQLQuery = "Insert Into Site (siteCode, siteMarketron, siteWeb, siteNoDaysDelq, siteCMCmtCode, siteWMCmtCode, sitePMCmtCode, siteOMCmtCode, siteOMNoWeeks, siteDaysRetainSpots, siteAdminArttCode, siteDaysResendISCI, siteVehicleStn, siteISCIExport, siteShowVehType, siteNoMnthRetain, siteLastDateArch, siteDayRetainPost)
                '                       Values (1,      0,              0,        0,             0,             0,             0,             0,             0,             0,                   0,                 0,                  'N',            'B',            'N',             999,              '1970-01-01',     0)"
                tgSite.lCode = 0
                tgSite.iMarketron = 0
                tgSite.iWeb = 0
                tgSite.iNoDaysDelq = 0
                'tgSite.iDaysResendISCI = 0
                tgSite.iUnused3 = 0
                tgSite.iCMCmtCode = 0
                tgSite.iWMCmtCode = 0
                tgSite.iPMCmtCode = 0
                tgSite.iOMCmtCode = 0
                tgSite.iOMNoWeeks = 0
                tgSite.iDaysRetainSpots = 0
                tgSite.lAdminArttCode = 0
                tgSite.sVehicleStn = "N"
                'tgSite.sISCIExport = "B"
                ''tgSite.iNoMnthRetain = 999
                ''gPackDate "1/1/1970", tgSite.iLastDateArch(0), tgSite.iLastDateArch(1)
                tgSite.sUsingServAgree = "N"
                tgSite.sSyncMulticast = "N"
                tgSite.iBandCmtCode = 0
                tgSite.sUnused2 = ""
                tgSite.sShowVehType = "N"
                tgSite.sExportCart4 = ""    ' Use 4 character Cart Numbers
                tgSite.sExportCart5 = ""    ' Use 4 character Cart Numbers
                tgSite.sCCEMail = ""
                tgSite.iDayRetainPost = 0       ' Number of days to retain web
                                                         ' posted spots
                tgSite.sChngPswd = ""    ' Flag to allow web users to change
                                                         ' passwords or not - Y/N
                gPackDate "1/1/1970", tgSite.iOMMinDate(0), tgSite.iOMMinDate(1)
                tgSite.sEmailHost = ""   ' Host name/SMTP server address or
                                                         ' URL
                tgSite.iEmailPort = 25       ' Port Number for SMTP server
                tgSite.sEmailAcctName = ""  ' Account name for SMTP
                                                         ' credentials.
                tgSite.sEmailPassword = ""   ' Password for the SMTP credentials
                tgSite.sEmailFromAddress = ""    ' The email address of the sender.
                tgSite.sEmailFromName = ""  ' Used as from email name.  Like
                                                         ' "Dial Global Netowrk" or "ABC
                                                         ' Clearance Dept."
                tgSite.iNCRWks = 60
                tgSite.sWebSuppressLog = "N"
                tgSite.sMultiVehWebPost = "N"
                tgSite.iWebNoDyKeepMiss = 60
                tgSite.sAllowBonusSpots = "N"
                tgSite.sShowContrDate = "N"
                tgSite.sDDF092710 = "Y"
                tgSite.sUsingViero = "N"
                tgSite.sUsingStationID = "N"
                tgSite.sStationToXDS = "N"
                tgSite.sAgreementToXDS = "N"
                tgSite.sMissedDateTime = "N"
                tgSite.sCompliantBy = "P"
                tgSite.iWebNoDyViewPost = 7
                gPackDate "1/1/1970", tgSite.iRqtDate(0), tgSite.iRqtDate(1)
                gPackTime "12AM", tgSite.iRqtTime(0), tgSite.iRqtTime(1)
                tgSite.sGenTransparent = "N"
                tgSite.sProgToXDS = "N"
                gPackDate "1/1/1970", tgSite.iSSBDate(0), tgSite.iSSBDate(1)
                gPackTime "12AM", tgSite.iSSBTime(0), tgSite.iSSBTime(1)
                tgSite.sSupportXDSDelay = "N"
                tgSite.sAllowAutoPost = "N"
                tgSite.sWithinMissMonth = "N"       ' Within missed standard month (Y=Yes; N or Blank=No). Test for Y
                tgSite.sLastWk1stWk = "N"           ' Miised in last week of the standard month, MG allowed in 1st
                                                    ' week of next month (Y=Yes; N or Blank = No). Test for Y
                tgSite.sSkipHiatusWk = "N"          ' Disallow MG in skipped Weeks (Hiatus). (Y=Yes; N or Blank=No). Test for Y
                tgSite.sValidDaysOnly = "N"         ' Book only into valid flight days. (Y=Yes; N or Blank=No). Test for Y
                tgSite.sTimeRange = "N"             ' MG with time range (O=Flight Order Times; P=Pledge Times; N or
                                                    ' Blank=Any Time).Test for O or P
                tgSite.sISCIPolicy = "O"            ' ISCI Policy: O or Blank=Order; A=Advertiser after Order failed. Test for A
                tgSite.iMGDays = -1
                tgSite.iCompetSepTime = 0
                tgSite.sAllowMGSpots = "Y"
                tgSite.sAllowReplSpots = "Y"
                tgSite.sNoMissedReason = "N"
                tgSite.sDefaultEstDay = "N"
                tgSite.sWebPostInFuture = "Y"
                tgSite.sMissedMGBypass = "N"
                'tgSite.sUnused = ""
                tgSite.sRADARMultiAir = "S"
                'added by Dan 9/22/20
                tgSite.lAstMaxLastValue = 0
                gPackDate "1/1/1970", tgSite.iAstMaxDate(0), tgSite.iAstMaxDate(1)
                '9926
                tgSite.iUMCmtCode = 0
                ilRet = btrInsert(hgSite, tgSite, imSiteRecLen, INDEXKEY0)
            Else
                ilRet = btrGetFirst(hgSite, tgSite, imSiteRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If (tgSpf.sGUseAffSys = "Y") And ((Trim$(tgSite.sDDF092710) = "") Or (Trim$(tgSite.sDDF092710) = "N")) Then
                    MsgBox "Run Affiliate System to complete DDF Install", vbInformation + vbOKOnly, "Information"
                ElseIf (Trim$(tgSite.sDDF092710) = "") Or (Trim$(tgSite.sDDF092710) = "N") Then
                    'Set ULF fields
                    imUlfRecLen = Len(tmUlf)
                    ilRet = btrGetFirst(hgUlf, tmUlf, imUlfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    Do While ilRet = BTRV_ERR_NONE
                        tmUlf.iTrafJobNo = -1
                        tmUlf.iAffTaskNo = -1
                        ilRet = btrUpdate(hgUlf, tmUlf, imUlfRecLen)
                        ilRet = btrGetNext(hgUlf, tmUlf, imUlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop
                    'Clear Flag in site
                    tgSite.sDDF092710 = "Y"
                    ilRet = btrUpdate(hgSite, tgSite, imSiteRecLen)
                End If
                'Dan 12/11/14 Dick said to not send if signed in as csi, so I changed call to new global
'                If blSendGGWarning Then 'shutdown not far off, tell service people
'                    mSendGGWarningMsg
'                End If
                If blSendGGWarning And Not bgInternalGuide Then 'shutdown not far off, tell service people
                    slTemp = Trim$(tgSpf.sGClient) & sgCR & sgLF & Trim$(tgSpf.sGAddr(0)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(1)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(2))
                    slTemp = slTemp & sgCR & sgLF & "The above client system will shutdown in " & igGGFlag & " days"
                    slTemp = slTemp & sgCR & sgLF & "Please extend the shutdown date to a new date which you can obtain from the Counterpoint service manager"
                    gSendServiceEmail "Client System Shutdown approaching", slTemp
                End If
            End If
        Else
            MsgBox "Call Counterpoint as DDFReorg needs to be run", vbOKOnly + vbCritical, "Error"
            mTrafficStartUp = False
           ' btrDestroy hmSite
            Exit Function
        End If
        'btrDestroy hmSite
    'Else
    '    MsgBox "Call Counterpoint as DDFReorg probably needs to be run, error # " & str$(ilRet), vbOkOnly + vbCritical, "Error"
    '    mTrafficStartUp = False
    '    btrDestroy hmSite
    '    Exit Function
    'End If

    If Not mCheckParticipant() Then
        mTrafficStartUp = False
        btrDestroy hmSaf
        Exit Function
    End If
    If Not mCheckCopyUsage() Then
        mTrafficStartUp = False
        btrDestroy hmSaf
        Exit Function
    End If


    For ilLoop = 0 To 12 Step 1
        cmcTask(ilLoop).Visible = True
    Next ilLoop
    For ilLoop = LBound(igJobShowing) To UBound(igJobShowing) Step 1
        igJobShowing(ilLoop) = False
    Next ilLoop
    cmcReport.Visible = True
    'plcDate.Visible = True
    If igGGFlag = 0 Then
        For ilLoop = 0 To 12 Step 1
            cmcTask(ilLoop).Enabled = False
            cmcTask(ilLoop).FontBold = False
            'cmcTask(ilLoop).BackColor = RED
        Next ilLoop
        mnuAccessorySetCredit.Enabled = False
        mnuAccessoryShowCredits.Enabled = False
    Else
        gSetWinStatusFromSPF igWinStatus()
        'If tgSpf.sGUsePropSys = "Y" Then
            If igWinStatus(BUDGETSJOB) = 0 Then
                cmcTask(0).Enabled = False
                cmcTask(0).FontBold = False
                'cmcTask(0).BackColor = vbRed
            Else
                If igWinStatus(BUDGETSJOB) = 1 Then
                    'cmcTask(0).BackColor = vbYellow
                    'cmcTask(0).ForeColor = vbYellow
                    cmcTask(0).FontBold = False
                Else
                    'cmcTask(0).BackColor = vbButtonFace
                End If
            End If
        'Else
        '    cmcTask(0).Enabled = False
        '    'cmcTask(0).BackColor = vbRed
        'End If
        If igWinStatus(RATECARDSJOB) = 0 Then
            cmcTask(1).Enabled = False
            cmcTask(1).FontBold = False
            'cmcTask(1).BackColor = vbRed
        Else
            If igWinStatus(RATECARDSJOB) = 1 Then
                'cmcTask(1).BackColor = vbYellow
                'cmcTask(1).ForeColor = vbYellow
                cmcTask(1).FontBold = False
            Else
                'cmcTask(1).BackColor = vbButtonFace
            End If
        End If
        'If tgSpf.sGUsePropSys = "Y" Then
            If igWinStatus(PROPOSALSJOB) = 0 Then
                cmcTask(2).Enabled = False
                cmcTask(2).FontBold = False
                'cmcTask(2).BackColor = vbRed
            Else
                If igWinStatus(PROPOSALSJOB) = 1 Then
                    'cmcTask(2).BackColor = vbYellow
                    'cmcTask(2).ForeColor = vbYellow
                    cmcTask(2).FontBold = False
                Else
                    'cmcTask(2).BackColor = vbButtonFace
                End If
            End If
        'Else
        '    cmcTask(2).Enabled = False
        '    'cmcTask(2).BackColor = vbRed
        'End If
        If igWinStatus(CONTRACTSJOB) = 0 Then
            cmcTask(3).Enabled = False
            cmcTask(3).FontBold = False
            'cmcTask(3).BackColor = vbRed
        Else
            If igWinStatus(CONTRACTSJOB) = 1 Then
                'cmcTask(3).BackColor = vbYellow
                'cmcTask(3).ForeColor = vbYellow
                cmcTask(3).FontBold = False
            Else
                'cmcTask(3).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(PROGRAMMINGJOB) = 0 Then
            cmcTask(4).Enabled = False
            cmcTask(4).FontBold = False
            'cmcTask(4).BackColor = vbRed
        Else
            If igWinStatus(PROGRAMMINGJOB) = 1 Then
                'cmcTask(4).BackColor = vbYellow
                'cmcTask(4).ForeColor = vbYellow
                cmcTask(4).FontBold = False
            Else
                'cmcTask(4).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(SPOTSJOB) = 0 Then
            cmcTask(5).Enabled = False
            cmcTask(5).FontBold = False
            'cmcTask(5).BackColor = vbRed
        Else
            If igWinStatus(SPOTSJOB) = 1 Then
                'cmcTask(5).BackColor = vbYellow
                'cmcTask(5).ForeColor = vbYellow
                cmcTask(5).FontBold = False
            Else
                'cmcTask(5).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(COPYJOB) = 0 Then
            cmcTask(6).Enabled = False
            cmcTask(6).FontBold = False
            'cmcTask(6).BackColor = vbRed
        Else
            If igWinStatus(COPYJOB) = 1 Then
                'cmcTask(6).BackColor = vbYellow
                'cmcTask(6).ForeColor = vbYellow
                cmcTask(6).FontBold = False
            Else
                'cmcTask(6).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(LOGSJOB) = 0 Then
            cmcTask(7).Enabled = False
            cmcTask(7).FontBold = False
            'cmcTask(7).BackColor = vbRed
        Else
            If igWinStatus(LOGSJOB) = 1 Then
                'cmcTask(7).BackColor = vbYellow
                'cmcTask(7).ForeColor = vbYellow
                cmcTask(7).FontBold = False
            Else
                ''cmcTask(7).BackColor = vbButtonFace
            End If
        End If
        cmcTask(8).Left = cmcTask(7).Left + cmcTask(7).Width
        cmcTask(9).Left = cmcTask(8).Left + cmcTask(8).Width
        cmcTask(10).Left = cmcTask(9).Left + cmcTask(9).Width
        cmcTask(11).Left = cmcTask(10).Left + cmcTask(10).Width
        cmcTask(12).Left = cmcTask(11).Left + cmcTask(11).Width
        cmcReport.Left = cmcTask(12).Left + cmcTask(12).Width
        cmcReport.Width = lgCurrHRes * 15 - cmcReport.Left
        If tgSpf.sSystemType <> "R" Then
            If tgSpf.sGUseAffFeed = "Y" Then
'                If igWinStatus(STATIONFEEDJOB) = 0 Then
'                    cmcTask(8).Enabled = False
'                    'cmcTask(8).BackColor = vbRed
'                Else
'                    If igWinStatus(STATIONFEEDJOB) = 1 Then
'                        'cmcTask(8).BackColor = vbYellow
'                        'cmcTask(8).ForeColor = vbYellow
'                        cmcTask(8).FontBold = False
'                    Else
'                        ''cmcTask(8).BackColor = vbButtonFace
'                    End If
'                End If
                mnuExportStnFeed.Enabled = True
            Else
'                cmcTask(8).Enabled = False
'                'cmcTask(8).BackColor = vbRed
                mnuExportStnFeed.Enabled = False
            End If
            cmcTask(8).Enabled = False
            cmcTask(8).Visible = False
            cmcTask(9).Left = cmcTask(9).Left - cmcTask(8).Width
            cmcTask(10).Left = cmcTask(10).Left - cmcTask(8).Width
            cmcTask(11).Left = cmcTask(11).Left - cmcTask(8).Width
            cmcTask(12).Left = cmcTask(12).Left - cmcTask(8).Width
            cmcReport.Left = cmcReport.Left - cmcTask(8).Width
            cmcReport.Width = cmcReport.Width + cmcTask(8).Width
        Else
            mnuExportStnFeed.Enabled = False
            cmcTask(8).Enabled = False
            cmcTask(8).Visible = False
            cmcTask(9).Left = cmcTask(9).Left - cmcTask(8).Width
            cmcTask(10).Left = cmcTask(10).Left - cmcTask(8).Width
            cmcTask(11).Left = cmcTask(11).Left - cmcTask(8).Width
            cmcTask(12).Left = cmcTask(12).Left - cmcTask(8).Width
            cmcReport.Left = cmcReport.Left - cmcTask(8).Width
            cmcReport.Width = cmcReport.Width + cmcTask(8).Width
        End If
        If igWinStatus(POSTLOGSJOB) = 0 Then
            cmcTask(9).Enabled = False
            cmcTask(9).FontBold = False
            'cmcTask(9).BackColor = vbRed
        Else
            If igWinStatus(POSTLOGSJOB) = 1 Then
                'cmcTask(9).BackColor = vbYellow
                'cmcTask(9).ForeColor = vbYellow
                cmcTask(9).FontBold = False
            Else
                ''cmcTask(9).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(INVOICESJOB) = 0 Then
            cmcTask(10).Enabled = False
            cmcTask(10).FontBold = False
            'cmcTask(10).BackColor = vbRed
        Else
            If igWinStatus(INVOICESJOB) = 1 Then
                'cmcTask(10).BackColor = vbYellow
                'cmcTask(10).ForeColor = vbYellow
                cmcTask(10).FontBold = False
            Else
                ''cmcTask(10).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(COLLECTIONSJOB) = 0 Then
            cmcTask(11).Enabled = False
            cmcTask(11).FontBold = False
            'cmcTask(11).BackColor = vbRed
        Else
            If igWinStatus(COLLECTIONSJOB) = 1 Then
                'cmcTask(11).BackColor = vbYellow
                'cmcTask(11).ForeColor = vbYellow
                cmcTask(11).FontBold = False
            Else
                ''cmcTask(11).BackColor = vbButtonFace
            End If
        End If
        If igWinStatus(SLSPCOMMSJOB) = 0 Then
            cmcTask(12).Enabled = False
            cmcTask(12).FontBold = False
            'cmcTask(12).BackColor = vbRed
        Else
            If igWinStatus(SLSPCOMMSJOB) = 1 Then
                'cmcTask(12).BackColor = vbYellow
                'cmcTask(12).ForeColor = vbYellow
                cmcTask(12).FontBold = False
            Else
                ''cmcTask(12).BackColor = vbButtonFace
            End If
        End If
        '11/16/09:  Make find active for all users
        'If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
            'cmcFind.Left = cmcReport.Left
            'cmcFind.Top = cmcReport.Top
            '''Moved to after tgSaf is populate as it is erase above
            '''cmcReport.Move cmcFind.Left + cmcFind.Width + 15, cmcReport.Top, pbcMsgArea.Width - cmcFind.Left - cmcFind.Width + 15
            ''If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then
            ''    cmcEMail.Visible = True
            ''    cmcEMail.Move pbcMsgArea.Width - cmcEMail.Width + 15, cmcReport.Top
            ''    cmcReport.Move cmcFind.Left + cmcFind.Width, cmcReport.Top, cmcEMail.Left - (cmcFind.Left + cmcFind.Width)
            ''Else
            ''    cmcEMail.Visible = False
            ''    cmcReport.Move cmcFind.Left + cmcFind.Width + 15, cmcReport.Top, pbcMsgArea.Width - cmcFind.Left - cmcFind.Width + 15
            ''End If
            'cmcFind.Visible = True
        'Else
        '    cmcFind.Visible = False
        'End If
        
        If (igWinStatus(COLLECTIONSJOB) = 2) Then
        Else
            mnuAccessorySetCredit.Enabled = False
            mnuAccessoryShowCredits.Enabled = False
        End If

    End If
    '12/23/15: Move from above
    cmcFind.Left = cmcReport.Left
    cmcFind.Top = cmcReport.Top
    cmcFind.Visible = True
    blShowAffiliate = False
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        blShowAffiliate = True
    Else
        slSQLQuery = "Select ustName, ustPassword From Ust Where "
        slSQLQuery = slSQLQuery & " ustState <> 1 "
        slSQLQuery = slSQLQuery & " and ustName = '" & Trim$(tgUrf(0).sName) & "'"
        slSQLQuery = slSQLQuery & " and ustPassword = '" & Trim$(tgUrf(0).sPassword) & "'"
        Set rst_Ust = gSQLSelectCall(slSQLQuery)
        If Not rst_Ust.EOF Then
            blShowAffiliate = True
        End If
        rst_Ust.Close
    End If
    If blShowAffiliate Then
        cmcAffiliate.Move cmcFind.Left + cmcFind.Width, cmcReport.Top, 3 * cmcAffiliate.Width
        cmcAffiliate.Visible = True
    Else
        cmcAffiliate.Visible = False
    End If
    ' Dan M 4/13/09 limit user's menu items  --set counterpoint date
    ' Dan M 1/11/10 whoops.  No need to limit user seeing counterpoint date
   ' mLimitUserMenu
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        If btrRecordLength(hmSsf) <> 20 Then
            MsgBox "Call Counterpoint as SSFVB3_6 needs to be run", vbOKOnly + vbCritical, "Error"
            mTrafficStartUp = False
            btrDestroy hmSaf
            Exit Function
        End If
        btrDestroy hmSsf
    Else
        MsgBox "Call Counterpoint as SSF missing, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
        mTrafficStartUp = False
        btrDestroy hmSsf
        Exit Function
    End If
    'aleways populate- 3/17/03
    'we might have to re install if it acts to slow
    'If Not igUsingTimeStamp Then
        mPopulateListBoxes
    'Else
    '    gAnyClustersDef
    '    gAnyRepDef
    'End If

    If Not mCheckBuyer() Then
        mTrafficStartUp = False
        btrDestroy hmSaf
        Exit Function
    End If

    If Not mCheckSeason() Then
        mTrafficStartUp = False
        Exit Function
    End If

    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then
        cmcEMail.Visible = True
        cmcEMail.Move pbcMsgArea.Width - cmcEMail.Width + 15, cmcReport.Top
        If cmcAffiliate.Visible Then
            cmcReport.Move cmcAffiliate.Left + cmcAffiliate.Width, cmcReport.Top, cmcEMail.Left - (cmcAffiliate.Left + cmcAffiliate.Width)
        Else
            cmcReport.Move cmcFind.Left + cmcFind.Width, cmcReport.Top, cmcEMail.Left - (cmcFind.Left + cmcFind.Width)
        End If
    Else
        cmcEMail.Visible = False
        If cmcAffiliate.Visible Then
            cmcReport.Move cmcAffiliate.Left + cmcAffiliate.Width + 15, cmcReport.Top, pbcMsgArea.Width - cmcAffiliate.Left - cmcAffiliate.Width + 15
        Else
            cmcReport.Move cmcFind.Left + cmcFind.Width + 15, cmcReport.Top, pbcMsgArea.Width - cmcFind.Left - cmcFind.Width + 15
        End If
    End If
    
    If tgSaf(0).sHideDemoOnBR <> "Y" Then
        mnuImportImpressions.Enabled = False
    End If


    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    ' 5676: don't assume c drive
    slRootPath = sgRootDrive & "CSI\"
    slPhotoPath = slRootPath & "RptLogo.jpg"
    slBitmapPath = slRootPath & "RptLogo.Bmp"
    slLogoPath = sgLogoPath & "RptLogo.Bmp"
    ilRet = 0
    On Error GoTo mTrafficStartUpErr:
    slDateTime1 = FileDateTime(sgLogoPath & "rptlogo.jpg")
    bgRptLogoJpgExists = True
    sgRptLogoExt = "jpg"
    If ilRet <> 0 Then      'jpg does not exist, assume .bmp
        bgRptLogoJpgExists = False
        sgRptLogoExt = "bmp"
    Else                    'jpg exists, see if .bmp exists.  If so, rename so it doesnt conflict when renaming for customized logs
        slDateTime1 = FileDateTime(sgLogoPath & "rptlogo.bmp")
        If ilRet = 0 Then           'rptlogo.bmp exists, rename it
            Name sgLogoPath & "rptlogo.bmp" As Trim$(sgLogoPath) & "rptlogo1.bmp"
        End If

    End If
    On Error GoTo 0
    
'8-18-14 do not copy the logos to Root drive
'    slDateTime1 = FileDateTime(slBitmapPath)
'    If ilRet <> 0 Then
'        ilRet = 0
'        MkDir slRootPath
'        If ilRet = 0 Then
'            FileCopy slLogoPath, slBitmapPath
'        Else
'            FileCopy slLogoPath, slBitmapPath
'        End If
'    Else
'        ilRet = 0
'        slDateTime2 = FileDateTime(slLogoPath)
'        If ilRet = 0 Then
'            If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
'                FileCopy slLogoPath, slBitmapPath
'            End If
'        End If
'    End If
'    On Error GoTo 0


        'ttp 5260
    '8-18-14 no longer copy logos to root drive
'    slLogoPath = sgLogoPath & "RptLogo.jpg"
'    If Dir(slLogoPath) > "" Then
'        If Dir(slPhotoPath) = "" Then
'            FileCopy slLogoPath, slPhotoPath
'        'ok, both exist.  is logopath's more recent?
'        Else
'            '5461
'            slDateTime1 = FileDateTime(slLogoPath)
'            slDateTime2 = FileDateTime(slPhotoPath)
'            If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
'                FileCopy slLogoPath, slPhotoPath
'            End If
'        End If
'    End If
    
    
    
'    slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
'    If ilRet <> 0 Then
'        ilRet = 0
'        MkDir "C:\CSI"
'        If ilRet = 0 Then
'            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
'        Else
'            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
'        End If
'    Else
'        ilRet = 0
'        slDateTime2 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
'        If ilRet = 0 Then
'            If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
'                FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
'            End If
'        End If
'    End If
'    On Error GoTo 0
'        'ttp 5260
'    If Dir(sgLogoPath & "RptLogo.jpg") > "" Then
'        If Dir("c:\csi\RptLogo.jpg") = "" Then
'            FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
'        'ok, both exist.  is logopath's more recent?
'        Else
'            '5461
'            slDateTime1 = FileDateTime(sgLogoPath & "RptLogo.jpg")
'            slDateTime2 = FileDateTime("C:\CSI\RptLogo.jpg")
'            If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
'                FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
'            End If
'        End If
'    End If
    mCheckDiskFreeSpace
    mCheckShortDateForm
    gCheckForContFiles
    If (Not igShowPicture) Or (igRemoteUser) Or (igTestSystem) Then
        If igSportsSystem <> 0 Then
            Sports.Show vbModeless
            pbcMsgArea.Visible = False
            mDisableMenu
        Else
            Basic10.Show vbModeless
        End If
    End If
    igIsSignedOn = True  'D.S. 3/10/03
    'If the user is signed onto the Test system don't let them backup data to test
    If igTestSystem = True Then
        mnuTrafficDataToTest.Enabled = False
    End If

    'Do we show the menu item to backup the Affiliate system
    If tgSpf.sGUseAffSys = "Y" Then
        'mnuBackupAffiliate.Enabled = True
        mnuBackupTrafficAndAffiliate.Enabled = True
    Else
        'mnuBackupAffiliate.Enabled = False
        mnuBackupTrafficAndAffiliate.Enabled = True 'False
    End If
    
    ilRet = gVffRead()
    mnuExportAirWave.Enabled = False
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).sExportAirWave = "Y" Then
            mnuExportAirWave.Enabled = True
            Exit For
        End If
    Next ilVff

    'For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
    '    If tgVff(ilVff).sPostLogSource = "S" Then
    '        mnuAccessoryAcquisitionCheck.Visible = True
    '        Exit For
    '    End If
    'Next ilVff

    'gCheckForContFiles

    mCheckV81Flags
    
    ilRet = gDeleteLockRec_ByUser(tgUrf(0).iCode)

    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    sgDateRangeStartDate = Format$(llNowDate - 365, "m/d/yy")
    sgDateRangeEndDate = Format$(llNowDate - 1, "m/d/yy")
    ' Dan M 4/09/09 Limit Guide
    mLimitGuideMenus
    '8297
'    Set ogPdf = New PDFSplitMerge.CPDFSplitMergeObj
'    ogPdf.SetCode "989888DFE6F0FF0056681A"
    
    '07-13-15
    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then
    '    mnuSendToEDS.Enabled = True
        bgEDSIsActive = True
    Else
    '    mnuSendToEDS.Enabled = False
        bgEDSIsActive = False
    End If
    lgShellAndWaitID = 0
    Load TrafficAffiliateCom
    mTrafficStartUp = True
    Exit Function

ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
mTrafficStartUpErr:
    ilRet = err.Number
    Resume Next
ERRNOLOCATIONNAME:
    mTrafficStartUp = False
    gMsgBox "Traffic.Ini [Locations] 'Name' is not valid. Traffic will now shutdown.", vbCritical
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnloadForms                    *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove any active form.  Test if*
'*                     the user stopped the unload     *
'*                                                     *
'*******************************************************
Private Sub mUnloadForms(ilSignOn As Integer)
'
'   mUnloadForms iSignon
'   Where:
'       iSignon (I)- True = Unload Signon, False = leave Signon loaded
'
'    'Unload RateCard
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set RateCard = Nothing   'Remove data segment
'    'Unload Collect
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set Collect = Nothing   'Remove data segment
'    'Unload Copy
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set Copy = Nothing   'Remove data segment
'    'Unload PostLog
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set PostLog = Nothing   'Remove data segment
'    'Unload Program
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set Program = Nothing   'Remove data segment
'    Unload IconTraf
'    Set IconTraf = Nothing
'    'Unload Spots
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set Spots = Nothing   'Remove data segment
'    Unload CAvail
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Set CAvail = Nothing   'Remove data segment
'    'Unload CBill
'    'If igStopCancel Then    'User stop the unload
'    '    Exit Sub
'    'End If
'    'Set CBill = Nothing   'Remove data segment
'    Unload Contract
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Unload CCancel
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Set CCancel = Nothing   'Remove data segment
'    Unload CPackage
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Set CPackage = Nothing   'Remove data segment
'    Unload CRevNo
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Set CRevNo = Nothing   'Remove data segment
'    Unload CShift
'    If igStopCancel Then    'User stop the unload
'        Exit Sub
'    End If
'    Set CShift = Nothing   'Remove data segment
'    Set Contract = Nothing   'Remove data segment
'    'Traffic!cmcAuxWnd.Value = False
'    'Traffic!cmcBasicWnd.Value = False
''    Unload auxwnd
''    If igStopCancel Then    'User stop the unload
''        Exit Sub
''    End If
''    Unload Basic10
''    If igStopCancel Then    'User stop the unload
''        Exit Sub
''    End If
'    If ilSignOn Then        'User stop the unload
'        Unload MstPict
'        Set MstPict = Nothing   'Remove data segment
'        'Unload Signon
'        'Set Signon = Nothing   'Remove data segment
'    End If
    Dim ilLoop As Integer
    '2/7/09:  Loaded forms like Contract are unloaded automatically prior to unloading traffic
    For ilLoop = Forms.Count - 1 To 0 Step -1
        Unload Forms(ilLoop)
    Next ilLoop
End Sub

Private Sub mnuAccessoryUtilitiesSSFCheck_Click()
    'Dim slStr As String
    'If igTestSystem Then
    '    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'Else
    '    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    'End If
    'sgCommandStr = slStr
    'On Error Resume Next
    'SsfCheck.Show vbModal
'    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)
'        ChDir sgCurDir
'    End If
    'gChDrDir

End Sub



Private Sub pbcMsgArea_GotFocus()
'    'This control will get focus if no other form is active
'    'This control also get focus at various times- therefore the test on
'    'the number of forms running
'    If Forms.Count = igNoMinForms + igCalActive + igCalcActive Then 'Show basic 10 if only this form to showing
'        If Not igShowPicture Then
'            If igBasicOrAux = 2 Then
'                Traffic!cmcAuxWnd.Value = True
'            Else
'                Traffic!cmcBasicWnd.Value = True
'            End If
'        Else
'            If (igCalActive = 0) And (igCalcActive = 0) Then
'                'cmcJobs.Enabled = True
'                'cmcLists.Enabled = True
'                igShowPicture = False
'                'Traffic.plcHelp.Caption = " press JOBS or LISTS button to continue"
'                'Traffic!plcHelp.Cls
'                'Traffic!plcHelp.CurrentX = 0
'                'Traffic!plcHelp.CurrentY = 0
'                'Traffic!plcHelp.Print " press JOBS or LISTS button to continue"
'            End If
'        End If
'    End If
End Sub




Private Sub RichTextBox1_Change()
'rich text box 9-01-09 Dan M  to modify rich text and add double space for reports
End Sub

Private Sub tmcCheckAlert_Timer()
'    Dim ilRet As Integer
'
'    If tgSpf.iGAlertInterval <> 0 Then
'        If ((tgSpf.iGAlertInterval <> 0) And (tgSpf.iGAlertInterval <= igAlertTimer)) Then
'            'Check if Alert exist
'            imAufRecLen = Len(tmAuf)
'            If (tgUrf(0).sCompAlert <> "N") Then
'                tmAufSrchKey1.sType = "C"
'                tmAufSrchKey1.sStatus = "R"
'                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'                If ilRet = BTRV_ERR_NONE Then
'                    tmcCheckAlert.Enabled = False
'                    igAlertFlash = -1
'                    igAlertTimer = 0
'                    tmcFlashAlert.Interval = 2000   'Every 2 second
'                    tmcFlashAlert.Enabled = True
'                    DoEvents
'                    Exit Sub
'                End If
'            End If
'            If (tgUrf(0).sReprintLogAlert <> "N") Then
'                tmAufSrchKey1.sType = "L"
'                tmAufSrchKey1.sStatus = "R"
'                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'                If ilRet = BTRV_ERR_NONE Then
'                    tmcCheckAlert.Enabled = False
'                    igAlertFlash = -1
'                    igAlertTimer = 0
'                    tmcFlashAlert.Interval = 2000   'Every 2 second
'                    tmcFlashAlert.Enabled = True
'                    DoEvents
'                    Exit Sub
'                End If
'            End If
'            tmcFlashAlert.Enabled = False
'            tmcCheckAlert.Enabled = True
'            mnuAlert.Visible = False
'            igAlertTimer = 1
'        Else
'            igAlertTimer = igAlertTimer + 1
'        End If
'    Else
'        tmcCheckAlert.Enabled = False
'    End If
    Dim ilRet As Integer

    ilRet = gAlertCheck()
End Sub

Private Sub tmcClock_Timer()
    Dim slDateTime As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slMessage As String
    Dim llCountdownEnteredTime As Long
    Dim llCefCode As Long
    Dim llAufCode As Long
    Dim slName As String

    mCheckGG

    ilRet = gAlertCheckBlock(tmAuf)
    If ilRet = 1 Then   'Block user
        imPrevAlertBlockChk = ilRet
        MstPict.Cls
        gUnpackTimeLong tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1), False, llCountdownEnteredTime
        If (igShowVersionNo <> -1) Or (llCountdownEnteredTime <> lmCountdownEnteredTime) Then
            If tmAuf.lCefCode > 0 Then
                hmCef = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    tmCefSrchKey.lCode = tmAuf.lCefCode
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    slMessage = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'End If
                        slMessage = Trim$(tmCef.sComment)
                    End If
                Else
                    slMessage = "System will be shut down in " & Trim$(str$(tmAuf.iCountdown)) & " minutes"
                End If
                btrDestroy hmCef
            Else
                slMessage = "System will be shut down in " & Trim$(str$(tmAuf.iCountdown)) & " minutes"
            End If
            igShowVersionNo = -1
            sgWallpaper = "Shutdown : " & tmAuf.iCountdown & "m"
            gWriteBkgd
            imCountdown = tmAuf.iCountdown
            lmCountdownEnteredTime = llCountdownEnteredTime
            MsgBox slMessage, vbInformation, "Notification"
        Else
            imCountdown = imCountdown - 1
            If imCountdown >= 0 Then
                igShowVersionNo = -1
                sgWallpaper = "Shutdown : " & imCountdown & "m"
                gWriteBkgd
            Else
                igShowVersionNo = -1
                sgWallpaper = "Shutdown"
                gWriteBkgd
            End If
        End If
    ElseIf ilRet = 2 Then   'Block Initiated by this user
        imPrevAlertBlockChk = ilRet
        MstPict.Cls
        gUnpackTimeLong tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1), False, llCountdownEnteredTime
        If (igShowVersionNo <> -1) Or (llCountdownEnteredTime <> lmCountdownEnteredTime) Then
            igShowVersionNo = -1
            sgWallpaper = "Shutdown: " & tmAuf.iCountdown & "m"
            gWriteBkgd
            imCountdown = tmAuf.iCountdown
            lmCountdownEnteredTime = llCountdownEnteredTime
        Else
            imCountdown = imCountdown - 1
            If imCountdown >= 0 Then
                igShowVersionNo = -1
                sgWallpaper = "Shutdown: " & imCountdown & "m"
                gWriteBkgd
            Else
                igShowVersionNo = -1
                sgWallpaper = "Shutdown"
                gWriteBkgd
            End If
        End If
    Else
        If igShowVersionNo = -1 Then
            MstPict.Cls
            igShowVersionNo = imShowVersionNo
            sgWallpaper = smWallpaper
            gWriteBkgd
            If imPrevAlertBlockChk <> 2 Then
                MsgBox "System Shut Down has been stopped!", vbInformation, "Notification"
            End If
            imPrevAlertBlockChk = ilRet
        End If
    End If

    Do
        llAufCode = gAlertCheckNotice(lgUlfCode, llCefCode)
        If llAufCode > 0 Then
            If llCefCode > 0 Then
                hmCef = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    tmCefSrchKey.lCode = llCefCode
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slMessage = gStripChr0(tmCef.sComment)
                        'If tmCef.iStrLen > 0 Then
                        If slMessage <> "" Then
                            'slMessage = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                            MsgBox slMessage, vbInformation, "Notification"
                        End If
                    End If
                End If
                btrDestroy hmCef
            End If
            ilRet = gAlertClear("A", "N", "", 0, 0, Trim$(str$(llAufCode)))
        Else
            Exit Do
        End If
    Loop While llAufCode > 0

    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    'slStr = Traffic.Caption
    'ilPos = InStr(1, slStr, " **", vbTextCompare)
    'If ilPos > 0 Then
    '    slStr = Left$(slStr, ilPos - 1)
    '    Traffic.Caption = slStr & " ** " & slDateTime & " **"
    'End If

' JD 11-24-08 - CSI_Server is eating up handles when this is called every minute.
'               Turning this off will stop the out of resource problem.
'               This is a temporary change to buy time until we can fix CSI_Server.
'    If gIsBackupRunning() Then
'        mnuDate.Caption = slDateTime & " ** Backup Executing **"
'    Else
    If sgDateBrannerMsg = "" Then
        gShowPartBranner
    End If
        mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg  '"                       "
'    End If

    If (tgWebUser.iPrimaryCode <> 0 Or tgWebUser.iSecondaryCode <> 0) And tgSpf.sSystemType = "R" Then
        'If the record is locked then another user is polling so get out
        If gStartWebPoll(tgWebUser.iPrimaryCode) = False Then
            Exit Sub
        End If
        ilPos = gPollWebForSpots()
    End If
    mCheckActiveLog

'
'    slStr = Format$(gNow(), "m/d/yy")
'    slStr = " " & gAddDayToDate(slStr)
'    'plcDate.Caption = slStr
'    plcDate.Cls
'    plcDate.CurrentX = 0
'    plcDate.CurrentY = 0
'    plcDate.Print slStr
'    slStr = Format$(gNow(), " h:mmAM/PM")
'    ''plcTime.Caption = slStr
'    'plcTime.Cls
'    'plcTime.CurrentX = 0
'    'plcTime.CurrentY = 0
'    'plcTime.Print slStr
End Sub

Private Sub mShowAlert(hlBitmap As Long)
    Dim ilRet As Integer
    Dim llMenuID As Long

    ' Get a handle to the top level menu
    hmAlert = GetMenu(Traffic.HWnd)
    lmMenuID = GetMenuItemID(hmAlert, 3)

    'If hmBitmap = 0 Then hmBitmap = CopyPictureImage(IconTraf!pbcAlert)
    ' And replace it with a bitmap.
    ilRet = ModifyMenuBynum(hmAlert, 3, MF_BITMAP Or MF_BYPOSITION Or MF_ENABLED, llMenuID, hlBitmap)

End Sub

Private Sub tmcFlashAlert_Timer()
    igAlertFlash = igAlertFlash + 1
    If igAlertFlash = 10 Then
        If ((tgSpf.iGAlertInterval <> 0) And (tgSpf.iGAlertInterval <= igAlertTimer)) Then
            tmcFlashAlert.Enabled = False
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
            tmcCheckAlert.Enabled = True
        Else
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
            tmcFlashAlert.Interval = 60000  'Every minute
            igAlertFlash = 0
            igAlertTimer = igAlertTimer + 1
        End If
        Exit Sub
    Else
        tmcFlashAlert.Interval = 2000  'Every 2 seconds
        If igAlertFlash Mod 2 = 0 Then
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
        Else
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmWhiteAlertBitmap
        End If
        DoEvents
    End If
End Sub

Private Sub mDisableMenu()
    mnuTaskJob.Enabled = False
    mnuTaskList.Enabled = False
    mnuTqskBackup.Enabled = False
    mnuResetWorkFiles.Enabled = False
    mnuRemoteUser.Enabled = False
    mnuFileImport.Enabled = False
    mnuFileExport.Enabled = False
    
    mnuAccessoryEditor.Enabled = False
    mnuAccessorySetCredit.Enabled = False
    mnuAccessoryShowCredits.Enabled = False
    mnuAccessoryUtilities.Enabled = False
    mnuAccessoryMessages.Enabled = False
    mnuAccessoryBlocks.Enabled = False
    mnuAccessoryUserStatus.Enabled = False
        
End Sub

'Private Sub mSendGGWarningMsg()
'    'Dan M 9/6/11 use email object
'    Dim slStr As String
'    'Dim tlEmail As EmailInformation
'
'    slStr = Trim$(tgSpf.sGClient) & sgCR & sgLF & Trim$(tgSpf.sGAddr(0)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(1)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(2))
'    slStr = slStr & sgCR & sgLF & "The above client system will shutdown in " & igGGFlag & " days"
'    slStr = slStr & sgCR & sgLF & "Please extend the shutdown date to a new date which you can obtain from the Counterpoint service manager"
'    Set ogEmailer = New CEmail
'
'    With ogEmailer
'        .FromAddress = "AClient@Counterpoint.net"
'        .FromName = Trim$(tgSpf.sGClient)
'        .AddTOAddress "Service@counterpoint.net", "Service"
''        .ToAddress = "Service@Counterpoint.net"
''        .ToName = "Service"
'        .Subject = "Client System Shutdown approaching"
'        .Message = slStr
'        .SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, False
'        If Not .Send() Then
'            slStr = "Email could not be sent from form traffic-mSendGGWarningMsg: "
'            gLogMsg slStr, "TrafficErrors.Txt", False
'        End If
'    End With
'    Set ogEmailer = Nothing
'
''    Dim slStr As String
''    Dim tlEmail As EmailInformation
'''Dan M 9.17/09 replaced pbEmail with aspnet mail.
''''11/15/08:  Remove sending message until the other E-Mail program can be used
'''Exit Sub
''
'''    Set objEmailSender = New EmailSender    'Initialize Sender and Email
'''    Set objMyEmail = New Email
'''
'''    Dim objReg As KeyCodeRegistration
'''    Set objReg = New KeyCodeRegistration
'''    Call objReg.EnterKeyCode("QwBvAHUAbgB0AGUAcgBwAG8AaQBuAHQAIABTAG8AZgB0AHcAYQByAGUAAABEAEIARQBBADAANQAxADUAOQA3AEQAQQA3ADUAQQAzADMANwA5ADQAQwBFADgARgBFADYANwA4ADAAMAAwAEIANgBEAEYARgBFAEEAMQAxADgAOQAzAEUAMAAyADEARABCAEEAQwA0ADkARgBCADMAQgBEADEAMAA3ADUANABCAEEANwBBADAANQBEAEIARQAzADQAMQA2ADUAMgA5AEIAQQA5ADkARAA0ADYAMgBEADUARABBADIAMgBCADMARQAwADAAQwA0ADAARAAxAEUAOABCAEMAOQBBAEIAQwBBAA==")
'''
'''    Set objMyEmail = New Email ' Reset Object Var
'''
'''    objMyEmail.FromEmail = "Service@Counterpoint.net"    'Asing the From Information
'''    objMyEmail.FromName = Trim$(tgSpf.sGClient)
''
''    'Reply-To
'''    objMyEmail.ReplyTo = "" 'txtReplyTo.Text
'''
'''    'Organization
'''    objMyEmail.Organization = "" 'txtOrganization.Text
'''
'''    'Recipient Information
'''    objMyEmail.ToEmail = "Service@Counterpoint.Net"
'''    objMyEmail.ToName = "Service"
'''
'''    'Priority
'''    objMyEmail.Priority = Priorities.Normal
''
'''    'Content Type
'''    objMyEmail.ContentType = ContentTypes.Text
''    slStr = Trim$(tgSpf.sGClient) & sgCR & sgLF & Trim$(tgSpf.sGAddr(0)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(1)) & sgCR & sgLF & Trim$(tgSpf.sGAddr(2))
''    slStr = slStr & sgCR & sgLF & "The above client system will shutdown in " & igGGFlag & " days"
''    slStr = slStr & sgCR & sgLF & "Please extend the shutdown date to a new date which you can obtain from the Counterpoint service manager"
'''    objMyEmail.BodyText = slStr
'''    objMyEmail.Subject = "Client System Shutdown approaching"
''
'''    Call objEmailSender.SendEmail(objMyEmail)    'Send the Email
''    With tlEmail
''        .sFromAddress = "Service@Counterpoint.net"
''        ' dan 11/04/09 make sure this is ok
''        .sFromName = Trim$(tgSpf.sGClient)
''        .sToAddress = "Service@Counterpoint.net"
''        .sToName = "Service"
''        .sSubject = "Client System Shutdown approaching"
''        .sMessage = slStr
''        '.bUserFromHasPriority = True
''    End With
''
''    If Not gEmailWithCounterpointHost(tlEmail) Then
''        slStr = "Email could not be sent from form traffic-mSendGGWarningMsg: "
''        gLogMsg slStr, "TrafficErrors.Txt", False
''    End If
'End Sub

'Private Sub objEmailSender_SendCompleted(ByVal EventArgs As PBEmail7.SendCompletedArgs)
'Dim slMsg As String
'
'    If EventArgs.Result.Result = Results.AllOk Then 'Show the Result in the Sendings List in the Email Tag position
'
'    Else
'        slMsg = "A general error has occured in the Email form: " & EventArgs.Result.Description
'        gLogMsg slMsg, "TrafficErrors.Txt", False
'        gMsgBox slMsg, vbCritical, "Email Error"
'    End If
'
'
'End Sub
Private Sub tmcInitAlert_Timer()
    If imReturnToAlerts Then
        tmcInitAlert.Enabled = False
        mnuAlert_Click
        imReturnToAlerts = False
        imReturnToDashboard = False
    End If
    'TTP 10937 - Contract Dashboard: new screen to show contract header information
    If imReturnToDashboard Then
        tmcInitAlert.Enabled = False
        Basic10.cmcDashboard_MouseUp 0, 0, 0, 0
        imReturnToDashboard = False
        imReturnToAlerts = False
    End If
    
End Sub

'Check if participant moved to Pif
Private Function mCheckParticipant() As Integer
    Dim llPifNoRecs As Long
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilHeaderCreated As Integer
    Dim ilSeqNo As Integer
    Dim tlVef As VEF

    hmPif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPif, "", sgDBPath & "Pif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        llPifNoRecs = btrRecords(hmPif)
        If llPifNoRecs = 0 Then
            imPifRecLen = Len(tmPif)
            hmVef = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                imVefRecLen = Len(tgMVef(LBound(tgMVef)))
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    ilHeaderCreated = False
                    ilSeqNo = 1
                    For ilIndex = LBound(tgMVef(ilVef).iMnfSSCode) To UBound(tgMVef(ilVef).iMnfSSCode) Step 1
                        If tgMVef(ilVef).iMnfSSCode(ilIndex) > 0 Then
                            If ilIndex = LBound(tgMVef(ilVef).iMnfSSCode) Then
                                tgMVef(ilVef).iOwnerMnfCode = tgMVef(ilVef).iMnfGroup(ilIndex)
                                tgMVef(ilVef).iSSMnfCode = tgMVef(ilVef).iMnfSSCode(ilIndex)
                                Do
                                    tmVefSrchKey.iCode = tgMVef(ilVef).iCode
                                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilRet = btrUpdate(hmVef, tgMVef(ilVef), imVefRecLen)
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                '11/26/17
                                gFileChgdUpdate "vef.btr", True

                            End If
                            tmPif.lCode = 0
                            tmPif.iVefCode = tgMVef(ilVef).iCode
                            gPackDate "1/1/1990", tmPif.iStartDate(0), tmPif.iStartDate(1)
                            gPackDate "12/31/2069", tmPif.iEndDate(0), tmPif.iEndDate(1)
                            tmPif.iSeqNo = ilSeqNo
                            tmPif.iMnfSSCode = tgMVef(ilVef).iMnfSSCode(ilIndex)
                            tmPif.iMnfGroup = tgMVef(ilVef).iMnfGroup(ilIndex)
                            tmPif.sUpdateRVF = tgMVef(ilVef).sUpdateRVF(ilIndex)
                            tmPif.sExtUpdateRvf = tgMVef(ilVef).sExtUpdateRvf(ilIndex)
                            tmPif.iProdPct = tgMVef(ilVef).iProdPct(ilIndex)
                            tmPif.sUnused = ""
                            ilRet = btrInsert(hmPif, tmPif, imPifRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                MsgBox "Call Counterpoint as Participant Creation Failed, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
                                mCheckParticipant = False
                                Exit Function
                            End If
                            ilSeqNo = ilSeqNo + 1
                        End If
                    Next ilIndex
                Next ilVef
            End If
        End If
    Else
        MsgBox "Call Counterpoint as Unable to Open Participant File, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
        mCheckParticipant = False
        Exit Function
    End If
    btrDestroy hmVef
    btrDestroy hmPif
    mCheckParticipant = True
    Exit Function
End Function

Private Sub mCheckDiskFreeSpace()
    Dim ilPos As Integer
    Dim slDriveLetter As String
    Dim llSectorsPerCluster As Long
    Dim llBytesPerSector As Long
    Dim llNumberOfFreeClusters As Long
    Dim llTotalNumberOfClusters As Long
    Dim llRet As Long
    Dim flFreeGigBytes As Single

    If tgUrf(0).iSlfCode <= 0 Then
        ilPos = InStr(1, sgDBPath, ":", vbTextCompare)
        If ilPos > 0 Then
            slDriveLetter = Left$(sgDBPath, ilPos)
        Else
            ilPos = InStr(1, sgDBPath, "\", vbTextCompare)
            If ilPos > 0 Then
                ilPos = InStr(ilPos + 2, sgDBPath, "\", vbTextCompare)
                If ilPos > 0 Then
                    slDriveLetter = Left$(sgDBPath, ilPos)
                End If
            End If
        End If
        llRet = GetDiskFreeSpace(slDriveLetter, llSectorsPerCluster, llBytesPerSector, llNumberOfFreeClusters, llTotalNumberOfClusters)
        If llRet <> 0 Then
            flFreeGigBytes = (CDbl(llNumberOfFreeClusters) * llSectorsPerCluster * llBytesPerSector) / CDbl(1073741824)
            If flFreeGigBytes <= 5# Then
                llRet = 10 * flFreeGigBytes
                MsgBox "Warning: only " & gLongToStrDec(llRet, 1) & " gig space remains on the server, inform your IT department that more free disk space is required", vbOKOnly + vbExclamation, "WARNING"
            End If
        End If
    End If

End Sub

Private Sub mCheckShortDateForm()
    Dim slDate As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer

    slDate = Format$(Now, "m/d/yy")
    ilPos1 = InStr(1, slDate, "/", vbTextCompare)
    ilPos2 = InStr(ilPos1 + 1, slDate, "/", vbTextCompare)
    If ilPos2 <= 0 Then
        MsgBox "Please change your Regional definition of Short Date Format to m/d/yy", vbOKOnly + vbExclamation, "WARNING"
    Else
        If ilPos2 + 2 < Len(slDate) Then
            MsgBox "Please change your Regional definition of Short Date Format to m/d/yy", vbOKOnly + vbExclamation, "WARNING"
        End If
    End If
End Sub

Private Function mSignOnOff(ilSignOn As Integer) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slPCName As String
    Dim slMACAddr As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDBType As String
    Dim slDate As String

    'hmUlf = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    mSignOnOff = False
    '    Exit Function
    'End If
    imUlfRecLen = Len(tmUlf)  'btrRecordLength(hlUrf)  'Get and save record length
    slPCName = Trim$(gGetComputerName())
    slMACAddr = Trim$(gGetMACs_AdaptInfo())
    slNowDate = Format$(Now, "m/d/yy")
    slNowTime = Format$(Now, "h:mm:ssAM/PM")
    If igTestSystem Then
        tmUlfSrchKey1.sDBType = "T"
    Else
        tmUlfSrchKey1.sDBType = "P"
    End If
    slDBType = tmUlfSrchKey1.sDBType
    ilFound = False
    tmUlfSrchKey1.sSystemType = "T"
    tmUlfSrchKey1.iUrfCode = tgUrf(0).iCode
    tmUlfSrchKey1.iUstCode = 0
    tmUlfSrchKey1.iUieCode = 0
    ilRet = btrGetEqual(hgUlf, tmUlf, imUlfRecLen, tmUlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        If (tmUlf.sDBType = slDBType) And (tmUlf.sSystemType = "T") And (tmUlf.iUrfCode = tgUrf(0).iCode) And (StrComp(Trim$(tmUlf.sPCName), slPCName, vbTextCompare) = 0) And (StrComp(Trim$(tmUlf.sPCMACAddr), slMACAddr, vbTextCompare) = 0) Then
            ilFound = True
            If ilSignOn Then
                gUnpackDate tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1), slDate
                If gDateValue(slDate) = gDateValue("12/31/2069") Then
                    'User Aborted, set UAF
                    mSetUAFToAborted tmUlf.lCode
                End If
                'Update Sign On time
                gPackDate slNowDate, tmUlf.iSignOnDate(0), tmUlf.iSignOnDate(1)
                gPackTime slNowTime, tmUlf.iSignOnTime(0), tmUlf.iSignOnTime(1)
                gPackDate "12/31/2069", tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
                gPackTime "12am", tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)

            Else
                gPackDate slNowDate, tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
                gPackTime slNowTime, tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
            End If
            ilRet = btrUpdate(hgUlf, tmUlf, imUlfRecLen)
            Exit Do
        End If
        ilRet = btrGetNext(hgUlf, tmUlf, imUlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If (Not ilFound) And (ilSignOn) Then
        'Create record with Sign On time
        tmUlf.lCode = 0
        If igTestSystem Then
            tmUlf.sDBType = "T"
        Else
            tmUlf.sDBType = "P"
        End If
        tmUlf.sSystemType = "T"
        tmUlf.iUrfCode = tgUrf(0).iCode
        tmUlf.iUstCode = 0
        tmUlf.iUieCode = 0
        gPackDate slNowDate, tmUlf.iSignOnDate(0), tmUlf.iSignOnDate(1)
        gPackTime slNowTime, tmUlf.iSignOnTime(0), tmUlf.iSignOnTime(1)
        gPackDate "12/31/2069", tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
        gPackTime "12am", tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
        gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
        gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
        tmUlf.sPCName = slPCName
        tmUlf.sPCMACAddr = slMACAddr
        tmUlf.sTimeZone = Left$(gGetLocalTZName(), 1)
        tmUlf.iTrafJobNo = -1
        tmUlf.iTrafListNo = 0
        tmUlf.iTrafRnfCode = 0
        tmUlf.iAffTaskNo = -1
        tmUlf.iAffSubtaskNo = 0
        tmUlf.iAffRptNo = 0
        ilRet = btrInsert(hgUlf, tmUlf, imUlfRecLen, INDEXKEY0)
        ilFound = True
    End If
    If ilFound Then
        lgActiveUlfCode = tmUlf.lCode
        sgActiveLogDate = slNowDate
        sgActiveLogTime = slNowTime
    End If
    'btrDestroy hmUlf
    lgUlfCode = tmUlf.lCode
    ilRet = gObtainSAF()
    If ilSignOn Then
        '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
        igLogActivityStatus = 32123 'Start Log Activity
        gUserActivityLog "L", "Traffic.Frm"
    Else
        '4/2/11: Add setting and call.
        If igLogActivityStatus = 32123 Then
            igLogActivityStatus = -32123    'End Log Activity
            gUserActivityLog "", ""
        End If
    End If
    mSignOnOff = True
End Function

Private Sub tmcMonitor_Timer()
    Dim llDateMonitorChecked As Long
    Dim llTimeMonitorChecked As Long
    Dim ilTask As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim llRunningDate As Long
    Dim llRunningTime As Long
    Dim ll1stStartDate As Long
    Dim ll1stStartTime As Long
    Dim llStartDate As Long
    Dim llStartTime As Long
    Dim llTestStartDate As Long
    Dim llTestStartTime As Long
    Dim ll1stEndDate As Long
    Dim ll1stEndTime As Long
    Dim llEndDate As Long
    Dim llEndTime As Long
    Dim llTestEndDate As Long
    Dim llTestEndTime As Long
    Dim blFound As Boolean
    Dim llNextDate As Long
    Dim ilDay As Integer
    Dim llCurDate As Long
    Dim llCurTime As Long
    Dim slCurDate As String
    Dim slCurTime As String
    Dim llInvDate As Long
    Dim llEMailDate As Long
    Dim llEMailTime As Long
    '8793
    Dim slDebug As String
    
    llDateMonitorChecked = gDateValue(Format(Now, "m/d/yy"))
    llTimeMonitorChecked = gTimeToLong(Format(Now, "h:mm:ssAM/PM"), True)
    If (lgDateMonitorChecked <> llDateMonitorChecked) Or ((llTimeMonitorChecked - lgTimeMonitorChecked >= MONITORRUNINTERVAL)) Then
        For ilTask = 0 To UBound(tgTaskInfo) Step 1
            '8793
            slDebug = ""
            If tgTaskInfo(ilTask).iMenuIndex > 0 Then
                tmTmfSrchKey1.sTaskCode = Trim$(tgTaskInfo(ilTask).sTaskCode)
                ilRet = btrGetEqual(hgTmf, tmTmf, imTmfRecLen, tmTmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'gUnpackDate tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), slDate
                    'gUnpackTime tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), "M", "1", slTime
                    'Select Case smZone
                    '    Case "E"
                    '        slStr = slDate & " " & slTime
                    '    Case "C"
                    '        slStr = DateAdd("h", 1, slDate & " " & slTime)
                    '    Case "M"
                    '        slStr = DateAdd("h", 2, slDate & " " & slTime)
                    '    Case "P"
                    '        slStr = DateAdd("h", 3, slDate & " " & slTime)
                    'End Select
                    'slDate = Format$(slStr, "m/d/yy")
                    'slTime = Format$(slStr, "h:mm:ssAM/PM")
                    'If (gDateValue(slDate) <> lgMonitorRunningDate) Or (gTimeToLong(slTime, False) <> lgMonitorRunningTime) Then
                    '7967
                    If tmTmf.sService = "W" Then
                        '8129
                        tmTmf.sStatus = gMonitorVendor(tmcWebConnectIssue)
                    End If
                    gUnpackDateLong tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), llRunningDate
                    gUnpackTimeLong tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), True, llRunningTime
                    gUnpackDateLong tmTmf.iStartRunDate(0), tmTmf.iStartRunDate(1), llStartDate
                     '8129 this is Dan forces to yellow
                    If tmTmf.sService = "W" And tmTmf.sStatus = "Y" Then
                        llRunningDate = lmDate1970
                    End If
                    If (llRunningDate = lmDate1970) And (llStartDate = lmDate1970) Then
                        If tgTaskInfo(ilTask).lColor <> LIGHTYELLOW Then
                            tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                            mSetMonitorInfo ilTask
                        End If
                    Else
                        If (llRunningDate <> tgTaskInfo(ilTask).lRunningDate) Or (llRunningTime <> tgTaskInfo(ilTask).lRunningTime) Then
                            If tgTaskInfo(ilTask).lColor <> DARKGREEN Then
                                tgTaskInfo(ilTask).lColor = DARKGREEN
                                mSetMonitorInfo ilTask
                            End If
                            tgTaskInfo(ilTask).lRunningDate = llRunningDate
                            tgTaskInfo(ilTask).lRunningTime = llRunningTime
                            tgTaskInfo(ilTask).lElapsedTime = 0
                        Else
                            If (tmTmf.sStatus = "E") Or (tmTmf.sStatus = "A") Then
                                tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVAL Then
                                    slDebug = "DEBUG 1- elapsed time = " & tgTaskInfo(ilTask).lElapsedTime
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        tgTaskInfo(ilTask).lColor = vbRed
                                        mSetMonitorInfo ilTask
                                    End If
                                '7967
                                ElseIf tmTmf.sService = "W" Then
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        tgTaskInfo(ilTask).lColor = vbRed
                                        mSetMonitorInfo ilTask
                                    End If
                                    tgTaskInfo(ilTask).lElapsedTime = 0
                                End If
                            ElseIf tmTmf.sRunMode = "C" Then
                                tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVAL Then
                                    slDebug = "DEBUG 2- elapsed time = " & tgTaskInfo(ilTask).lElapsedTime
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        '7967 exclude web service
                                        If tmTmf.sService <> "W" Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                    tgTaskInfo(ilTask).lElapsedTime = 0
                                End If
                            Else
                                'Periodic
                                slCurDate = Format(Now, "m/d/yy")
                                slCurTime = Format(Now, "h:mm:ssAM/PM")
                                Select Case smZone
                                    Case "E"
                                        slStr = slCurDate & " " & slCurTime
                                    Case "C"
                                        slStr = DateAdd("h", 1, slCurDate & " " & slCurTime)
                                    Case "M"
                                        slStr = DateAdd("h", 2, slCurDate & " " & slCurTime)
                                    Case "P"
                                        slStr = DateAdd("h", 3, slCurDate & " " & slCurTime)
                                End Select
                                llCurDate = gDateValue(Format$(slStr, "m/d/yy"))
                                llCurTime = gTimeToLong(Format$(slStr, "h:mm:ssAM/PM"), False)
                                gUnpackDateLong tmTmf.i1stStartRunDate(0), tmTmf.i1stStartRunDate(1), ll1stStartDate
                                gUnpackTimeLong tmTmf.i1stStartRunTime(0), tmTmf.i1stStartRunTime(1), False, ll1stStartTime
                                gUnpackDateLong tmTmf.i1stEndRunDate(0), tmTmf.i1stEndRunDate(1), ll1stEndDate
                                gUnpackTimeLong tmTmf.i1stEndRunTime(0), tmTmf.i1stEndRunTime(1), True, ll1stEndTime
                                gUnpackDateLong tmTmf.iStartRunDate(0), tmTmf.iStartRunDate(1), llStartDate
                                gUnpackTimeLong tmTmf.iStartRunTime(0), tmTmf.iStartRunTime(1), False, llStartTime
                                gUnpackDateLong tmTmf.iEndRunDate(0), tmTmf.iEndRunDate(1), llEndDate
                                gUnpackTimeLong tmTmf.iEndRunTime(0), tmTmf.iEndRunTime(1), True, llEndTime
                                'If ((llStartDate = llEndDate) And (llEndTime >= llStartTime)) Or ((llStartDate + 1 = llEndDate) And (llEndTime < llStartTime)) Then
                                If (ll1stStartDate = llStartDate) And (ll1stStartTime >= llStartTime) Then
                                    'has started on today but not completed
                                    llTestStartDate = ll1stStartDate
                                    llTestEndDate = ll1stEndDate
                                    llTestStartTime = ll1stStartTime
                                    llTestEndTime = ll1stEndTime
                                Else
                                    'last run was yesterday
                                    llTestStartDate = llStartDate
                                    llTestEndDate = llEndDate
                                    llTestStartTime = llStartTime
                                    llTestEndTime = llEndTime
                                End If
                                If ((llTestStartDate = llTestEndDate) And (llTestEndTime >= llTestStartTime)) Or ((llTestStartDate + 1 = llTestEndDate) And (llTestEndTime < llTestStartTime)) Then
                                    'Determine next date that it should run
                                    If Trim(tmTmf.sMonthPeriod) = "" Then
                                        blFound = False
                                        If llTestEndTime >= llTestStartTime Then
                                            llNextDate = llTestEndDate + 1
                                        Else
                                            llNextDate = llTestEndDate
                                        End If
                                        Do
                                            ilDay = gWeekDayLong(llNextDate)
                                            Select Case ilDay
                                                Case 0
                                                    If tmTmf.sMo = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 1
                                                    If tmTmf.sTu = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 2
                                                    If tmTmf.sWe = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 3
                                                    If tmTmf.sTh = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 4
                                                    If tmTmf.sFr = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 5
                                                    If tmTmf.sSa = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 6
                                                    If tmTmf.sSu = "Y" Then
                                                        blFound = True
                                                    End If
                                            End Select
                                            If blFound Then
                                                Exit Do
                                            End If
                                            llNextDate = llNextDate + 1
                                        Loop While Not blFound
                                    Else
                                        Select Case tmTmf.sMonthPeriod
                                            Case "SE"
                                                slDate = gObtainEndStd(Format(llTestStartDate + 1, "m/d/yy"))
                                            Case "CE"
                                                slDate = gObtainEndCal(Format(llTestStartDate + 1, "m/d/yy"))
                                            Case "IE"
                                                gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llInvDate
                                                slDate = gObtainEndStd(Format(llInvDate + 1, "m/d/yy"))
                                        End Select
                                        llNextDate = gDateValue(DateAdd("d", tmTmf.iDaysAfter, slDate))
                                    End If
                                    If (llCurDate < llNextDate) Or (llCurDate = llNextDate) And (llCurTime < llTestStartTime) Then
                                        If tgTaskInfo(ilTask).lColor <> DARKGREEN Then
                                            tgTaskInfo(ilTask).lColor = DARKGREEN
                                            mSetMonitorInfo ilTask
                                        End If
                                    Else
                                        slDebug = "DEBUG 3- current date = " & Format(llCurDate, sgSQLDateForm) & " next date = " & Format(llNextDate, sgSQLDateForm) & " current time = " & gFormatTimeLong(llCurTime, "A", "1") & " test start time = " & gFormatTimeLong(llTestStartTime, "A", "1")
                                        If tgTaskInfo(ilTask).lColor <> vbRed Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                ElseIf (llTestStartDate > llTestEndDate) Or ((llTestStartDate = llTestEndDate) And (llTestStartTime > llTestEndTime)) Then
                                    tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                    If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVALPERIOD Then
                                        slDebug = "DEBUG 4- test start date = " & Format(llTestStartDate, sgSQLDateForm) & " test end date = " & Format(llTestEndDate, sgSQLDateForm) & " test start time = " & gFormatTimeLong(llTestStartTime, "A", "1") & " test end time = " & gFormatTimeLong(llTestEndTime, "A", "1")
                                        If tgTaskInfo(ilTask).lColor <> vbRed Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                        tgTaskInfo(ilTask).lElapsedTime = 0
                                    ElseIf tgTaskInfo(ilTask).lColor <> vbRed Then
                                        If tgTaskInfo(ilTask).lColor <> LIGHTYELLOW Then
                                            tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                End If
                            End If
                            If tgTaskInfo(ilTask).lColor = vbRed Then
                                slCurDate = Format(Now, "m/d/yy")
                                slCurTime = Format(Now, "h:mm:ssAM/PM")
                                Select Case smZone
                                    Case "E"
                                        slStr = slCurDate & " " & slCurTime
                                    Case "C"
                                        slStr = DateAdd("h", 1, slCurDate & " " & slCurTime)
                                    Case "M"
                                        slStr = DateAdd("h", 2, slCurDate & " " & slCurTime)
                                    Case "P"
                                        slStr = DateAdd("h", 3, slCurDate & " " & slCurTime)
                                End Select
                                llCurDate = gDateValue(Format$(slStr, "m/d/yy"))
                                llCurTime = gTimeToLong(Format$(slStr, "h:mm:ssAM/PM"), False)
                                gUnpackDateLong tmTmf.iEMailSentDate(0), tmTmf.iEMailSentDate(1), llEMailDate
                                If (llEMailDate <> llCurDate) Then
                                    gUnpackDateLong tmTmf.iEMailReqDate(0), tmTmf.iEMailReqDate(1), llEMailDate
                                    If (llEMailDate <> llCurDate) Then
                                        gPackDateLong llCurDate, tmTmf.iEMailReqDate(0), tmTmf.iEMailReqDate(1)
                                        gPackTimeLong llCurTime, tmTmf.iEMailReqTime(0), tmTmf.iEMailReqTime(1)
                                        ilRet = btrUpdate(hgTmf, tmTmf, imTmfRecLen)
                                    Else
                                        gUnpackTimeLong tmTmf.iEMailReqTime(0), tmTmf.iEMailReqTime(1), False, llEMailTime
                                        If llCurTime - llEMailTime > MONITOREMAILINTERVAL And emMonitoringAllowed <> MonitorNoEmail Then
                                            gPackDateLong llCurDate, tmTmf.iEMailSentDate(0), tmTmf.iEMailSentDate(1)
                                            ilRet = btrUpdate(hgTmf, tmTmf, imTmfRecLen)
                                            '8793
                                            If Len(slDebug) = 0 Then
                                                slDebug = "Hit no debug points!"
                                            End If
                                            slDebug = smMonitorID & " User: " & Trim$(tgUrf(0).sName) & " Time Zone: " & smZone & " Debug Info: Running Date = " & Format$(llRunningDate, sgSQLDateForm) & " Running Time = " & gFormatTimeLong(llRunningTime, "A", "1") & " " & slDebug
                                            'Dan: Generate E-Mail at this point...don't send if logged in as csi
                                            If Not bgInternalGuide Then
                                                'gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Traffic has detected a problem for " & Trim$(tgSpf.sGClient) & " at " & gNow()
                                                '8793
                                                'gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Traffic has detected a problem for " & Trim$(tgSpf.sGClient) & " at " & Now
                                                gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Traffic has detected a problem for " & Trim$(tgSpf.sGClient) & " at " & Now & slDebug
                                            Else
                                                'test only!
                                               ' mTempSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Traffic has detected a problem for " & Trim$(tgSpf.sGClient) & " at " & Now & slDebug
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilTask
        lgDateMonitorChecked = llDateMonitorChecked
        lgTimeMonitorChecked = llTimeMonitorChecked
    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    mnuTaskExit_Click
End Sub

Private Sub mCheckActiveLog()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llNowDate                     llNowTime                                               *
'******************************************************************************************

    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilRet As Integer
    Dim llMinuteDiff As Long

    slNowDate = Format$(Now, "m/d/yy")
    slNowTime = Format$(Now, "h:mm:ssAM/PM")
    If sgActiveLogDate = "" Then
        Exit Sub
    End If
    If sgActiveLogTime = "" Then
        Exit Sub
    End If
    If gValidDate(sgActiveLogDate) = False Then
        Exit Sub
    End If
    llMinuteDiff = DateDiff("n", sgActiveLogDate & " " & sgActiveLogTime, slNowDate & " " & slNowTime)
    If llMinuteDiff > 60 Then
        'hmUlf = CBtrvTable(TWOHANDLES)
        'ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet = BTRV_ERR_NONE Then
            imUlfRecLen = Len(tmUlf)  'btrRecordLength(hlUrf)  'Get and save record length
            tmUlfSrchKey.lCode = lgActiveUlfCode
            ilRet = btrGetEqual(hgUlf, tmUlf, imUlfRecLen, tmUlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
                ilRet = btrUpdate(hgUlf, tmUlf, imUlfRecLen)
            End If
        'End If
        sgActiveLogDate = slNowDate
        sgActiveLogTime = slNowTime
        'btrDestroy hmUlf
    End If
End Sub

Private Function mCheckBuyer() As Integer
    Dim ilRet As Integer
    Dim ilPass As Integer
    Dim ilPnfBuyer As Integer
    Dim llLastCode As Long

    ReDim tmPnfBuyer(0 To 0) As TYPESORT

    hmCHF = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        imCHFRecLen = Len(tmChf)

        ilRet = btrGetFirst(hmCHF, tmChf, imCHFRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If tmChf.iPnfBuyer <> 0 Then
                btrDestroy hmCHF
                mCheckBuyer = True
                Exit Function
            End If
            If UBound(tmPnfBuyer) <= LBound(tmPnfBuyer) Then
                ilRet = mGetPnfBuyer()
                If Not ilRet Then
                    btrDestroy hmCHF
                    mCheckBuyer = False
                    Exit Function
                End If
            End If
            ilPnfBuyer = 0
            If Trim$(tmChf.sBuyer) <> "" Then
                ilRet = mBinarySearchPnfBuyer(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.sBuyer)
            Else
                ilRet = -1
            End If
            If ilRet <> -1 Then
                ilPnfBuyer = tmPnfBuyer(ilRet).lRecPos
            Else
                If tmChf.iAgfCode <= 0 Then
                    ilRet = gBinarySearchAdf(tmChf.iAdfCode)
                    If ilRet <> -1 Then
                        ilPnfBuyer = tgCommAdf(ilRet).iPnfBuyer
                    End If
                Else
                    ilRet = gBinarySearchAgf(tmChf.iAgfCode)
                    If ilRet <> -1 Then
                        ilPnfBuyer = tgCommAgf(ilRet).iPnfBuyer
                    End If
                End If
            End If
            llLastCode = tmChf.lCode
            If ilPnfBuyer <> 0 Then
                tmChf.iPnfBuyer = ilPnfBuyer
                ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
            End If
            Do
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            Loop While (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = llLastCode)
        Loop

        For ilPass = 0 To 1 Step 1
            If ilPass = 0 Then
                hmRvf = CBtrvTable(TWOHANDLES)
                ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            Else
                hmRvf = CBtrvTable(TWOHANDLES)
                ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            End If
            If ilRet = BTRV_ERR_NONE Then
                imRvfRecLen = Len(tmRvf)
                ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                Do While ilRet = BTRV_ERR_NONE
                    If tmRvf.iPnfBuyer <> 0 Then
                        btrDestroy hmCHF
                        btrDestroy hmRvf
                        mCheckBuyer = True
                        Exit Function
                    End If
                    If UBound(tmPnfBuyer) <= LBound(tmPnfBuyer) Then
                        ilRet = mGetPnfBuyer()
                        If Not ilRet Then
                            btrDestroy hmCHF
                            btrDestroy hmRvf
                            mCheckBuyer = False
                            Exit Function
                        End If
                    End If

                    ilPnfBuyer = 0
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) Then
                        If Trim$(tmChf.sBuyer) <> "" Then
                            ilRet = mBinarySearchPnfBuyer(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.sBuyer)
                        Else
                            ilRet = -1
                        End If
                    Else
                        ilRet = -1
                    End If
                    If ilRet <> -1 Then
                        ilPnfBuyer = tmPnfBuyer(ilRet).lRecPos
                    Else
                        If tmRvf.iAgfCode <= 0 Then
                            ilRet = gBinarySearchAdf(tmRvf.iAdfCode)
                            If ilRet <> -1 Then
                                ilPnfBuyer = tgCommAdf(ilRet).iPnfBuyer
                            End If
                        Else
                            ilRet = gBinarySearchAgf(tmRvf.iAgfCode)
                            If ilRet <> -1 Then
                                ilPnfBuyer = tgCommAgf(ilRet).iPnfBuyer
                            End If
                        End If
                    End If
                    llLastCode = tmRvf.lCode
                    If ilPnfBuyer <> 0 Then
                        tmRvf.iPnfBuyer = ilPnfBuyer
                        ilRet = btrUpdate(hmRvf, tmRvf, imRvfRecLen)
                    End If
                    Do
                        ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop While (ilRet = BTRV_ERR_NONE) And (tmRvf.lCode = llLastCode)
                Loop
            Else
                btrDestroy hmRvf
                MsgBox "Call Counterpoint as Unable to Open Receivable File, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
                mCheckBuyer = False
                Exit Function
            End If
            btrDestroy hmRvf
        Next ilPass

    Else
        btrDestroy hmCHF
        MsgBox "Call Counterpoint as Unable to Open Contract File, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
        mCheckBuyer = False
        Exit Function
    End If
    btrDestroy hmCHF
    Erase tmPnfBuyer
    mCheckBuyer = True
    Exit Function
End Function

Private Function mGetPnfBuyer() As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim slCode As String

    hmPnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        imPnfRecLen = Len(tmPnf)
        ilRet = btrGetFirst(hmPnf, tmPnf, imPnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If tmPnf.sType = "B" Then
                If tmPnf.iAdfCode > 0 Then
                    slName = "ADF"
                    slCode = Trim$(str$(tmPnf.iAdfCode))
                Else
                    slName = "AGF"
                    slCode = Trim$(str$(tmPnf.iAgfCode))
                End If
                Do While Len(slCode) < 5
                    slCode = "0" & slCode
                Loop
                'Contract buyer defined with a length of 20 characters, so limit compare to the 20 characters
                tmPnfBuyer(UBound(tmPnfBuyer)).sKey = slName & slCode & Left$(tmPnf.sName, 20)
                tmPnfBuyer(UBound(tmPnfBuyer)).lRecPos = tmPnf.iCode
                ReDim Preserve tmPnfBuyer(0 To UBound(tmPnfBuyer) + 1) As TYPESORT
            End If
            ilRet = btrGetNext(hmPnf, tmPnf, imPnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If UBound(tmPnfBuyer) > 0 Then
            ArraySortTyp fnAV(tmPnfBuyer(), 0), UBound(tmPnfBuyer), 0, LenB(tmPnfBuyer(0)), 0, LenB(tmPnfBuyer(0).sKey), 0
        End If
    Else
        btrDestroy hmPnf
        MsgBox "Call Counterpoint as Unable to Open Buyer Name File, error # " & str$(ilRet), vbOKOnly + vbCritical, "Error"
        mGetPnfBuyer = False
        Exit Function
    End If
    btrDestroy hmPnf
    mGetPnfBuyer = True
    Exit Function

End Function
Private Function mGetDBID() As Integer
    Dim tlNrf As NRF
    Dim ilRet As Integer
    Dim ilNrfRecLen As Integer
    Dim hlNrf As Integer

    sgRepDBID = ""
    sgNetDBID = ""
    hlNrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlNrf, "", sgDBPath & "Nrf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        ilNrfRecLen = Len(tlNrf)
        ilRet = btrGetFirst(hlNrf, tlNrf, ilNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                If tlNrf.sType = "R" Then
                    sgRepDBID = tlNrf.sDBID
                    Exit Do
                End If
            ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                If tlNrf.sType = "N" Then
                    sgNetDBID = tlNrf.sDBID
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hlNrf, tlNrf, ilNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
        mGetDBID = True
    Else
        mGetDBID = False
    End If
    btrDestroy hlNrf
    Exit Function
End Function

Private Function mBinarySearchPnfBuyer(ilAdfCode As Integer, ilAgfCode As Integer, slBuyerName) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim slName As String
    Dim slCode As String

    If ilAgfCode <= 0 Then
        slName = "ADF"
        slCode = Trim$(str$(ilAdfCode))
    Else
        slName = "AGF"
        slCode = Trim$(str$(ilAgfCode))
    End If
    Do While Len(slCode) < 5
        slCode = "0" & slCode
    Loop
    slName = slName & slCode & Trim(slBuyerName)
    ilMin = LBound(tmPnfBuyer)
    ilMax = UBound(tmPnfBuyer) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If StrComp(Trim$(tmPnfBuyer(ilMiddle).sKey), slName, vbTextCompare) = 0 Then
            'found the match
            mBinarySearchPnfBuyer = ilMiddle
            Exit Function
        ElseIf StrComp(Trim$(tmPnfBuyer(ilMiddle).sKey), slName, vbTextCompare) > 0 Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchPnfBuyer = -1
End Function
Private Sub mLimitGuideMenus()
Dim Control As Control
Dim slVisibleMenuItems(11) As String
Dim c As Integer
Dim blSkip As Boolean
If (Trim$(tgUrf(0).sName) = sgSUName And Not bgInternalGuide) Then
    mSetVisibleMenuList slVisibleMenuItems
     For Each Control In Me.Controls
        If TypeOf Control Is Menu Then
        blSkip = False
        On Error Resume Next
            For c = 0 To 11
                If Control.Name = slVisibleMenuItems(c) Then
                    blSkip = True
                    Exit For
                End If
            Next c
            If Not blSkip Then
                Control.Visible = False
            End If
        End If
     Next
End If
End Sub
Private Sub mSetVisibleMenuList(ByRef slVisibleMenuItems() As String)
slVisibleMenuItems(0) = "mnuFile"
slVisibleMenuItems(1) = "mnuTaskJob"
slVisibleMenuItems(2) = "mnuTaskList"
slVisibleMenuItems(3) = "mnuTaskExit"
slVisibleMenuItems(4) = "mnuAccessory"
slVisibleMenuItems(5) = "mnuAccessoryCalendar"
slVisibleMenuItems(6) = "mnuAccessoryCalculator"
slVisibleMenuItems(7) = "mnuHelp"
slVisibleMenuItems(8) = "mnuContentHelp"
slVisibleMenuItems(9) = "mnuSearchHelp"
slVisibleMenuItems(10) = "mnuUsingHelp"
slVisibleMenuItems(11) = "mnuAboutHelp"
End Sub
'Private Sub mLimitUserMenu()
'If tgUrf(0).sChangeCSIDate = "I" Then
'    mnuAccessoryDate.Visible = True
'Else
'    mnuAccessoryDate.Visible = False
'End If
'End Sub



Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'web browser added so pic can show animated gif  Dan 9-01-09
End Sub

Private Sub mCheckAndPopEDIServiceNames()
    Dim ilRet As Integer
    Dim ilDonovanFd As Integer
    Dim ilKatzFd As Integer
    Dim ilRadioInvFd As Integer
    Dim ilWarrenLambFd As Integer
    Dim ilSpotDataFd As Integer     '11-12-10
    Dim ilMediaBankFd As Integer    '11-12-10
    Dim ilPDFEMailFd As Integer
    Dim ilMOceanOmni As Integer     'TTP 10897 - EDI new format: MOcean Omni
    Dim ilLoop As Integer
    Dim ilSave As Integer
    Dim ilUpper As Integer
    
    ReDim tgEDIServiceInfo(0 To 0) As EDISERVICEINFO
    hmArf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmArf, "", sgDBPath & "ARF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    imArfRecLen = Len(tmArf)
    ilDonovanFd = False
    ilKatzFd = False
    ilRadioInvFd = False
    ilWarrenLambFd = False
    ilSpotDataFd = False
    ilMediaBankFd = False
    ilPDFEMailFd = False
    ilMOceanOmni = False
    ilRet = btrGetFirst(hmArf, tmArf, imArfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tmArf.sType = "A" Then
            ilSave = False
            'If StrComp(Trim$(tmArf.sID), "Donovan", vbTextCompare) = 0 Then
            If (StrComp(Trim$(tmArf.sID), "Donovan", vbTextCompare) = 0) Or (StrComp(Trim$(tmArf.sID), "MOcean DS", vbTextCompare) = 0) Then
                ilDonovanFd = True
                ilSave = True
            End If
            If StrComp(Trim$(tmArf.sID), "Katz", vbTextCompare) = 0 Then
                ilKatzFd = True
                ilSave = True
            End If
            If StrComp(Trim$(tmArf.sID), "RadioInv", vbTextCompare) = 0 Then
                ilRadioInvFd = True
                ilSave = True
            End If
            If StrComp(Trim$(tmArf.sID), "WarrenLamb", vbTextCompare) = 0 Then
                ilWarrenLambFd = True
                ilSave = True
            End If
            If StrComp(Trim$(tmArf.sID), "SpotData", vbTextCompare) = 0 Then
                ilSpotDataFd = True
                ilSave = True
            End If
            'If StrComp(Trim$(tmArf.sID), "MediaBank", vbTextCompare) = 0 Then
            If (StrComp(Trim$(tmArf.sID), "MediaBank", vbTextCompare) = 0) Or (StrComp(Trim$(tmArf.sID), "MOcean OX", vbTextCompare) = 0) Then
                ilMediaBankFd = True
                ilSave = True
            End If
            'TTP 10897 - EDI new format: MOcean Omni
            If (StrComp(Trim$(tmArf.sID), "Omni", vbTextCompare) = 0) Then
                ilMOceanOmni = True
                ilSave = True
            End If
            If StrComp(Trim$(tmArf.sID), "PDF EMail", vbTextCompare) = 0 Then
                ilPDFEMailFd = True
                '10/12/16: PDF EMail test so that only CSI will see it
                '11-21-16 remove test to test for CSI on PDF email.  Full feature implemented
'                If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                    ilSave = True
'                End If
            End If
            If ilSave Then
                ilUpper = UBound(tgEDIServiceInfo)
                tgEDIServiceInfo(ilUpper).tArf = tmArf
                tgEDIServiceInfo(ilUpper).hEDI = -1
                tgEDIServiceInfo(ilUpper).sPathFile = ""
                tgEDIServiceInfo(ilUpper).sUsed = "N"
                ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
            End If
        End If
        ilRet = btrGetNext(hmArf, tmArf, imArfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    For ilLoop = 0 To 3 Step 1
        tmArf.sNmAd(ilLoop) = ""
    Next ilLoop
    tmArf.sType = "A"
    tmArf.sSepInvByVeh = "N"
    tmArf.sEMail = ""
    tmArf.sFTP = ""
    tmArf.sSendISCITo = ""
    tmArf.sWebSite = ""
    tmArf.sContactName = ""
    tmArf.sContactPhone = ""
    tmArf.sContactFax = ""
    If Not ilDonovanFd Then
        tmArf.iCode = 0
        tmArf.sID = "MOcean DS" '"Donovan"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilKatzFd Then
        tmArf.iCode = 0
        tmArf.sID = "Katz"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilRadioInvFd Then
        tmArf.iCode = 0
        tmArf.sID = "RadioInv"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilWarrenLambFd Then
        tmArf.iCode = 0
        tmArf.sID = "WarrenLamb"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilSpotDataFd Then
        tmArf.iCode = 0
        tmArf.sID = "SpotData"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilMediaBankFd Then
        tmArf.iCode = 0
        tmArf.sID = "MOcean OX" '"MediaBank"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    'TTP 10897 - EDI new format: MOcean Omni
    If Not ilMOceanOmni Then
        tmArf.iCode = 0
        tmArf.sID = "Omni"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        ilUpper = UBound(tgEDIServiceInfo)
        tgEDIServiceInfo(ilUpper).tArf = tmArf
        tgEDIServiceInfo(ilUpper).hEDI = -1
        tgEDIServiceInfo(ilUpper).sUsed = "N"
        ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
    End If
    If Not ilPDFEMailFd Then
        tmArf.iCode = 0
        tmArf.sID = "PDF EMail"
        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, 0)
        '10/12/16: PDF EMail test so that only CSI will see it
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            ilUpper = UBound(tgEDIServiceInfo)
            tgEDIServiceInfo(ilUpper).tArf = tmArf
            tgEDIServiceInfo(ilUpper).hEDI = -1
            tgEDIServiceInfo(ilUpper).sUsed = "N"
            ReDim Preserve tgEDIServiceInfo(0 To ilUpper + 1) As EDISERVICEINFO
        End If
    End If
    'Check if name needs to be changed
    For ilLoop = 0 To UBound(tgEDIServiceInfo) - 1 Step 1
        If (StrComp(Trim$(tgEDIServiceInfo(ilLoop).tArf.sID), "Donovan", vbTextCompare) = 0) Then
            tmArfSrchKey.iCode = tgEDIServiceInfo(ilLoop).tArf.iCode
            ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmArf.sID = "MOcean DS"
                ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
            End If
        End If
        If (StrComp(Trim$(tgEDIServiceInfo(ilLoop).tArf.sID), "MediaBank", vbTextCompare) = 0) Then
            tmArfSrchKey.iCode = tgEDIServiceInfo(ilLoop).tArf.iCode
            ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmArf.sID = "MOcean OX"
                ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
            End If
        End If
    Next ilLoop
    ilRet = btrClose(hmArf)
    btrDestroy hmArf
End Sub

Private Function mUpdatesAvailable() As Boolean
    Dim myInfo As csiInstallerInfo
    Dim ilPos As Integer
    Dim slCsiSystemPath As String
    Dim blRet As Boolean
    Dim MyFile As FileSystemObject
    Dim ilVersion As Integer
    
    ilVersion = mGetCsiVersion()
    Set MyFile = New FileSystemObject
    Set myInfo = New csiInstallerInfo
    ilPos = InStr(1, sgExePath, "\exe", vbTextCompare)
    slCsiSystemPath = Mid$(sgExePath, 1, ilPos) & "Setup\ClientInstall\System32"
    If Not MyFile.FolderExists(slCsiSystemPath) Then
        ilPos = InStrRev(sgDBPath, "\", Len(sgDBPath) - 1)
        If ilPos > 0 Then
            slCsiSystemPath = Mid$(sgDBPath, 1, ilPos) & "Setup\ClientInstall\System32"
        End If
        If Not MyFile.FolderExists(slCsiSystemPath) Then
            ilPos = InStrRev(sgImportPath, "\", Len(sgImportPath) - 1)
            If ilPos > 0 Then
                slCsiSystemPath = Mid$(sgImportPath, 1, ilPos) & "Setup\ClientInstall\System32"
            End If
            If Not MyFile.FolderExists(slCsiSystemPath) Then
               '5676
                gLogMsg "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", "UpdateErrors.Txt", False
                'gMsgBox "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", -1, ""
                mUpdatesAvailable = False
                GoTo CLEANUP
            End If
        End If
    End If
    If myInfo.Start(slCsiSystemPath, ilVersion) Then
        If Not myInfo.CompareCheckers() Then
            If myInfo.FailureReason = NOERROR Then
                blRet = True
            Else
                '5676
                gLogMsg "Problem in Traffic:mUpdatesAvailable.  " & myInfo.FailureMessage, "UpdateErrors.Txt", False
                'gMsgBox "Problem in Traffic:mUpdatesAvailable.  " & myInfo.FailureMessage, -1, ""
            End If
        End If
    Else
        '5676
        gLogMsg "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", "UpdateErrors.Txt", False
        'gMsgBox "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", -1, ""
    End If
    mUpdatesAvailable = blRet
CLEANUP:
    Set MyFile = Nothing
    Set myInfo = Nothing
End Function
Private Function mGetCsiVersion() As Integer
    
    mGetCsiVersion = CInt(App.Major & App.Minor)
End Function
Private Sub mSetUAFToAborted(llUlfCode As Long)
    Dim ilRet As Integer
    Dim slStatus As String
    Dim llUaf As Long
    ReDim llUafCode(0 To 0) As Long
    
    hmUaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUaf, "", sgDBPath & "Uaf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    imUafRecLen = Len(tmUaf)  'Get and save ARF record length
    slStatus = "I"
    tmUafSrchKey2.lUlfCode = llUlfCode
    tmUafSrchKey2.sStatus = slStatus
    gPackDate Format(Now, "m/d/yy"), tmUafSrchKey2.iStartDate(0), tmUafSrchKey2.iStartDate(1)
    gPackTime "11:59:59PM", tmUafSrchKey2.iStartTime(0), tmUafSrchKey2.iStartTime(1)
    ilRet = btrGetGreaterOrEqual(hmUaf, tmUaf, imUafRecLen, tmUafSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmUaf.lUlfCode = llUlfCode) And (tmUaf.sStatus = slStatus)
        llUafCode(UBound(llUafCode)) = tmUaf.lCode
        ReDim Preserve llUafCode(0 To UBound(llUafCode) + 1) As Long
        ilRet = btrGetNext(hmUaf, tmUaf, imUafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    llUaf = 0
    Do While llUaf <= UBound(llUafCode) - 1
        tmUafSrchKey0.lCode = llUafCode(llUaf)
        ilRet = btrGetEqual(hmUaf, tmUaf, imUafRecLen, tmUafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmUaf.sStatus = "A"
            gPackDate Format(Now, "m/d/yy"), tmUaf.iEndDate(0), tmUaf.iEndDate(1)
            gPackTime Format(Now, "h:mm:ssAM/PM"), tmUaf.iEndTime(0), tmUaf.iEndTime(1)
            ilRet = btrUpdate(hmUaf, tmUaf, imUafRecLen)
            If ilRet <> BTRV_ERR_CONFLICT Then
                llUaf = llUaf + 1
            End If
        Else
            llUaf = llUaf + 1
        End If
    Loop
    btrClose hmUaf
    btrDestroy hmUaf
End Sub


Private Sub mExport(slExportName As String)
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim blStart As Boolean
    Dim ilRet As Integer
    
    sgExportMenuItem = slExportName
    On Error GoTo LoadErr
    ilRet = 0
    AppActivate "CSI Exports"
    If ilRet = 0 Then
        MsgBox "Operation Disallow as Export module is busy"
        Exit Sub
    End If
    ExportMenu.Show vbModal
    
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub mImportResearch(slImportName As String)
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim blStart As Boolean
    Dim ilRet As Integer
    
    If tgSpf.sGUsePropSys <> "Y" Then
        Exit Sub
    End If
    sgImportResearchMenuItem = slImportName
    On Error GoTo LoadErr
    ilRet = 0
    AppActivate "Import Research"
    If ilRet = 0 Then
        MsgBox "Operation Disallow as Import Research module is busy"
        Exit Sub
    End If
    ImportResearchMenu.Show vbModal
    
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
End Sub
Private Function mCheckCopyUsage() As Integer
    Dim ilRet As Integer
    Dim llCufNoRecs As Long
    Dim llCrfNoRecs As Long
    Dim ilAdd As Integer
    Dim ilVef As Integer
    Dim slActiveDate As String
    
    mCheckCopyUsage = True
    hmCuf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCuf, "", sgDBPath & "Cuf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        llCufNoRecs = btrRecords(hmCuf)
        If llCufNoRecs = 0 Then
            imCufRecLen = Len(tmCuf)
            hmCrf = CBtrvTable(ONEHANDLE)
            ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                llCrfNoRecs = btrRecords(hmCrf)
                If llCrfNoRecs > 0 Then
                    imCrfRecLen = Len(tmCrf)
                    slActiveDate = Format$(gNow(), "m/d/yy")
                    slActiveDate = gObtainStartStd(slActiveDate)    'Get tuesday of the current week
                    hmMcf = CBtrvTable(ONEHANDLE)
                    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        imMcfRecLen = Len(tmMcf)
                        hmCif = CBtrvTable(TWOHANDLES)
                        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                        If ilRet = BTRV_ERR_NONE Then
                            hmCnf = CBtrvTable(ONEHANDLE)
                            ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrGetFirst(hmMcf, tmMcf, imMcfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
                                Do While ilRet = BTRV_ERR_NONE
                                    ilAdd = True
                                    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
                                        If tmMcf.iVefCode > 0 Then
                                            ilVef = gBinarySearchVef(tmMcf.iVefCode)
                                            If ilVef = -1 Then
                                                ilAdd = False
                                            End If
                                        End If
                                    End If
                                    '12/8/12: Check all media codes
                                    'If (ilAdd) And (tmMcf.sReuse = "Y") Then
                                    If (ilAdd) Then
                                        ReDim tmCifInfo(0 To 0) As CIFINFO
                                        tmCifInfo(0).lFirst = -1
                                        ReDim tmCrfDates(0 To 0) As CRFDATES
                                        tmCrfDates(0).lNext = -1
                                        lmChgCount = 0
                                        lmTotalCount = 0
                                        gBuildCifArray hmCif, tmMcf.iCode, tmCifInfo()
                                        gBuildRotDates hmCrf, hmCnf, slActiveDate, tmCifInfo(), tmCrfDates()
                                        gCreateCufDates hmCif, hmCuf, tmCifInfo(), tmCrfDates()
                                        gUpdateCifDates hmCif, tmCifInfo(), tmCrfDates(), lmChgCount, lmTotalCount, False
                                    End If
                                    ilRet = btrGetNext(hmMcf, tmMcf, imMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
                                Loop
                                If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) <> MEDIACODEBYVEH Then
                                    If (tgSpf.sUseCartNo = "N") Or (tgSpf.sUseCartNo = "B") Then
                                        ReDim tmCifInfo(0 To 0) As CIFINFO
                                        tmCifInfo(0).lFirst = -1
                                        ReDim tmCrfDates(0 To 0) As CRFDATES
                                        tmCrfDates(0).lNext = -1
                                        lmChgCount = 0
                                        lmTotalCount = 0
                                        gBuildCifArray hmCif, 0, tmCifInfo()
                                        gBuildRotDates hmCrf, hmCnf, slActiveDate, tmCifInfo(), tmCrfDates()
                                        gCreateCufDates hmCif, hmCuf, tmCifInfo(), tmCrfDates()
                                        gUpdateCifDates hmCif, tmCifInfo(), tmCrfDates(), lmChgCount, lmTotalCount, False
                                    End If
                                End If
                            End If
                            On Error Resume Next
                            btrDestroy hmCnf
                        End If
                        On Error Resume Next
                        btrDestroy hmCif
                    End If
                    On Error Resume Next
                    btrDestroy hmMcf
                End If
            End If

        End If
    End If
    On Error Resume Next
    btrDestroy hmCuf
    btrDestroy hmCrf
    mCheckCopyUsage = True
    Exit Function
End Function

Private Function mCheckSeason() As Integer
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    
    imGhfRecLen = Len(tmGhf)
    imGsfRecLen = Len(tmGsf)
    imVffRecLen = Len(tmVff)
    hmGhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        If btrRecords(hmGhf) > 0 Then
            hmVff = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                hmGsf = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrGetFirst(hmGhf, tmGhf, imGhfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        If Trim$(tmGhf.sSeasonName) = "" Then
                            Do
                                llStartDate = 999999
                                llEndDate = 0
                                tmGsfSrchKey1.lghfcode = tmGhf.lCode
                                tmGsfSrchKey1.iGameNo = 0
                                ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                                Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.lghfcode = tmGhf.lCode)
                                    '12/19/12: Include cancel games to avoid date conflicts
                                    'If tmGsf.sGameStatus <> "C" Then
                                        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                        If llDate < llStartDate Then
                                            llStartDate = llDate
                                        End If
                                        If llDate > llEndDate Then
                                            llEndDate = llDate
                                        End If
                                    'End If
                                    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                If llStartDate <> 999999 Then
                                    gPackDateLong llStartDate, tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1)
                                    gPackDateLong llEndDate, tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1)
                                    If tmGhf.iSeasonStartDate(1) = tmGhf.iSeasonEndDate(1) Then
                                        tmGhf.sSeasonName = Trim$(str(tmGhf.iSeasonStartDate(1)))
                                    Else
                                        tmGhf.sSeasonName = Trim$(str(tmGhf.iSeasonStartDate(1))) & "-" & Trim$(str(tmGhf.iSeasonEndDate(1)))
                                    End If
                                    ilRet = btrUpdate(hmGhf, tmGhf, imGhfRecLen)
                                    tmVffSrchKey1.iCode = tmGhf.iVefCode
                                    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmVff.lSeasonGhfCode = tmGhf.lCode
                                        ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
                                    End If
                                End If
                                ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            Loop Until ilRet <> BTRV_ERR_NONE
                        End If
                    End If
                End If
                btrDestroy hmGsf
            End If
            btrDestroy hmVff
        End If
    End If
    btrDestroy hmGhf
    mCheckSeason = True
    Exit Function
    
End Function



Private Sub mInitMonitorMenu()
    Dim ilTask As Integer
    Dim ilMenuIndex As Integer
    Dim ilRet As Integer
    
    ' Get a handle to the top level menu
    lmDate1970 = gDateValue("1/1/1970")
    smZone = Left$(gGetLocalTZName(), 1)
    hmMonitor = GetMenu(Traffic.HWnd)
    gInitTaskInfo
    ilMenuIndex = 5
    If mMonitoringAllowed() Then
        For ilTask = 0 To UBound(tgTaskInfo) Step 1
            tmTmfSrchKey1.sTaskCode = Trim$(tgTaskInfo(ilTask).sTaskCode)
            ilRet = btrGetEqual(hgTmf, tmTmf, imTmfRecLen, tmTmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If (tmTmf.sService <> "N") And (Trim$(tmTmf.sService) <> "") Then
                    gUnpackDateLong tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), tgTaskInfo(ilTask).lRunningDate
                    gUnpackTimeLong tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), True, tgTaskInfo(ilTask).lRunningTime
                    mnuMonitor(ilTask).Visible = True
                    tgTaskInfo(ilTask).iMenuIndex = ilMenuIndex
                    tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                    tgTaskInfo(ilTask).lElapsedTime = 0
                    mSetMonitorInfo ilTask
                    ilMenuIndex = ilMenuIndex + 1
                End If
            End If
        Next ilTask
        lgDateMonitorChecked = gDateValue(Format(Now, "m/d/yy"))
        lgTimeMonitorChecked = gTimeToLong(Format(Now, "h:mm:ssAM/PM"), False)
        tmcMonitor.Interval = 100 * MONITORTIMEINTERVAL
        tmcMonitor.Enabled = True
    End If
End Sub

Private Sub mSetMonitorColor(llBgColor As Long)
    IconTraf!pbcMonitor.BackColor = llBgColor
    If (llBgColor = DARKGREEN) Or (llBgColor = vbRed) Then
        IconTraf!pbcMonitor.ForeColor = vbWhite
    Else
        IconTraf!pbcMonitor.ForeColor = vbBlack
    End If
End Sub

Private Sub mShowMonitor(ilMenuNumber As Integer)
    Dim ilRet As Integer
    Dim hlMonitorBitmap As Long
    Dim llMenuID As Long

    hlMonitorBitmap = mCopyPictureImage(IconTraf!pbcMonitor)
    
    llMenuID = GetMenuItemID(hmMonitor, ilMenuNumber)

    'If hmBitmap = 0 Then hmBitmap = CopyPictureImage(IconTraf!pbcAlert)
    ' And replace it with a bitmap.
    ilRet = ModifyMenuBynum(hmMonitor, ilMenuNumber, MF_BITMAP Or MF_BYPOSITION Or MF_ENABLED, llMenuID, hlMonitorBitmap)
End Sub

Private Sub mSetMonitorInfo(ilTask As Integer)
    mSetMonitorColor tgTaskInfo(ilTask).lColor
    IconTraf!pbcMonitor.CurrentX = 15
    IconTraf!pbcMonitor.CurrentY = 0
    IconTraf!pbcMonitor.Print Trim$(tgTaskInfo(ilTask).sTaskCode)
    If mnuAlert.Visible = False Then
        mShowMonitor tgTaskInfo(ilTask).iMenuIndex - 1
    Else
        mShowMonitor tgTaskInfo(ilTask).iMenuIndex
    End If
    mSetSubmenuStatus ilTask
End Sub

Private Sub mSetSubmenuStatus(ilTask As Integer)
    Select Case Trim$(tgTaskInfo(ilTask).sTaskCode)
        Case "CSS"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCSSStatus.Visible = True
            Else
                mnuCSSStatus.Visible = False
            End If
        Case "SSB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSSBStatus.Visible = True
            Else
                mnuSSBStatus.Visible = False
            End If
        Case "AEQ"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuAEQStatus.Visible = True
            Else
                mnuAEQStatus.Visible = False
            End If
        Case "ASI"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuASIStatus.Visible = True
            Else
                mnuASIStatus.Visible = False
            End If
        Case "ARQ"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuARQStatus.Visible = True
            Else
                mnuARQStatus.Visible = False
            End If
        Case "ASG"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuASGStatus.Visible = True
            Else
                mnuASGStatus.Visible = False
            End If
    
        Case "SC"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSCStatus.Visible = True
            Else
                mnuSCStatus.Visible = False
            End If
        Case "CE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCEStatus.Visible = True
            Else
                mnuCEStatus.Visible = False
            End If
        Case "SFE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSFEStatus.Visible = True
            Else
                mnuSFEStatus.Visible = False
            End If
        Case "ME"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuMEStatus.Visible = True
            Else
                mnuMEStatus.Visible = False
            End If
        Case "EPE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuEPEStatus.Visible = True
            Else
                mnuEPEStatus.Visible = False
            End If
        Case "ERE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuEREStatus.Visible = True
            Else
                mnuEREStatus.Visible = False
            End If
        Case "GPE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuGPEStatus.Visible = True
            Else
                mnuGPEStatus.Visible = False
            End If
        Case "BD"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuBDStatus.Visible = True
            Else
                mnuBDStatus.Visible = False
            End If
        Case "AMB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuAMBStatus.Visible = True
            Else
                mnuAMBStatus.Visible = False
            End If
        '7967
        Case "WVM"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuWviStatus.Visible = True
            Else
                mnuWviStatus.Visible = False
            End If
        'D.S. 03/21/18
        Case "PB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuPbStatus.Visible = True
            Else
                mnuPbStatus.Visible = False
            End If
        Case "TE"                           'Tableau interface
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnurestatus.Visible = True
            Else
                mnurestatus.Visible = False
            End If
        Case "RE"                           'RAB interface 1-31-20
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnurestatus.Visible = True
            Else
                mnurestatus.Visible = False
            End If
        Case "CRE"                           'TTP 9992 - Custom Rev Export
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnucrestatus.Visible = True
            Else
                mnucrestatus.Visible = False
            End If
    End Select
End Sub
Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llNow As Long
    
    On Error Resume Next
    
    If Not igBasic10Showing Then
        Exit Sub
    End If
    If imLastHourGGChecked = Hour(Now) Then
        Exit Sub
    End If
    imLastHourGGChecked = Hour(Now)
    
    '3/1/16: Don't set ggFlag as follows
    If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Or (igDemo) Then
        Exit Sub
    End If
    If bgInternalGuide Then
        Exit Sub
    End If
    
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
    llNow = gDateValue(Format$(gNow(), "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
            igWinStatus(ilLoop) = 0
        Next ilLoop
        tgSpf.sAExcl = "N"
        tgSpf.sAEDIC = "N"
        tgSpf.sAEDII = "N"
        tgSpf.sUseCartNo = "N"
        tgSpf.sGUsePropSys = "N"
        tgSpf.sGUseAffFeed = "N"
        igGGFlag = 0
        Basic10!pbcHor.Enabled = False
        For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
            Basic10!cmcTask(ilLoop).Enabled = False
            Basic10!lacTask(ilLoop).Enabled = False
            Basic10!lacTask(ilLoop).ForeColor = vbBlack
        Next ilLoop
    End If
    gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub
Private Sub tmcWebConnectIssue_Timer()
    '7967  in past, checking with web did 20 retries.  Now only try once and come here to retry if that fails
    Static ilMinutes As Integer
    
    ilMinutes = ilMinutes + 1
    '8133 10 minute span
    If ilMinutes = 10 Then
        ilMinutes = 0
        If Not bgVendorToWebAllowed Then
            If gVendorToWebAllowed() Then
                    'force task monitor timer to run test to web
                    igWVImportElapsed = 0
                    dgWvImportLast = 0
                    ' set by gVendorToWebAllowed bgVendorToWebAllowed = True
                    tmcWebConnectIssue.Enabled = False
            Else
                gLogMsg "VendorToWeb-retry of gVendorToWebAllowed failed", "TrafficErrors.Txt", False
            End If
        End If
    End If
End Sub
Private Function mMonitoringAllowed() As Boolean
    'also sets emMonitoringAllowed for testing in timer, and sets the monitor id value to be output in email
    Dim blRet As Boolean
    
    blRet = True
    If igTestSystem Then
        gLoadOption "TestLocations", "MonitorID", smMonitorID
    Else
        gLoadOption "Locations", "MonitorID", smMonitorID
    End If
    emMonitoringAllowed = MonitorNormal
    If Len(Trim$(smMonitorID)) > 0 Then
        Select Case UCase(smMonitorID)
            Case "OFF", "ICONSOFF"
                emMonitoringAllowed = MonitorOff
                blRet = False
            Case "EMAILOFF"
                emMonitoringAllowed = MonitorNoEmail
        End Select
        smMonitorID = " MonitorID: " & smMonitorID
    Else
        smMonitorID = " MonitorID is not set"
    End If
    mMonitoringAllowed = blRet
End Function
Private Sub mTempSendServiceEmail(slSubject As String, slBody As String)

    If slSubject = "" Then
        slSubject = "Automated message from a client"
    End If
    Set ogEmailer = New CEmail

    With ogEmailer
        .FromAddress = "AClient@Counterpoint.net"
        .FromName = Trim$(tgSpf.sGClient)
        .AddTOAddress "danmichaelson@counterpoint.net", "Dan"
        .Subject = slSubject
        .Message = slBody
        '10564
        '.SetHost "smtpauth.hosting.earthlink.net", 587, "emailsender@counterpoint.net", "Csi44Sic", True
        .SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, CSITLS
        If Not .Send() Then
            gLogMsg "Email could not be sent from modClass-gSendServiceEmail: ", "AffErrorLog.Txt", False
        End If
    End With
    Set ogEmailer = Nothing
End Sub


Private Sub mCheckV81Flags()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVpfIndex As Integer
    Dim slMedium As String * 1
    Dim slSQLQuery As String
    Dim llRet As Long
    Dim blConvert As Boolean
    Dim slFeatures8 As String * 1
    Dim saf_rst As ADODB.Recordset
    Dim vpf_rst As ADODB.Recordset
    
    blConvert = False
    hmPvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imPvfRecLen = Len(tmPvf) 'btrRecordLength(hmAdf)    'Get Cff size

    tmPvfSrchKey.lCode = tgMVef(ilVef).lPvfCode
    ilRet = btrGetFirst(hmPvf, tmPvf, imPvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    Do While ilRet = BTRV_ERR_NONE
        If Trim$(tmPvf.sType) = "" Then
            slMedium = ""
            For ilVef = LBound(tmPvf.iVefCode) To UBound(tmPvf.iVefCode) Step 1
                If tmPvf.iVefCode(ilVef) > 0 Then
                    ilVpfIndex = gBinarySearchVpf(tmPvf.iVefCode(ilVef))
                    If ilVpfIndex > 0 Then
                        If slMedium = " " Then
                            If tgVpf(ilVpfIndex).sGMedium <> "P" Then
                                slMedium = "A"
                            Else
                                slMedium = "P"
                            End If
                        Else
                            If tgVpf(ilVpfIndex).sGMedium <> "P" Then
                                If slMedium <> "A" Then
                                    slMedium = "M"
                                    Exit For
                                End If
                            Else
                                If slMedium <> "P" Then
                                    slMedium = "M"
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilVef
            If slMedium = " " Then
                slMedium = "A"
            End If
            DoEvents
            blConvert = True
            tmPvf.sType = slMedium
            ilRet = btrUpdate(hmPvf, tmPvf, imPvfRecLen)
            ilRet = btrGetNext(hmPvf, tmPvf, imPvfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Else
            Exit Do
        End If
    Loop
    If blConvert Then
        slSQLQuery = "Update chf_Contract_Header Set chfPodSpotDefined = 'N'"
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        slSQLQuery = "Update chf_Contract_Header Set chfPodSpotDefined = 'Y'"
        slSQLQuery = slSQLQuery & " Where chfCode in (Select Distinct clfChfCode From clf_Contract_Line Left Outer Join vpf_Vehicle_Options on clfVefCode = vpfvefKCode Where vpfGMedium = 'P')"
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    End If
    If blConvert Then
        slSQLQuery = "Select safFeatures8 From saf_Schd_Attributes Where safVefCode = 0"
        Set saf_rst = gSQLSelectCall(slSQLQuery)
        If Not saf_rst.EOF Then
            slFeatures8 = saf_rst!safFeatures8
            If Trim$(slFeatures8) = "" Then
                slFeatures8 = Chr(0)
            End If
            slSQLQuery = "Select count(1) From vpf_Vehicle_Options Where ((vpfGMedium <> 'P') and (vpfGMedium <> 'M'))"
            Set vpf_rst = gSQLSelectCall(slSQLQuery)
            If Not vpf_rst Then
                If vpf_rst(0).Value > 0 Then
                    slFeatures8 = Chr$(Asc(slFeatures8) Or PODAIRTIME)
                End If
            End If
            slSQLQuery = "Select count(1) From vpf_Vehicle_Options Where vpfGMedium = 'P'"
            Set vpf_rst = gSQLSelectCall(slSQLQuery)
            If Not vpf_rst Then
                If vpf_rst(0).Value > 0 Then
                    slFeatures8 = Chr$(Asc(slFeatures8) Or PODSPOTS)
                End If
            End If
            slSQLQuery = "Select count(1) From vpf_Vehicle_Options Where vpfGMedium = 'M'"
            Set vpf_rst = gSQLSelectCall(slSQLQuery)
            If Not vpf_rst Then
                If vpf_rst(0).Value > 0 Then
                    slFeatures8 = Chr$(Asc(slFeatures8) Or PODADSERVER)
                End If
            End If
            slSQLQuery = "Update saf_Schd_Attributes Set safFeatures8 = '" & slFeatures8 & "' Where safVefCode = 0"
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        End If
        On Error Resume Next
        tgSaf(0).sFeatures8 = slFeatures8
        saf_rst.Close
        vpf_rst.Close
    End If
    On Error Resume Next
    ilRet = btrClose(hmPvf)
    btrDestroy hmPvf
End Sub


