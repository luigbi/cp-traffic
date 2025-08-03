VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Traffic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2475
   ClientLeft      =   1785
   ClientTop       =   2355
   ClientWidth     =   7365
   ControlBox      =   0   'False
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
   ForeColor       =   &H80000008&
   Icon            =   "TRAFFIC2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   7365
   Visible         =   0   'False
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
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7365
      Begin VB.PictureBox pbc256SignOn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5940
         Picture         =   "TRAFFIC2.frx":08CA
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   525
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
         Left            =   2490
         ScaleHeight     =   225
         ScaleWidth      =   1005
         TabIndex        =   7
         Top             =   120
         Width           =   1035
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
         Left            =   5925
         TabIndex        =   6
         Top             =   825
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox mkcTempFaxNo 
         Height          =   210
         Left            =   2505
         TabIndex        =   5
         Top             =   1725
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
         Left            =   1665
         TabIndex        =   4
         Top             =   1755
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
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   735
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cdcSetup 
         Left            =   6690
         Top             =   75
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   0
         MaxFileSize     =   256
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
         Left            =   1485
         ScaleHeight     =   675
         ScaleWidth      =   2985
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   750
         Width           =   2985
      End
      Begin MSMask.MaskEdBox mkcPhone 
         Height          =   210
         Left            =   1605
         TabIndex        =   1
         Top             =   1560
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
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   795
         Left            =   6885
         TabIndex        =   9
         Top             =   390
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1402
         _Version        =   393217
         TextRTF         =   $"TRAFFIC2.frx":4FEE
      End
      Begin VB.Image imcCSLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6540
         Picture         =   "TRAFFIC2.frx":5068
         Top             =   1290
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imcHelp 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5415
         Picture         =   "TRAFFIC2.frx":5932
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image imcCSILogo 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5865
         Top             =   1365
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "Traffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    'pbcLogo.Picture = LoadPicture(sgLogoPath & "CustLogo.Bmp")
    On Error Resume Next
    pbcSymb.Picture = LoadPicture(sgLogoPath & "CustSymb.Bmp")
    mkcPhone.Text = "(___) ___-____ Ext(____)"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gEraseGlobalVar True
End Sub
Private Sub imcHelp_Click()
    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    Traffic!cdcSetup.HelpCommand = cdlHelpIndex 'cdlHelpIndex
    Traffic!cdcSetup.Action = 6
End Sub
