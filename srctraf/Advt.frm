VERSION 5.00
Begin VB.Form Advt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10905
   ClientLeft      =   49410
   ClientTop       =   3495
   ClientWidth     =   19110
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10905
   ScaleWidth      =   19110
   Begin VB.TextBox edcMegaphoneAdvID 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3825
      MaxLength       =   70
      TabIndex        =   93
      Top             =   8235
      Width           =   1335
   End
   Begin VB.PictureBox pbcSuppressNet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9645
      ScaleHeight     =   210
      ScaleWidth      =   1350
      TabIndex        =   70
      Top             =   3390
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox edcSuppressNet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   9630
      MaxLength       =   6
      TabIndex        =   71
      Top             =   3690
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox edcCRMID 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5520
      TabIndex        =   73
      Top             =   8250
      Width           =   1335
   End
   Begin VB.TextBox edcDirectRefID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   7185
      MaxLength       =   36
      TabIndex        =   72
      Top             =   6600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox edcAddrID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   62
      Top             =   6855
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox lbcName 
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
      Height          =   240
      ItemData        =   "Advt.frx":0000
      Left            =   9300
      List            =   "Advt.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.PictureBox pbcPolitical 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12105
      ScaleHeight     =   210
      ScaleWidth      =   795
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lbcTerms 
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
      Height          =   240
      ItemData        =   "Advt.frx":0004
      Left            =   15585
      List            =   "Advt.frx":0006
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   6090
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcRepInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12240
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5025
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcBonusOnInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12345
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3795
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcRepMG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12690
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3345
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5655
      TabIndex        =   1
      Top             =   45
      Width           =   3180
   End
   Begin VB.PictureBox pbcRateOnInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12315
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4230
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcPackage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14280
      ScaleHeight     =   210
      ScaleWidth      =   1005
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ListBox lbcCreditApproval 
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
      Height          =   240
      Left            =   9375
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   9180
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox edcRating 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   7110
      MaxLength       =   7
      TabIndex        =   49
      Top             =   6225
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox edcRefId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   7095
      MaxLength       =   36
      TabIndex        =   50
      Top             =   6945
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14400
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3765
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcPayable 
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
      Height          =   240
      Left            =   9420
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9585
      Visible         =   0   'False
      Width           =   5730
   End
   Begin VB.ListBox lbcBuyer 
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
      Height          =   240
      Left            =   9345
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5715
      Visible         =   0   'False
      Width           =   5730
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   14280
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   4725
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   15
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   4740
      Width           =   75
   End
   Begin VB.ListBox lbcLkBox 
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
      Height          =   240
      Left            =   13185
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcEDI 
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
      Height          =   240
      Index           =   0
      Left            =   13230
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcProd 
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
      Height          =   240
      Left            =   14580
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcInvSort 
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
      Height          =   240
      Left            =   12375
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5370
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcTax 
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
      Height          =   240
      Left            =   13050
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8190
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.ListBox lbcEDI 
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
      Height          =   240
      Index           =   1
      Left            =   13170
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9255
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcAgency 
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
      Height          =   240
      Left            =   9390
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5385
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcSPerson 
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
      Height          =   240
      Left            =   9210
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6705
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcPaymRating 
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
      Height          =   240
      Left            =   9345
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8550
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcExcl 
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
      Height          =   240
      Index           =   0
      Left            =   9390
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8250
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcExcl 
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
      Height          =   240
      Index           =   1
      Left            =   9300
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcCreditRestr 
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
      Height          =   240
      Left            =   9345
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   7155
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcComp 
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
      Height          =   240
      Index           =   0
      Left            =   9390
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7515
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcComp 
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
      Height          =   240
      Index           =   1
      Left            =   9390
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7890
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcBillTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   1
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   55
      Top             =   2505
      Width           =   60
   End
   Begin VB.PictureBox pbcBillTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   60
      ScaleHeight     =   120
      ScaleWidth      =   90
      TabIndex        =   54
      Top             =   2430
      Width           =   90
   End
   Begin VB.PictureBox pbcDmPriceType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14145
      ScaleHeight     =   210
      ScaleWidth      =   2850
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3375
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.CommandButton cmcCEDropDown 
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
      Left            =   8625
      Picture         =   "Advt.frx":0008
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3345
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcCEDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6870
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   3150
      Picture         =   "Advt.frx":0102
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4050
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6390
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDmDropDown 
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
      Left            =   8670
      Picture         =   "Advt.frx":01FC
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDmDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6870
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.ListBox lbcDemo 
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
      Height          =   240
      Index           =   1
      Left            =   12930
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6750
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox lbcDemo 
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
      Height          =   240
      Index           =   2
      Left            =   12930
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox lbcDemo 
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
      Height          =   240
      Index           =   3
      Left            =   12975
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   7425
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox lbcDemo 
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
      Height          =   240
      Index           =   0
      Left            =   12975
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7815
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox pbcDm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   8850
      Picture         =   "Advt.frx":02F6
      ScaleHeight     =   1770
      ScaleWidth      =   2910
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox pbcCE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   12030
      Picture         =   "Advt.frx":3190
      ScaleHeight     =   720
      ScaleWidth      =   4755
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.PictureBox plcCE 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   11910
      ScaleHeight     =   1170
      ScaleWidth      =   5100
      TabIndex        =   19
      Top             =   2025
      Visible         =   0   'False
      Width           =   5100
      Begin VB.PictureBox pbcCESTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   20
         Top             =   225
         Width           =   60
      End
      Begin VB.PictureBox pbcCETab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   30
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   27
         Top             =   840
         Width           =   75
      End
   End
   Begin VB.PictureBox pbcNotDirect 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   720
      Left            =   8820
      Picture         =   "Advt.frx":6F62
      ScaleHeight     =   720
      ScaleWidth      =   8490
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   60
      Width           =   8490
   End
   Begin VB.PictureBox plcNotDirect 
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   150
      ScaleHeight     =   765
      ScaleWidth      =   8565
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   9705
      Width           =   8625
   End
   Begin VB.PictureBox plcDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   8775
      ScaleHeight     =   2175
      ScaleWidth      =   3240
      TabIndex        =   28
      Top             =   765
      Visible         =   0   'False
      Width           =   3240
      Begin VB.PictureBox pbcDmTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   75
         ScaleHeight     =   90
         ScaleWidth      =   75
         TabIndex        =   38
         Top             =   2025
         Width           =   75
      End
      Begin VB.PictureBox pbcDmSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   29
         Top             =   225
         Width           =   60
      End
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   8655
      Top             =   5085
   End
   Begin VB.PictureBox pbcISCI 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14460
      ScaleHeight     =   210
      ScaleWidth      =   765
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox edcAgyCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   1365
      MaxLength       =   10
      TabIndex        =   17
      Top             =   8850
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox edcCreditLimit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   5205
      MaxLength       =   10
      TabIndex        =   42
      Top             =   6420
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   150
      MaxLength       =   40
      TabIndex        =   58
      Top             =   7875
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   195
      MaxLength       =   40
      TabIndex        =   57
      Top             =   7170
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   3435
      MaxLength       =   40
      TabIndex        =   61
      Top             =   9300
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   3510
      MaxLength       =   40
      TabIndex        =   60
      Top             =   8925
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   3465
      MaxLength       =   40
      TabIndex        =   59
      Top             =   8565
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcStnCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   1695
      MaxLength       =   10
      TabIndex        =   18
      Top             =   8475
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcRepCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   105
      MaxLength       =   10
      TabIndex        =   16
      Top             =   8925
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   150
      MaxLength       =   40
      TabIndex        =   56
      Top             =   6825
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcAbbr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   135
      MaxLength       =   7
      TabIndex        =   8
      Top             =   8460
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   150
      MaxLength       =   30
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton cmcSplitCue 
      Appearance      =   0  'Flat
      Caption         =   "S&plit Cue"
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
      Left            =   7320
      TabIndex        =   81
      Top             =   5805
      Width           =   1050
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge"
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
      Left            =   6180
      TabIndex        =   80
      Top             =   5805
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      Left            =   5055
      TabIndex        =   79
      Top             =   5805
      Width           =   1050
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
      Left            =   3930
      TabIndex        =   78
      Top             =   5805
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   2805
      TabIndex        =   77
      Top             =   5805
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
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
      Left            =   1680
      TabIndex        =   76
      Top             =   5805
      Width           =   1050
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
      Left            =   555
      TabIndex        =   75
      Top             =   5805
      Width           =   1050
   End
   Begin VB.PictureBox pbcAdvt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   1425
      Left            =   270
      Picture         =   "Advt.frx":DCAC
      ScaleHeight     =   1425
      ScaleWidth      =   8490
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   615
      Width           =   8490
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1095
      ScaleHeight     =   240
      ScaleWidth      =   450
      TabIndex        =   2
      Top             =   6315
      Width           =   450
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   420
      ScaleHeight     =   375
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   6240
      Width           =   435
   End
   Begin VB.PictureBox plcAdvt 
      ForeColor       =   &H00000000&
      Height          =   1560
      Left            =   210
      ScaleHeight     =   1500
      ScaleWidth      =   8550
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   8610
   End
   Begin VB.PictureBox pbcDirect 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   2820
      Left            =   330
      Picture         =   "Advt.frx":35546
      ScaleHeight     =   2820
      ScaleWidth      =   8490
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Width           =   8490
   End
   Begin VB.PictureBox plcDirect 
      ForeColor       =   &H00000000&
      Height          =   3180
      Left            =   210
      ScaleHeight     =   3120
      ScaleWidth      =   8565
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8625
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9615
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   4395
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9375
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame frcBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2865
      TabIndex        =   89
      Top             =   2145
      Width           =   3690
      Begin VB.OptionButton rbcBill 
         Caption         =   "Direct (Advertiser)"
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
         Index           =   1
         Left            =   1545
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   15
         Width           =   1965
      End
      Begin VB.OptionButton rbcBill 
         Caption         =   "Agency"
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
         Index           =   0
         Left            =   420
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   15
         Width           =   1020
      End
      Begin VB.Label lacBill 
         Caption         =   "Bill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   92
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.Label plcScreen 
      Caption         =   "Advertiser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1425
   End
   Begin VB.Label lacCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   7935
      TabIndex        =   87
      Top             =   390
      Width           =   870
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   45
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Advt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Advt.frm on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Advt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Advertiser input screen code
Option Explicit
Option Compare Text

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9

'Advertiser Field Areas
Dim tmCtrls(0 To 46)  As FIELDAREA  'TTP 10397 - Direct Advertiser list screen - third line of contract address appears in billing address field, billing address field 2nd and 3rd line can't be edited
Dim imLBCtrls As Integer
Dim tmDmCtrls(0 To 9) As FIELDAREA
Dim imLBDmCtrls As Integer
Dim tmCECtrls(0 To 4) As FIELDAREA
Dim imLBCECtrls As Integer
Dim tmDTCtrls(0 To 11)  As FIELDAREA
Dim tmNDTCtrls(0 To 11)  As FIELDAREA
Dim imFirstActivate As Integer
Dim imPrtStyle As Integer           '0=Wide; 1=Narrow
Dim imRateOnInv As Integer          '0=Yes; 1= No
Dim imISCI As Integer               '0=Yes; 1= No
Dim imRepMG As Integer              '0=Yes; 1= No
Dim imRepInv As Integer             'Rep Invoice generated: 0=Internal; 1=External
Dim imPolitical As Integer           '0=Yes; 1= No
Dim imBonusOnInv As Integer         '0=Yes; 1= No
Dim imState As Integer              '0=Active; 1= Dormant
Dim imPackage As Integer            '0=Daypart; 1=Time
Dim imMaxNoCtrls As Integer         'Set to Invoice Sort or Sales Tax
Dim imEdcChgMode As Integer
Dim imBoxNo As Integer              'Current Agency Box
Dim imDMBoxNo As Integer
Dim imCEBoxNo As Integer
Dim tmEDICode() As SORTCODE
Dim smEDICodeTag As String
Dim tmInvSortCode() As SORTCODE
Dim smInvSortCodeTag As String
Dim tmTermsCode() As SORTCODE
Dim smTermsCodeTag As String
Dim tmLkBoxCode() As SORTCODE
Dim smLkBoxCodeTag As String
Dim tmTaxSortCode() As SORTCODE
Dim smTaxSortCodeTag As String
Dim hmAdf As Integer                'Advertiser file handle
Dim tmAdf As ADF                    'ADF record image
Dim tmAdfSrchKey As INTKEY0         'ADF key record image
Dim imAdfRecLen As Integer          'ADF record length
Dim tmAdfx As ADFX 'L.Bianchi 05/31/2021
Dim hmPrf As Integer                'Product file handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer
Dim tmPrfSrchKey As PRFKEY1
Dim tmPrfSrchKey0 As LONGKEY0
Dim hmPnf As Integer                'Product file handle
Dim tmBPnf As PNF
Dim tmPPnf As PNF
Dim imPnfRecLen As Integer
Dim tmPnfSrchKey As INTKEY0
Dim tmPnfSrchKey2 As PNFKEY2
Dim imNewPnfCode() As Integer
Dim hmPDF As Integer                'Media code file handle
Dim imPdfRecLen As Integer          'McF record length
Dim tmPdfSrchKey1 As INTKEY0
Dim tmPdfSrchKey2 As INTKEY0
Dim tmPdf() As PDF
Dim bmPDFEMailChgd As Boolean
Dim hmAxf As Integer                'Product file handle
Dim tmAxf As AXF
Dim imAxfRecLen As Integer
Dim tmAxfSrchKey0 As LONGKEY0
Dim tmAxfSrchKey1 As INTKEY0
Dim bmAxfChg As Boolean
Dim hmSaf As Integer
Dim hmAgf As Integer

Dim imChgMode As Integer            'Change mode status (so change not entered when in change)
Dim imBSMode As Integer             'Backspace flag
Dim imSelectedIndex As Integer      'Index of selected record (0 if new)
Dim imTerminate As Integer          'True = terminating task, False= OK
Dim smProduct As String             'Product name, saved to determine if changed
Dim smBuyer As String               'Buyer name, saved to determine if changed
Dim smPayable As String             'Payable name, saved to determine if changed
Dim smOrigBuyer As String           'Buyer name, saved to determine if changed
Dim smOrigPayable As String         'Payable name, saved to determine if changed
Dim smSPerson As String             'Salesperson name, saved to determine if changed
Dim smAgency As String              'Agency name, saved to determine if changed
Dim smComp(0 To 1) As String        'Competitive name, saved to determine if changed
Dim smExcl(0 To 1) As String        'Exclusion name, saved to determine if changed
Dim smSvDemo(0 To 3) As String
Dim smSvTarget(0 To 3) As String    'Demo targets, saved to determine if changed
Dim smTarget(0 To 3) As String      'Demo targets
Dim smBkoutPool As String
Dim smLkBox As String               'Lock box, saved to determine if changed
Dim smInvSort As String
Dim smTerms As String
Dim smTax As String
Dim smEDIC As String                'EDI service for contracts, saved to determine if changed
Dim smEDII As String                'EDI service for invoices, saved to determine if changed
Dim imPriceType As Integer          '0=N/a; 1= CPM; 2=CPP
Dim imComboBoxIndex As Integer
Dim imUpdateAllowed As Integer      'User can update records
Dim imCreditRestrFirst As Integer   'First time at field-set default if required
Dim imPaymRatingFirst As Integer    'First time at field-set default if required
'Dim imPrtStyleFirst As Integer     'First time at field-set default if required
Dim imScriptFirst As Integer        'First time at field-set default if required
Dim imPriceTypeFirst As Integer     'First time at field-set default if required
Dim imCombo As Integer              'True=Combo salesperson; False=Standard allow salesperson
Dim imTaxFirst As Integer           'First time at field-set default if required
Dim imFirstFocus As Integer
Dim smPct90 As String
Dim smCurrAR As String
Dim smUnbilled As String
Dim smHiCredit As String
Dim smTotalGross As String
Dim smDateEntrd As String
Dim smNSFChks As String
Dim smDateLstInv As String
Dim smDateLstPaym As String
Dim smAvgToPay As String
Dim smLstToPay As String
Dim smNoInvPd As String
Dim imLbcMouseDown As Integer       'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer    'True = Processing arrow key- retain current list box visibly
                                    'False= Make list box invisible
Dim imDirProcess As Integer
Dim imTabDirection As Integer       '0=left to right (Tab); -1=right to left (Shift tab)
Dim imTaxDefined As Integer
Dim imBillTabDir As Integer         '0=Index 0 to pbcTab, 1=Index 0 to pbcSTab
Dim imPopReqd As Integer            'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imShortForm As Integer          'True=Only process min number of fields
Dim imChgSaveFlag As Integer        'Indicates if any changed saved
Dim lmKeyState As Long
Dim imSuppressNet As Integer        '0=No (default), 1=Yes  Suppress Net Amount for Trade Invoices  ' TTP 10622 - 2023-03-08 JJB

Dim smMegaphoneAdvID_Original As String
Dim imAvfCode As Integer

Public smAVIndicatorID As String
Public smXDSCue As String

Dim fmAdjFactorW As Single          'Width adjustment factor
Dim fmAdjFactorH As Single          'Width adjustment factor

'Advertiser
'LINE 1
    Const NAMEINDEX = 1                 'Name control/field
    Const ABBRINDEX = 2                 'Last name control/field
    Const POLITICALINDEX = 3            'Political control/index
    Const STATEINDEX = 4                'Active/Dormant control/index
    Const PRODINDEX = 5                 'Product control/index
'LINE 2
    Const SPERSONINDEX = 6              'Salesperson control/field
    Const AGENCYINDEX = 7               'Agency control/index
    Const REPCODEINDEX = 8              'Rep advertiser Code control/field
    Const AGYCODEINDEX = 9              'Agency advertiser code
    Const STNCODEINDEX = 10             'Station Agency Code control/field
'LINE 3
    Const COMPINDEX = 11
    Const DEMOINDEX = 12
    Const BKOUTPOOLINDEX = 13
    Const CREDITAPPROVALINDEX = 14
'LINE 4
    Const CREDITRESTRINDEX = 15         'Credit restriction control/field
    Const PAYMRATINGINDEX = 17          'Payment rating control/field
    Const CREDITRATINGINDEX = 18
    Const REPINVINDEX = 19
    Const INVSORTINDEX = 20             'Invoice sort control/field
    Const RATEONINVINDEX = 21
    Const ISCIINDEX = 22                'ISCI control/field
    
    ' what are the following 2 lines?? Where are they??  - JJB
    Const REPMGINDEX = 23               'Count Rep MG
    Const BONUSONINVINDEX = 24          'Bonus spots show on Invoices
    
    Const PACKAGEINDEX = 25
    Const REFIDINDEX = 26               'L.Bianchi 05/26/2021
'LINE 5
    Const DIRECTREFIDINDEX = 27         'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    Const DIRECTINDEX = 28              'Bill direct
    Const MEGAPHONEADVID = 29           ' Megaphone Advertiser ID
    Const CRMIDINDEX = 30               ' JD 09/19/22

'Direct Advertiser

Const CADDRINDEX = 31               'Contract address control/field (3 fields)
'TTP 10397 - Direct Advertiser list screen - third line of contract address appears in billing address field, billing address field 2nd and 3rd line can't be edited
Const BADDRINDEX = 34               'Billing address control/field (3 fields)
Const ADDRIDINDEX = 37
Const BUYERINDEX = 38               'Buyer control/field
Const PAYABLEINDEX = 39             'Buyer control/field
Const LKBOXINDEX = 40               'Lock box control/field
Const EDICINDEX = 41                'EDI service for contract control/field
Const EDIIINDEX = 42                'EDI service for Invoice control/field
Const TERMSINDEX = 43
Const TAXINDEX = 44                 'Tax 1 and 2 control/field
Const SUPPRESSNETINDEX = 45         'Suppress Net   ' TTP 10622 - 2023-03-08 JJB
Const UNUSEDINDEX = 46

Const CECOMPINDEX = 1
Const CEEXCLINDEX = 3
Const DMPRICETYPEINDEX = 1
Const DMDEMOINDEX = 2   '4, 6, 8
Const DMVALUEINDEX = 3  '5, 7, 9

Const PCT90INDEX = 1
Const CURRARINDEX = 2
Const UNBILLEDINDEX = 3
Const HICREDITINDEX = 4
Const TOTALGROSSINDEX = 5
Const DATEENTRDINDEX = 6
Const NSFCHKSINDEX = 7
Const DATELSTINVINDEX = 8
Const DATELSTPAYMINDEX = 9
Const AVGTOPAYINDEX = 10
Const LSTTOPAYINDEX = 11

Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim tlPnf As PNF
    If imChgMode Then  'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY, True) Then
            GoTo cbcSelectErr
        End If
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            lacCode.Caption = str$(tmAdf.iCode)
        Else
            lacCode.Caption = ""
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
        lacCode.Caption = ""
    End If
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    pbcAdvt.Cls
    pbcNotDirect.Cls
    pbcDirect.Cls
    mPaintAdvtTitle
    mPaintDirectTitle
    mPaintNotDirectTitle
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    For ilLoop = imLBCECtrls To UBound(tmCECtrls) Step 1
        tmCECtrls(ilLoop).sShow = ""
        mCESetShow ilLoop  'Set show strings
    Next ilLoop
    For ilLoop = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        mInitDmShow ilLoop  'Set show strings
    Next ilLoop
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If ilLoop <> DEMOINDEX And ilLoop <> BKOUTPOOLINDEX Then 'Wait until end for DEMOINDEX to aviod painting before all values set
            tmCtrls(ilLoop).sShow = ""
            mSetShow ilLoop  'Set show strings
        End If
    Next ilLoop
    If (smBkoutPool = "N") Or (smBkoutPool = "") Then
        If imSelectedIndex = 0 Then
            slStr = ""
        Else
            slStr = "N"
        End If
    Else
        slStr = "Y"
    End If
    
    gSetShow pbcAdvt, slStr, tmCtrls(BKOUTPOOLINDEX)
    mSetShow DEMOINDEX  'Set show strings
    'Paint not required as SetShow for DEMOINDEX contains a paint call
'        pbcAdvt_Paint
    pbcDirect_Paint
    pbcNotDirect_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
'    mSetCommands
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub edcMegaphoneAdvID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSuppressNet_Change()
' TTP 10622 - 2023-03-08 JJB
    mSetChg imBoxNo
End Sub

Private Sub edcSuppressNet_Click()
' TTP 10622 - 2023-03-08 JJB
    Call edcSuppressNet_Change
End Sub

Private Sub edcSuppressNet_GotFocus()
' TTP 10622 - 2023-03-08 JJB
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcSelect_DropDown()
    'Removing this code- when New name added, after save if drop down selected the name when blank
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igAdvtCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgAdvtName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgAdvtName    'New name
            End If
            cbcSelect_Change
            If sgAdvtName <> "" Then
                mSetCommands
                gFindMatch sgAdvtName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
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
Private Sub cmcCancel_Click()
    Dim ilLoop As Integer
    Dim tlPnf As PNF
    Dim ilRet As Integer
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    If igAdvtCallSource <> CALLNONE Then
        igAdvtCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub cmcCEDropDown_Click()
    Select Case imCEBoxNo
        Case CECOMPINDEX
            lbcComp(0).Visible = Not lbcComp(0).Visible
        Case CECOMPINDEX + 1
            lbcComp(1).Visible = Not lbcComp(1).Visible
        Case CEEXCLINDEX
            lbcExcl(0).Visible = Not lbcExcl(0).Visible
        Case CEEXCLINDEX + 1
            lbcExcl(1).Visible = Not lbcExcl(1).Visible
    End Select
    edcCEDropDown.SelStart = 0
    edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
    edcCEDropDown.SetFocus
End Sub
Private Sub cmcCEDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDmDropDown_Click()
    Select Case imDMBoxNo
        Case DMDEMOINDEX
            lbcDemo(0).Visible = Not lbcDemo(0).Visible
        Case DMDEMOINDEX + 2
            lbcDemo(1).Visible = Not lbcDemo(1).Visible
        Case DMDEMOINDEX + 4
            lbcDemo(2).Visible = Not lbcDemo(2).Visible
        Case DMDEMOINDEX + 6
            lbcDemo(3).Visible = Not lbcDemo(3).Visible
    End Select
    edcDmDropDown.SelStart = 0
    edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
    edcDmDropDown.SetFocus
End Sub
Private Sub cmcDmDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim tlPnf As PNF
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    imChgSaveFlag = False
    If igAdvtCallSource <> CALLNONE Then
        sgAdvtName = Trim$(edcName.Text) 'Save name for returning
        If rbcBill(1).Value Then
            If Trim$(edcAddrID.Text) <> "" Then
                sgAdvtName = sgAdvtName & ", " & Trim$(edcAddrID.Text) & "/Direct"
            Else
                sgAdvtName = sgAdvtName & "/Direct"
            End If
        End If
        If mSaveRecChg(False) = False Then
            sgAdvtName = "[New]"
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    'If no to save- remove any pnf created
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    If imChgSaveFlag Then
        sgAdvertiserTag = ""
        mPopulate
    End If
    If igAdvtCallSource <> CALLNONE Then
        If sgAdvtName = "[New]" Then
            igAdvtCallSource = CALLCANCELLED
        Else
            igAdvtCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To imMaxNoCtrls Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case NAMEINDEX
            lbcName.Visible = Not lbcName.Visible
        Case PRODINDEX
            lbcProd.Visible = Not lbcProd.Visible
        Case SPERSONINDEX
            lbcSPerson.Visible = Not lbcSPerson.Visible
        Case AGENCYINDEX
            lbcAgency.Visible = Not lbcAgency.Visible
        Case CREDITRESTRINDEX
            lbcCreditRestr.Visible = Not lbcCreditRestr.Visible
        Case PAYMRATINGINDEX
            lbcPaymRating.Visible = Not lbcPaymRating.Visible
        Case CREDITAPPROVALINDEX
            lbcCreditApproval.Visible = Not lbcCreditApproval.Visible
        Case INVSORTINDEX
            lbcInvSort.Visible = Not lbcInvSort.Visible
        Case BUYERINDEX
            lbcBuyer.Visible = Not lbcBuyer.Visible
        Case PAYABLEINDEX
            lbcPayable.Visible = Not lbcPayable.Visible
        Case LKBOXINDEX
            lbcLkBox.Visible = Not lbcLkBox.Visible
        Case EDICINDEX
            lbcEDI(0).Visible = Not lbcEDI(0).Visible
        Case EDIIINDEX
            lbcEDI(1).Visible = Not lbcEDI(1).Visible
        Case TERMSINDEX
            lbcTerms.Visible = Not lbcTerms.Visible
        Case TAXINDEX
            lbcTax.Visible = Not lbcTax.Visible
    End Select
    If imBoxNo = NAMEINDEX Then
        edcName.SelStart = 0
        edcName.SelLength = Len(edcName.Text)
        edcName.SetFocus
    Else
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    End If
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlPnf As PNF
    Dim rs As ADODB.Recordset
    Dim llCount As Long
    If imSelectedIndex > 0 Then
        If tgSpf.sRemoteUsers = "Y" Then
            slMsg = "Cannot erase - Remote User System in Use"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Bof.Btr", "BofAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Blackout references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Cdf.Btr", "CdfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Comments references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Chf.Btr", "ChfAdfCode")   'chfadfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Cif.Btr", "CifAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Inventory references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Crf.Btr", "CrfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Rotation references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Csf.Btr", "CsfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Script references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Odf.Btr", "OdfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a One Day Log references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Pjf.Btr", "PjfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract Projection references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Psf.Btr", "PsfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Package Spot references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Phf.Btr", "PhfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables History references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        '9/22/05- Product is delete so test is not necessary
        'ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Prf.Btr", "PrfAdfCode")
        'If ilRet Then
        '    Screen.MousePointer = vbDefault
        '    slMsg = "Cannot erase - a Product Name references name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Rvf.Btr", "RvfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Sdf.Btr", "SdfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Spot Detail references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Sif.Btr", "SifAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Short Title references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Shf.Btr", "ShfAdfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Spot History references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            ilRet = MsgBox("OK to remove " & Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID), vbOKCancel + vbQuestion, "Erase")
        Else
            ilRet = MsgBox("OK to remove " & Trim$(tmAdf.sName), vbOKCancel + vbQuestion, "Erase")
        End If
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        gGetSyncDateTime slSyncDate, slSyncTime
        slStamp = gFileDateTime(sgDBPath & "Adf.Btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE, False) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        End If
        tmPrfSrchKey.iAdfCode = tmAdf.iCode
        ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = tmAdf.iCode)
            'tmRec = tmPrf
            'ilRet = gGetByKeyForUpdate("PRF", hmPrf, tmRec)
            'tmPrf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    Exit Sub
            'End If
            ilRet = btrDelete(hmPrf)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            tmPrfSrchKey.iAdfCode = tmAdf.iCode
            ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Loop
        tmPnfSrchKey2.iAdfCode = tmAdf.iCode
        ilRet = btrGetGreaterOrEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
        Do While (ilRet = BTRV_ERR_NONE) And (tlPnf.iAdfCode = tmAdf.iCode)
            'tmRec = tlPnf
            'ilRet = gGetByKeyForUpdate("PNF", hmPnf, tmRec)
            'tlPnf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    Exit Sub
            'End If
            ilRet = btrDelete(hmPnf)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            tmPnfSrchKey2.iAdfCode = tmAdf.iCode
            ilRet = btrGetGreaterOrEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
        Loop
        'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
        For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
            tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
            ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmPnf)
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            End If
        Next ilLoop
        'ReDim imNewPnfCode(1 To 1) As Integer
        ReDim imNewPnfCode(0 To 0) As Integer
        ilRet = btrDelete(hmAxf)
        ilRet = btrDelete(hmAdf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Advt
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "ADF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmAdf.iRemoteID
'            tmDsf.lAutoCode = tmAdf.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
'                Exit Sub
'            End If
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If Traffic!lbcAdvertiser.Tag <> "" Then
        '    If slStamp = Traffic!lbcAdvertiser.Tag Then
        '        Traffic!lbcAdvertiser.Tag = FileDateTime(sgDBPath & "ADF.Btr")
         '   End If
        'End If
        If sgAdvertiserTag <> "" Then
            If slStamp = sgAdvertiserTag Then
                sgAdvertiserTag = gFileDateTime(sgDBPath & "ADF.Btr")
            End If
        End If
        'Traffic!lbcAdvertiser.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgAdvertiser()
        cbcSelect.RemoveItem imSelectedIndex
        'Remove from tgCommAdf
        For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            If tgCommAdf(ilLoop).iCode = tmAdf.iCode Then
                For ilIndex = ilLoop To UBound(tgCommAdf) - 1 Step 1
                    tgCommAdf(ilIndex) = tgCommAdf(ilIndex + 1)
                Next ilIndex
                ReDim Preserve tgCommAdf(LBound(tgCommAdf) To UBound(tgCommAdf) - 1) As ADFEXT
                Exit For
            End If
        Next ilLoop
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
        'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
        For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
            tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
            ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmPnf)
                On Error GoTo cmcEraseErr
                gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Advt
                On Error GoTo 0
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
            End If
        Next ilLoop
        'ReDim imNewPnfCode(1 To 1) As Integer
        ReDim imNewPnfCode(0 To 0) As Integer
        Screen.MousePointer = vbDefault
    End If
    
    'L.Bianchi '05/31/2021' start
    tmAdfx.iCode = tmAdf.iCode
    SQLQuery = "select * from ADFX_Advertisers WHERE adfxCode =" & tmAdfx.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        SQLQuery = "DELETE FROM ADFX_Advertisers WHERE adfxCode = " & tmAdfx.iCode
         If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "Advt-cmcErase"
         End If
    End If
    tmAdfx.iCode = 0
    tmAdfx.sRefID = ""
    tmAdfx.sDirectRefId = "" 'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    tmAdfx.lCrmId = 0
    'L.Bianchi '05/31/2021' End
    
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcAdvt.Cls
    mPaintAdvtTitle
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub cmcMerge_Click()
    Dim slMsg As String
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If tgSpf.sRemoteUsers = "Y" Then
    '    slMsg = "Cannot Merge - Remote User System in Use"
    If tgUrf(0).iRemoteID > 0 Then
        slMsg = "Remote User Cannot Run Merge"
        ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Merge")
        Exit Sub
    End If
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbYesNo + vbQuestion, "Merge Advertiser")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("Are all other users off the traffic system", vbYesNo + vbQuestion, "Merge Advertiser")
    If ilRet = vbNo Then
        Exit Sub
    End If
    igMergeCallSource = ADVERTISERSLIST
    Merge.Show vbModal
    Screen.MousePointer = vbHourglass
    pbcAdvt.Cls
    pbcNotDirect.Cls
    pbcDirect.Cls
    mPaintAdvtTitle
    mPaintDirectTitle
    mPaintNotDirectTitle
    cbcSelect.Clear
    mPopulate
    cbcSelect.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub cmcSplitCue_Click()
    sgAVIndicatorID = smAVIndicatorID
    sgXDSCue = smXDSCue
    AdvtSplitID.Show vbModal
    If igAdvtSplitID <> 0 Then
        smAVIndicatorID = sgAVIndicatorID
        smXDSCue = sgXDSCue
        bmAxfChg = True
    End If
    mSetCommands
End Sub
Private Sub cmcSplitCue_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex, SETFORREADONLY, False) Then
            GoTo cmcUndoErr
        End If
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcAdvt.Cls
        mPaintAdvtTitle
        pbcAdvt_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    mClearCtrlFields
    pbcAdvt.Cls
    mPaintAdvtTitle
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = Trim$(edcName.Text)   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
      
    If mSaveRec_VAC() = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    imBoxNo = -1
    
    ''Must reset display so altered flag is cleared and setcommand will turn select on
    'If imSvSelectedIndex <> 0 Then
    '    cbcSelect.Text = slName
    'Else
    '    cbcSelect.ListIndex = 0
    'End If
    'cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmAdf.iCode
    cbcSelect.Clear
    sgAdvertiserTag = ""
    mPopulate
    For ilLoop = 0 To UBound(tgAdvertiser) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tgAdvertiser(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Val(slCode) = ilCode Then
            If cbcSelect.ListIndex = ilLoop + 1 Then
                cbcSelect_Change
            Else
                cbcSelect.ListIndex = ilLoop + 1
            End If
            Exit For
        End If
    Next ilLoop
    mProdPop ilCode
    mBuyerPop ilCode, "", -1
    mPayablePop ilCode, "", -1
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    Else
        cmcCancel.SetFocus
    End If
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus ActiveControl
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub edcAbbr_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcAbbr_GotFocus()
    If edcAbbr.Text = "" Then
        edcAbbr.Text = Left$(edcName.Text, 7)
        mSetChg imBoxNo   'Change event not generated
        mSetCommands
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAbbr_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcAddrID_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcAddrID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAddrID_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcAgyCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcAgyCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcBAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcBAddr_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
    If Index > 0 Then
        If (edcBAddr(0).Text = "") Then
            edcBAddr(1).Text = ""
            edcBAddr(2).Text = ""
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub edcCAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcCAddr_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcCEDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imCEBoxNo
        Case CECOMPINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcCEDropDown, lbcComp(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(0).ListIndex = 1
            End If
        Case CECOMPINDEX + 1
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcCEDropDown, lbcComp(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(1).ListIndex = 1
            End If
        Case CEEXCLINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcCEDropDown, lbcExcl(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(0).ListIndex = 1
            End If
        Case CEEXCLINDEX + 1
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcCEDropDown, lbcExcl(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcExcl(1).ListIndex = 1
            End If
    End Select
    imLbcArrowSetting = False
    mCESetChg imCEBoxNo
End Sub
Private Sub edcCEDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcCEDropDown_GotFocus()
    Select Case imCEBoxNo
        Case CECOMPINDEX
            If lbcComp(0).ListCount = 1 Then
                lbcComp(0).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcCESTab.SetFocus
                'Else
                '    pbcCETab.SetFocus
                'End If
                'Exit Sub
            End If
        Case CECOMPINDEX + 1
            If lbcComp(1).ListCount = 1 Then
                lbcComp(1).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcCESTab.SetFocus
                'Else
                '    pbcCETab.SetFocus
                'End If
                'Exit Sub
            End If
        Case CEEXCLINDEX
            If lbcExcl(0).ListCount = 1 Then
                lbcExcl(0).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcCESTab.SetFocus
                'Else
                '    pbcCETab.SetFocus
                'End If
                'Exit Sub
            End If
        Case CEEXCLINDEX + 1
            If lbcExcl(1).ListCount = 1 Then
                lbcExcl(1).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcCESTab.SetFocus
                'Else
                '    pbcCETab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcCEDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcCEDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcCEDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcCEDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imCEBoxNo
            Case CECOMPINDEX
                gProcessArrowKey Shift, KeyCode, lbcComp(0), imLbcArrowSetting
            Case CECOMPINDEX + 1
                gProcessArrowKey Shift, KeyCode, lbcComp(1), imLbcArrowSetting
            Case CEEXCLINDEX
                gProcessArrowKey Shift, KeyCode, lbcExcl(0), imLbcArrowSetting
            Case CEEXCLINDEX + 1
                gProcessArrowKey Shift, KeyCode, lbcExcl(1), imLbcArrowSetting
        End Select
        edcCEDropDown.SelStart = 0
        edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
    End If
End Sub
Private Sub edcCEDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imCEBoxNo
            Case CECOMPINDEX, CECOMPINDEX + 1, CEEXCLINDEX, CEEXCLINDEX + 1
                If imTabDirection = -1 Then  'Right To Left
                    pbcCESTab.SetFocus
                Else
                    pbcCETab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcCreditLimit_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCreditLimit_GotFocus()
    If lbcCreditRestr.ListIndex <> 1 Then
        pbcTab.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcCreditLimit_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcCreditLimit.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcCreditLimit.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcCreditLimit.Text
    slStr = Left$(slStr, edcCreditLimit.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCreditLimit.SelStart - edcCreditLimit.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Function GetTabState() As Boolean
    GetTabState = False
    lmKeyState = GetKeyState(VK_TAB)
    If lmKeyState And -256 Then
        GetTabState = True
    End If
End Function

Private Sub edcDirectRefID_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCRMID_Change()
    mSetChg CRMIDINDEX
End Sub

Private Sub edcMegaphoneAdvID_Change()
    mSetChg MEGAPHONEADVID
End Sub



Private Sub edcCRMID_LostFocus()
'    If Len(edcCRMID.Text) > 0 And IsNumeric(edcCRMID.Text) Then
'        If Val(edcCRMID.Text) < 2147483647 Then
'            tmAdfx.lCrmId = CLng(edcCRMID.Text)
'        End If
'    Else
'        edcCRMID.Text = ""
'    End If
End Sub

Private Sub edcDirectRefID_LostFocus()
'    lmKeyState = GetKeyState(VK_TAB)
'    If lmKeyState = -127 Then       ' Tab key with Shift key down
'        'imTabDirection = -1
'        'pbcSTab.SetFocus
'        imBoxNo = REFIDINDEX
'        mEnableBox imBoxNo
'    ElseIf lmKeyState = -128 Then   ' Tab key
'        imTabDirection = 0
'        pbcTab.SetFocus
'        imBoxNo = CRMIDINDEX
'        'mEnableBox imBoxNo
'    End If
End Sub

Private Sub edcDmDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imDMBoxNo
        Case DMDEMOINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDmDropDown, lbcDemo(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo(0).ListIndex = 0
            End If
        Case DMDEMOINDEX + 2
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDmDropDown, lbcDemo(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo(1).ListIndex = 0
            End If
        Case DMDEMOINDEX + 4
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDmDropDown, lbcDemo(2), imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo(2).ListIndex = 0
            End If
        Case DMDEMOINDEX + 6
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDmDropDown, lbcDemo(3), imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo(3).ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
    mDmSetChg imDMBoxNo
End Sub
Private Sub edcDmDropDown_GotFocus()
    Select Case imDMBoxNo
        Case DMDEMOINDEX
            If lbcDemo(0).ListCount = 1 Then
                lbcDemo(0).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcDmSTab.SetFocus
                'Else
                '    pbcDmTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case DMDEMOINDEX + 2
            If lbcDemo(1).ListCount = 1 Then
                lbcDemo(1).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcDmSTab.SetFocus
                'Else
                '    pbcDmTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case DMDEMOINDEX
            If lbcDemo(2).ListCount = 1 Then
                lbcDemo(2).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcDmSTab.SetFocus
                'Else
                '    pbcDmTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case DMDEMOINDEX + 6
            If lbcDemo(3).ListCount = 1 Then
                lbcDemo(3).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcDmSTab.SetFocus
                'Else
                '    pbcDmTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcDmDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDmDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDmDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (imDMBoxNo = DMVALUEINDEX) Or (imDMBoxNo = DMVALUEINDEX + 2) Or (imDMBoxNo = DMVALUEINDEX + 4) Or (imDMBoxNo = DMVALUEINDEX + 6) Then
        If imPriceType = 2 Then 'CPP (no decimal place)
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDmDropDown.Text
            slStr = Left$(slStr, edcDmDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDmDropDown.SelStart - edcDmDropDown.SelLength)
            If gCompNumberStr(slStr, "99999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Else
            ilPos = InStr(edcDmDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDmDropDown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDmDropDown.Text
            slStr = Left$(slStr, edcDmDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDmDropDown.SelStart - edcDmDropDown.SelLength)
            If gCompNumberStr(slStr, "999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub edcDmDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imDMBoxNo
            Case DMDEMOINDEX
                gProcessArrowKey Shift, KeyCode, lbcDemo(0), imLbcArrowSetting
                edcDmDropDown.SelStart = 0
                edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
            Case DMDEMOINDEX + 2
                gProcessArrowKey Shift, KeyCode, lbcDemo(1), imLbcArrowSetting
                edcDmDropDown.SelStart = 0
                edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
            Case DMDEMOINDEX + 4
                gProcessArrowKey Shift, KeyCode, lbcDemo(2), imLbcArrowSetting
                edcDmDropDown.SelStart = 0
                edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
            Case DMDEMOINDEX + 6
                gProcessArrowKey Shift, KeyCode, lbcDemo(3), imLbcArrowSetting
                edcDmDropDown.SelStart = 0
                edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
        End Select
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case PRODINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcProd.ListIndex = 0
            End If
            smProduct = edcDropDown.Text
        Case SPERSONINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSPerson, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSPerson.ListIndex = 1
            End If
        Case AGENCYINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAgency, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAgency.ListIndex = 1
            End If
        Case CREDITRESTRINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCreditRestr, imBSMode, imComboBoxIndex
        Case PAYMRATINGINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcPaymRating, imBSMode, imComboBoxIndex
        Case CREDITAPPROVALINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCreditApproval, imBSMode, imComboBoxIndex
        Case INVSORTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcInvSort, imBSMode, slStr)
            If ilRet = 1 Then
                lbcInvSort.ListIndex = 1
            End If
        Case BUYERINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcBuyer, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcBuyer.ListIndex = 0
            End If
            smBuyer = edcDropDown.Text
        Case PAYABLEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcPayable, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcPayable.ListIndex = 0
            End If
            smPayable = edcDropDown.Text
        Case LKBOXINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcLkBox, imBSMode, slStr)
            If ilRet = 1 Then
                lbcLkBox.ListIndex = -1
            End If
        Case EDICINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcEDI(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcEDI(0).ListIndex = 0
            End If
        Case EDIIINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcEDI(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcEDI(1).ListIndex = 0
            End If
        Case TERMSINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcTerms, imBSMode, slStr)
            If ilRet = 1 Then
                lbcTerms.ListIndex = 1
            End If
        Case TAXINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTax, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case PRODINDEX
'            If lbcProd.ListCount = 1 Then
'                lbcProd.ListIndex = 0
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
        Case SPERSONINDEX
            If lbcSPerson.ListCount = 1 Then
                lbcSPerson.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case AGENCYINDEX
            If lbcAgency.ListCount = 1 Then
                lbcAgency.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case CREDITRESTRINDEX
            If lbcCreditRestr.ListCount = 1 Then
                lbcCreditRestr.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case PAYMRATINGINDEX
            If lbcPaymRating.ListCount = 1 Then
                lbcPaymRating.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case CREDITAPPROVALINDEX
        Case INVSORTINDEX
'            If lbcInvSort.ListCount = 2 Then
'                lbcInvSort.ListIndex = 1
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
        Case BUYERINDEX
        Case PAYABLEINDEX
        Case LKBOXINDEX
'            If lbcLkBox.ListCount = 1 Then
'                lbcLkBox.ListIndex = 0
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
        Case EDICINDEX
            If lbcEDI(0).ListCount = 1 Then
                lbcEDI(0).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case EDIIINDEX
            If lbcEDI(1).ListCount = 1 Then
                lbcEDI(1).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case TERMSINDEX
        Case TAXINDEX
            If lbcTax.ListCount = 1 Then
                lbcTax.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    'If Not gCheckKeyAscii(ilKey) Then
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case PRODINDEX
                gProcessArrowKey Shift, KeyCode, lbcProd, imLbcArrowSetting
            Case SPERSONINDEX
                gProcessArrowKey Shift, KeyCode, lbcSPerson, imLbcArrowSetting
            Case AGENCYINDEX
                gProcessArrowKey Shift, KeyCode, lbcAgency, imLbcArrowSetting
            Case CREDITRESTRINDEX
                gProcessArrowKey Shift, KeyCode, lbcCreditRestr, imLbcArrowSetting
            Case PAYMRATINGINDEX
                gProcessArrowKey Shift, KeyCode, lbcPaymRating, imLbcArrowSetting
            Case CREDITAPPROVALINDEX
                gProcessArrowKey Shift, KeyCode, lbcCreditApproval, imLbcArrowSetting
            Case INVSORTINDEX
                gProcessArrowKey Shift, KeyCode, lbcInvSort, imLbcArrowSetting
            Case BUYERINDEX
                gProcessArrowKey Shift, KeyCode, lbcBuyer, imLbcArrowSetting
            Case PAYABLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcPayable, imLbcArrowSetting
            Case LKBOXINDEX
                gProcessArrowKey Shift, KeyCode, lbcLkBox, imLbcArrowSetting
            Case EDICINDEX
                gProcessArrowKey Shift, KeyCode, lbcEDI(0), imLbcArrowSetting
            Case EDIIINDEX
                gProcessArrowKey Shift, KeyCode, lbcEDI(1), imLbcArrowSetting
            Case TERMSINDEX
                gProcessArrowKey Shift, KeyCode, lbcTerms, imLbcArrowSetting
            Case TAXINDEX
                gProcessArrowKey Shift, KeyCode, lbcTax, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_LostFocus()
    '9760
    If (imBoxNo <> PRODINDEX) And (imBoxNo <> SPERSONINDEX) And (imBoxNo <> AGENCYINDEX) And (imBoxNo <> BUYERINDEX) And (imBoxNo <> PAYABLEINDEX) And (imBoxNo <> LKBOXINDEX) And (imBoxNo <> EDICINDEX) And (imBoxNo <> EDIIINDEX) And (imBoxNo <> INVSORTINDEX) And (imBoxNo <> TERMSINDEX) Then
        edcDropDown.Text = gRemoveIllegalPastedChar(edcDropDown.Text)
    End If
    Select Case imBoxNo
        Case PRODINDEX
            smProduct = edcDropDown.Text
        Case BUYERINDEX
            smBuyer = edcDropDown.Text
        Case PAYABLEINDEX
            smPayable = edcDropDown.Text
    End Select
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case PRODINDEX, SPERSONINDEX, AGENCYINDEX, BUYERINDEX, PAYABLEINDEX, LKBOXINDEX, INVSORTINDEX, EDICINDEX, EDIIINDEX, TERMSINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcName_Change()
    Dim slStr As String
    Dim ilRet As Integer

    If Not imEdcChgMode Then
        imEdcChgMode = True
        imLbcArrowSetting = True
        ilRet = gOptionalLookAhead(edcName, lbcName, imBSMode, slStr)
        If ilRet = 1 Then   'input was ""
            lbcName.ListIndex = -1
        End If
        imEdcChgMode = False
    End If
    imLbcArrowSetting = False
    mSetChg NAMEINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function- altered flag set so field is saved
End Sub
Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcName_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcName.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    'If Not gCheckKeyAscii(ilKey) Then
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcName_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcName, imLbcArrowSetting
        edcName.SelStart = 0
        edcName.SelLength = Len(edcName.Text)
    End If
End Sub

Private Sub edcName_LostFocus()
    '9760
    edcName.Text = gRemoveIllegalPastedChar(edcName.Text)
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
End Sub
Private Sub edcRating_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRating_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRating_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRepCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRepCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcStnCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcStnCode_GotFocus()
    gCtrlGotFocus ActiveControl
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
    If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcAdvt.Enabled = False
        pbcDirect.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        lacBill.Enabled = False
        rbcBill(0).Enabled = False
        rbcBill(1).Enabled = False
        pbcBillTab(0).Enabled = False
        pbcBillTab(1).Enabled = False
        imUpdateAllowed = False
    Else
        pbcAdvt.Enabled = True
        pbcDirect.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        lacBill.Enabled = True
        rbcBill(0).Enabled = True
        rbcBill(1).Enabled = True
        pbcBillTab(0).Enabled = True
        pbcBillTab(1).Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
    Me.KeyPreview = True
    Advt.Refresh
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
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imCEBoxNo > 0 Then
            mCEEnableBox imCEBoxNo
        ElseIf imDMBoxNo > 0 Then
            mDmEnableBox imDMBoxNo
        Else
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
        fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        If fmAdjFactorW < 1# Then
            fmAdjFactorW = 1#
        Else
            Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
            Me.Width = 12700
            'Me.height = 11385
        End If
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        'Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        fmAdjFactorH = 1#
        fmAdjFactorW = IIF(fmAdjFactorW > 1.425, 1.425, fmAdjFactorW)  ' TTP 10622 - 2023-03-08 JJB
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    If Not igManUnload Then
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Erase imNewPnfCode
    Erase tmEDICode
    Erase tmInvSortCode
    Erase tmTermsCode
    Erase tmLkBoxCode
    Erase tmTaxSortCode
    Erase tmPdf

    ilRet = btrClose(hmAxf)
    btrDestroy hmAxf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmPDF)
    btrDestroy hmPDF
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    ilRet = btrClose(hmAxf)
    btrDestroy hmAxf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    
    Set Advt = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lacBill_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lbcAgency_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAgency, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcAgency_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcAgency_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcAgency_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcAgency_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAgency, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcBuyer_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcBuyer, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcBuyer_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcBuyer_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBuyer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcBuyer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcBuyer, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcComp_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcComp(Index), edcCEDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcComp_DblClick(Index As Integer)
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcComp_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcComp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcComp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcComp(Index), edcCEDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcCESTab.SetFocus
        Else
            pbcCETab.SetFocus
        End If
    End If
End Sub
Private Sub lbcCreditApproval_Click()
    gProcessLbcClick lbcCreditApproval, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCreditApproval_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcCreditRestr_Click()
    gProcessLbcClick lbcCreditRestr, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCreditRestr_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDemo_Click(Index As Integer)
    gProcessLbcClick lbcDemo(Index), edcDmDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcDemo_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEDI_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcEDI(Index), edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcEDI_DblClick(Index As Integer)
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcEDI_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEDI_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcEDI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcEDI(Index), edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcExcl_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcExcl(Index), edcCEDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcExcl_DblClick(Index As Integer)
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcExcl_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcExcl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcExcl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcExcl(Index), edcCEDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcCESTab.SetFocus
        Else
            pbcCETab.SetFocus
        End If
    End If
End Sub
Private Sub lbcInvSort_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcInvSort, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcInvSort_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcInvSort_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcInvSort_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcInvSort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcInvSort, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcLkBox_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcLkBox, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcLkBox_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcLkBox_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcLkBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcLkBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcLkBox, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcName_Click()
    gProcessLbcClick lbcName, edcName, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPayable_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcPayable, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcPayable_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcPayable_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPayable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcPayable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcPayable, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcPaymRating_Click()
    gProcessLbcClick lbcPaymRating, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcPaymRating_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcProd_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcProd_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcProd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcProd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcSPerson_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSPerson, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSPerson_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSPerson_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSPerson_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSPerson_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSPerson, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcTax_Click()
    gProcessLbcClick lbcTax, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcTax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAgencyBranch                   *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Agency *
'*                      and process communication      *
'*                      back from agency               *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mAgencyBranch() As Integer
'
'   ilRet = mAgencyBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcAgency, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mAgencyBranch = False
        Exit Function
    End If
    If igWinStatus(AGENCIESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mAgencyBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(AGENCIESLIST)) Then
    '    imDoubleClickName = False
    '    mAgencyBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igAgyCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgAgyName = ""
    Else
        sgAgyName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & Trim$(str$(igAgyCallSource)) & "\" & sgAgyName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & Trim$(str$(igAgyCallSource)) & "\" & sgAgyName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAgyCallSource)) & "\" & sgAgyName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAgyCallSource)) & "\" & sgAgyName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Agency.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Agency.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAgyName)
    igAgyCallSource = Val(sgAgyName)
    ilParse = gParseItem(slStr, 2, "\", sgAgyName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAgencyBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igAgyCallSource = CALLDONE Then  'Done
        igAgyCallSource = CALLNONE
'        gSetMenuState True
        lbcAgency.Clear
        sgAgencyTag = ""
'        sgCommAgfStamp = ""
        mAgencyPop
        If imTerminate Then
            mAgencyBranch = False
            Exit Function
        End If
        gFindMatch sgAgyName, 1, lbcAgency
        sgAgyName = ""
        If gLastFound(lbcAgency) > 0 Then
            imChgMode = True
            lbcAgency.ListIndex = gLastFound(lbcAgency)
            edcDropDown.Text = lbcAgency.List(lbcAgency.ListIndex)
            imChgMode = False
            mAgencyBranch = False
            mSetChg AGENCYINDEX
        Else
            imChgMode = True
            lbcAgency.ListIndex = 1
            edcDropDown.Text = lbcAgency.List(1)
            imChgMode = False
            mSetChg AGENCYINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igAgyCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAgyCallSource = CALLNONE
        sgAgyName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAgyCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igAgyCallSource = CALLNONE
        sgAgyName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAgencyPop                      *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency list box       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAgencyPop()
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcAgency.ListIndex
    If ilIndex > 1 Then
        slName = lbcAgency.List(ilIndex)
    End If
    'Repopulate if required- if agency changed by another user while in this screen
    'ilRet = gPopAgyBox(Advt, lbcAgency, Traffic!lbcAgency)
    ilRet = gPopAgyBox(Advt, lbcAgency, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcAgency.AddItem "[None]", 0
        lbcAgency.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcAgency
            If gLastFound(lbcAgency) > 1 Then
                lbcAgency.ListIndex = gLastFound(lbcAgency)
            Else
                lbcAgency.ListIndex = -1
            End If
        Else
            lbcAgency.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBuyerPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Buyer Personnel       *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mBuyerPop(ilAdvtCode As Integer, slRetainName As String, ilReturnCode As Integer)
'
'   mBuyerPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilPnf As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    ilIndex = lbcBuyer.ListIndex
    If ilIndex > 0 Then
        slName = lbcBuyer.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'If imSelectedIndex > 0 Then 'Change mode
        'If imSelectedIndex = 0 Then
        '    ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "B", False, lbcBuyer, lbcBuyerCode)
        'Else
            'ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "B", True, 2, lbcBuyer, lbcBuyerCode)
            ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "B", True, 2, lbcBuyer, tgBuyerCode(), sgBuyerCodeTag)
        'End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mBuyerPopErr
            gCPErrorMsg ilRet, "mBuyerPop (gPopPersonnelBox)", Advt
            On Error GoTo 0
            'Filter out any contact not associated with this agency
            If imSelectedIndex = 0 Then
                For ilLoop = UBound(tgBuyerCode) - 1 To 0 Step -1 'lbcBuyerCode.ListCount - 1 To 0 Step -1
                    ilFound = False
                    slNameCode = tgBuyerCode(ilLoop).sKey  'lbcBuyerCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If (StrComp(Trim$(slRetainName), Trim$(Left$(slName, 30)), 1) = 0) And (slRetainName <> "") Or (Val(slCode) = ilReturnCode) And (slRetainName <> "") Then
                        ilFound = True
                    Else
                        'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilCode = Val(slCode)
                        'For ilPnf = 1 To UBound(imNewPnfCode) - 1 Step 1
                        For ilPnf = 0 To UBound(imNewPnfCode) - 1 Step 1
                            If imNewPnfCode(ilPnf) = ilCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilPnf
                    End If
                    If Not ilFound Then
                        lbcBuyer.RemoveItem ilLoop
                        gRemoveItemFromSortCode ilLoop, tgBuyerCode()
                        'lbcBuyerCode.RemoveItem ilLoop
                    End If
                Next ilLoop
            End If
            lbcBuyer.AddItem "[None]", 0  'Force as first item on list
            lbcBuyer.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 0 Then
                gFindMatch slName, 1, lbcBuyer
                If gLastFound(lbcBuyer) > 0 Then
                    lbcBuyer.ListIndex = gLastFound(lbcBuyer)
                Else
                    lbcBuyer.ListIndex = -1
                End If
            Else
                lbcBuyer.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    'Else
    '    If lbcBuyer.ListCount = 0 Then
    '        lbcBuyer.AddItem "[None]", 0  'Force as first item on list
    '        lbcBuyer.AddItem "[New]", 0  'Force as first item on list
    '    End If
    'End If
    Exit Sub
mBuyerPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCEEnableBox                    *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mCEEnableBox(ilBoxNo As Integer)
'
'   mCEEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCECtrls Or ilBoxNo > UBound(tmCECtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case CECOMPINDEX   'Competitive
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(0).height = gListBoxHeight(lbcComp(0).ListCount, 7)
            edcCEDropDown.Width = tmCECtrls(ilBoxNo).fBoxW - cmcCEDropDown.Width
            edcCEDropDown.MaxLength = 20
            gMoveFormCtrl pbcCE, edcCEDropDown, tmCECtrls(ilBoxNo).fBoxX, tmCECtrls(ilBoxNo).fBoxY
            cmcCEDropDown.Move edcCEDropDown.Left + edcCEDropDown.Width, edcCEDropDown.Top
            lbcComp(0).Move edcCEDropDown.Left, edcCEDropDown.Top + edcCEDropDown.height
            imChgMode = True
            If lbcComp(0).ListIndex < 0 Then
                lbcComp(0).ListIndex = 1   '[None]
            End If
            If lbcComp(0).ListIndex < 0 Then
                edcCEDropDown.Text = ""
            Else
                edcCEDropDown.Text = lbcComp(0).List(lbcComp(0).ListIndex)
            End If
            imChgMode = False
            edcCEDropDown.SelStart = 0
            edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
            edcCEDropDown.Visible = True
            cmcCEDropDown.Visible = True
            edcCEDropDown.SetFocus
        Case CECOMPINDEX + 1 'Competitive
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(1).height = gListBoxHeight(lbcComp(1).ListCount, 7)
            edcCEDropDown.Width = tmCECtrls(ilBoxNo).fBoxW - cmcCEDropDown.Width
            edcCEDropDown.MaxLength = 20
            gMoveFormCtrl pbcCE, edcCEDropDown, tmCECtrls(ilBoxNo).fBoxX, tmCECtrls(ilBoxNo).fBoxY
            cmcCEDropDown.Move edcCEDropDown.Left + edcCEDropDown.Width, edcCEDropDown.Top
            lbcComp(1).Move edcCEDropDown.Left, edcCEDropDown.Top + edcCEDropDown.height
            imChgMode = True
            If lbcComp(1).ListIndex < 0 Then
                lbcComp(1).ListIndex = 1   '[None]
            End If
            If lbcComp(1).ListIndex < 0 Then
                edcCEDropDown.Text = ""
            Else
                edcCEDropDown.Text = lbcComp(1).List(lbcComp(1).ListIndex)
            End If
            imChgMode = False
            edcCEDropDown.SelStart = 0
            edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
            edcCEDropDown.Visible = True
            cmcCEDropDown.Visible = True
            edcCEDropDown.SetFocus
        Case CEEXCLINDEX   'Program exclusion
            mExclPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExcl(0).height = gListBoxHeight(lbcExcl(0).ListCount, 11)
            edcCEDropDown.Width = tmCECtrls(ilBoxNo).fBoxW - cmcCEDropDown.Width
            edcCEDropDown.MaxLength = 20
            gMoveFormCtrl pbcCE, edcCEDropDown, tmCECtrls(ilBoxNo).fBoxX, tmCECtrls(ilBoxNo).fBoxY
            cmcCEDropDown.Move edcCEDropDown.Left + edcCEDropDown.Width, edcCEDropDown.Top
            lbcExcl(0).Move edcCEDropDown.Left, edcCEDropDown.Top + edcCEDropDown.height
            imChgMode = True
            If lbcExcl(0).ListIndex < 0 Then
                lbcExcl(0).ListIndex = 1   '[None]
            End If
            If lbcExcl(0).ListIndex < 0 Then
                edcCEDropDown.Text = ""
            Else
                edcCEDropDown.Text = lbcExcl(0).List(lbcExcl(0).ListIndex)
            End If
            imChgMode = False
            edcCEDropDown.SelStart = 0
            edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
            edcCEDropDown.Visible = True
            cmcCEDropDown.Visible = True
            edcCEDropDown.SetFocus
        Case CEEXCLINDEX + 1 'Program exclusion
            mExclPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExcl(1).height = gListBoxHeight(lbcExcl(1).ListCount, 11)
            edcCEDropDown.Width = tmCECtrls(ilBoxNo).fBoxW - cmcCEDropDown.Width
            edcCEDropDown.MaxLength = 20
            gMoveFormCtrl pbcCE, edcCEDropDown, tmCECtrls(ilBoxNo).fBoxX, tmCECtrls(ilBoxNo).fBoxY
            cmcCEDropDown.Move edcCEDropDown.Left + edcCEDropDown.Width, edcCEDropDown.Top
            lbcExcl(1).Move edcCEDropDown.Left, edcCEDropDown.Top + edcCEDropDown.height
            imChgMode = True
            If lbcExcl(1).ListIndex < 0 Then
                lbcExcl(1).ListIndex = 1   '[None]
            End If
            If lbcExcl(1).ListIndex < 0 Then
                edcCEDropDown.Text = ""
            Else
                edcCEDropDown.Text = lbcExcl(1).List(lbcExcl(1).ListIndex)
            End If
            imChgMode = False
            edcCEDropDown.SelStart = 0
            edcCEDropDown.SelLength = Len(edcCEDropDown.Text)
            edcCEDropDown.Visible = True
            cmcCEDropDown.Visible = True
            edcCEDropDown.SetFocus
    End Select
    mCESetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCESetChg                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mCESetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim ilLoop As Integer   'For loop control parameter
    If ilBoxNo < imLBCECtrls Or ilBoxNo > UBound(tmCECtrls) Then
        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case CECOMPINDEX 'Competitive
            gSetChgFlag smComp(0), lbcComp(0), tmCECtrls(CECOMPINDEX)
        Case CECOMPINDEX + 1 'Competitive
            gSetChgFlag smComp(1), lbcComp(1), tmCECtrls(CECOMPINDEX + 1)
        Case CEEXCLINDEX 'Program exclusion
            gSetChgFlag smExcl(0), lbcExcl(0), tmCECtrls(CEEXCLINDEX)
        Case CEEXCLINDEX + 1 'Program exclusion
            gSetChgFlag smExcl(1), lbcExcl(1), tmCECtrls(CEEXCLINDEX + 1)
    End Select
    tmCtrls(COMPINDEX).iChg = False
    For ilLoop = imLBCECtrls To UBound(tmCECtrls) Step 1
        If tmCECtrls(ilLoop).iChg Then
            tmCtrls(COMPINDEX).iChg = True
        End If
    Next ilLoop
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCESetFocus                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mCESetFocus(ilBoxNo As Integer)
'
'   mCESetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCECtrls Or ilBoxNo > UBound(tmCECtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case CECOMPINDEX   'Competitive
            edcCEDropDown.SetFocus
        Case CECOMPINDEX + 1 'Competitive
            edcCEDropDown.SetFocus
        Case CEEXCLINDEX   'Program exclusion
            edcCEDropDown.SetFocus
        Case CEEXCLINDEX + 1 'Program exclusion
            edcCEDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCESetShow                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mCESetShow(ilBoxNo As Integer)
'
'   mCESetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCECtrls) Or (ilBoxNo > UBound(tmCECtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case CECOMPINDEX   'Competitive
            lbcComp(0).Visible = False
            edcCEDropDown.Visible = False
            cmcCEDropDown.Visible = False
            If lbcComp(0).ListIndex > 0 Then
                slStr = lbcComp(0).Text
            Else
                slStr = ""
            End If
            gSetShow pbcCE, slStr, tmCECtrls(ilBoxNo)
        Case CECOMPINDEX + 1 'Competitive
            lbcComp(1).Visible = False
            edcCEDropDown.Visible = False
            cmcCEDropDown.Visible = False
            If lbcComp(1).ListIndex > 0 Then
                slStr = lbcComp(1).Text
            Else
                slStr = ""
            End If
            gSetShow pbcCE, slStr, tmCECtrls(ilBoxNo)
        Case CEEXCLINDEX   'Program exclusion
            lbcExcl(0).Visible = False
            edcCEDropDown.Visible = False
            cmcCEDropDown.Visible = False
            If lbcExcl(0).ListIndex > 0 Then
                slStr = lbcExcl(0).Text
            Else
                slStr = ""
            End If
            gSetShow pbcCE, slStr, tmCECtrls(ilBoxNo)
        Case CEEXCLINDEX + 1 'Program exclusion
            lbcExcl(1).Visible = False
            edcCEDropDown.Visible = False
            cmcCEDropDown.Visible = False
            If lbcExcl(1).ListIndex > 0 Then
                slStr = lbcExcl(1).Text
            Else
                slStr = ""
            End If
            gSetShow pbcCE, slStr, tmCECtrls(ilBoxNo)
    End Select
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
Private Function mCETestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'

    If (ilCtrlNo = CECOMPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcComp(0), "", "Competitive must be specified", tmCECtrls(CECOMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imCEBoxNo = CECOMPINDEX
            End If
            mCETestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CECOMPINDEX + 1) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcComp(1), "", "Competitive Code must be specified", tmCECtrls(CECOMPINDEX + 1).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imCEBoxNo = CECOMPINDEX + 1
            End If
            mCETestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CEEXCLINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcExcl(0), "", "Program Exclusion must be specified", tmCECtrls(CEEXCLINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imCEBoxNo = CEEXCLINDEX
            End If
            mCETestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CEEXCLINDEX + 1) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcExcl(1), "", "Program Exclusion Code must be specified", tmCECtrls(CEEXCLINDEX + 1).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imCEBoxNo = CEEXCLINDEX + 1
            End If
            mCETestFields = NO
            Exit Function
        End If
    End If
    mCETestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
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
'       ilOnlyAddr (I)- Clear only fields after address
'
    Dim ilLoop As Integer
    tmAdf.iCode = 0
    edcName.Text = ""
    edcAbbr.Text = ""
    edcRefId.Text = "" 'L.Bianchi 04/15/2021
    edcDirectRefID.Text = "" 'JW - 7/30/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    imState = -1
    lbcProd.Clear
    'lbcProdCode.Clear
    'lbcProdCode.Tag = ""
    ReDim tgProdCode(0 To 0) As SORTCODE
    sgProdCodeTag = ""
    lbcProd.ListIndex = -1
    lbcSPerson.ListIndex = -1   'This will set the text to "", setting text will not set index to -1
    lbcAgency.ListIndex = -1
    edcRepCode.Text = ""
    edcAgyCode.Text = ""
    edcStnCode.Text = ""
    lbcComp(0).ListIndex = -1
    lbcComp(1).ListIndex = -1
    lbcExcl(0).ListIndex = -1
    lbcExcl(1).ListIndex = -1
'    cbcPriceType.ListIndex = -1
    imPriceType = -1
    lbcDemo(0).ListIndex = -1
    lbcDemo(1).ListIndex = -1
    lbcDemo(2).ListIndex = -1
    lbcDemo(3).ListIndex = -1
    smTarget(0) = ""
    smTarget(1) = ""
    smTarget(2) = ""
    smTarget(3) = ""
    smSvTarget(0) = ""
    smSvTarget(1) = ""
    smSvTarget(2) = ""
    smSvTarget(3) = ""
    smBkoutPool = ""
'    edcTarget(0).Text = ""
'    edcTarget(1).Text = ""
'    edcTarget(2).Text = ""
'    edcTarget(3).Text = ""
    lbcCreditRestr.ListIndex = -1
    edcCreditLimit.Text = ""
    lbcPaymRating.ListIndex = -1
    lbcCreditApproval.ListIndex = -1
    edcRating.Text = ""
    imPolitical = -1
    imRateOnInv = -1
    imISCI = -1
    imPackage = -1
    imRepMG = -1
    imRepInv = -1
    imBonusOnInv = -1
    lbcInvSort.ListIndex = -1
    For ilLoop = 0 To 2 Step 1
        edcCAddr(ilLoop).Text = ""
        edcBAddr(ilLoop).Text = ""
    Next ilLoop
    edcAddrID.Text = ""
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    lbcBuyer.Clear
    'lbcBuyerCode.Clear
    'lbcBuyerCode.Tag = ""
    ReDim tgBuyerCode(0 To 0) As SORTCODE
    sgBuyerCodeTag = ""
    lbcBuyer.ListIndex = -1
    lbcPayable.Clear
    'lbcPayableCode.Clear
    'lbcPayableCode.Tag = ""
    ReDim tgPayableCode(0 To 0) As SORTCODE
    sgPayableCodeTag = ""
    lbcPayable.ListIndex = -1
    lbcLkBox.ListIndex = -1
    lbcEDI(0).ListIndex = -1
    lbcEDI(1).ListIndex = -1
    imPrtStyle = -1
    lbcTax.ListIndex = -1
    smProduct = ""
    smSPerson = ""
    smAgency = ""
    smComp(0) = ""
    smComp(1) = ""
    smExcl(0) = ""
    smExcl(1) = ""
    smLkBox = ""
    smEDIC = ""
    smEDII = ""
    smInvSort = ""
    smTerms = ""
    smTax = ""
    smBuyer = ""
    smPayable = ""
    smOrigBuyer = ""
    smOrigPayable = ""
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
        gSetShow pbcAdvt, "", tmCtrls(ilLoop)   'This is here since number of controls change and cbcSelect does not clear the direct area
    Next ilLoop
    For ilLoop = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        tmDmCtrls(ilLoop).iChg = False
    Next ilLoop
    For ilLoop = imLBCECtrls To UBound(tmCECtrls) Step 1
        tmCECtrls(ilLoop).iChg = False
    Next ilLoop
    smPct90 = ""
    smCurrAR = ""
    smUnbilled = ""
    smHiCredit = ""
    smTotalGross = ""
    smNSFChks = ""
    smDateEntrd = ""
    smDateLstInv = ""
    smDateLstPaym = ""
    smAvgToPay = ""
    smLstToPay = ""
    smNoInvPd = ""
    rbcBill(0).Value = True
    imCreditRestrFirst = True
    imPaymRatingFirst = True
'    imPrtStyleFirst = True
    imScriptFirst = True
    imPriceTypeFirst = True
    imTaxFirst = True
    bmAxfChg = False
    smAVIndicatorID = ""
    smXDSCue = ""
    bmPDFEMailChgd = False
    ReDim tmPdf(0 To 0) As PDF
    
    tmCtrls(CRMIDINDEX).sShow = ""
    edcCRMID.Text = ""
    edcMegaphoneAdvID.Text = ""
    tmAdfx.lCrmId = 0
    edcSuppressNet.Text = "  "  ' TTP 10622 - 2023-03-08 JJB
    imSuppressNet = 0          ' TTP 10622 - 2023-03-08 JJB
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompBranch                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      competitive and process             *
'*                      communication back from        *
'*                      competitive                    *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mCompBranch(ilIndex As Integer) As Integer
'
'   ilRet = mCompBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcCEDropDown, lbcComp(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcCEDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mCompBranch = False
        Exit Function
    End If
    If igWinStatus(COMPETITIVESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mCompBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(COMPETITIVESLIST)) Then
    '    imDoubleClickName = False
    '    mCompBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "C"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcCEDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mCompBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcComp(ilIndex).Clear
        sgCompCodeTag = ""
        sgCompMnfStamp = ""
        ilRet = csiSetStamp("COMPMNF", sgCompMnfStamp)
        mCompPop
        If imTerminate Then
            mCompBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcComp(ilIndex)
        sgMNmName = ""
        If gLastFound(lbcComp(ilIndex)) > 0 Then
            imChgMode = True
            lbcComp(ilIndex).ListIndex = gLastFound(lbcComp(ilIndex))
            edcCEDropDown.Text = lbcComp(ilIndex).List(lbcComp(ilIndex).ListIndex)
            imChgMode = False
            mCompBranch = False
            mCESetChg imCEBoxNo
        Else
            imChgMode = True
            lbcComp(ilIndex).ListIndex = 1
            edcCEDropDown.Text = lbcComp(ilIndex).List(1)
            imChgMode = False
            mCESetChg imCEBoxNo
            edcCEDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCEEnableBox imCEBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCEEnableBox imCEBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate competitive list      *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mCompPop()
'
'   mCompPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slComp(0 To 1) As String      'Competitive name, saved to determine if changed
    ReDim ilComp(0 To 1) As Integer      'Competitive name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilComp(0) = lbcComp(0).ListIndex
    ilComp(1) = lbcComp(1).ListIndex
    If ilComp(0) > 1 Then
        slComp(0) = lbcComp(0).List(ilComp(0))
    End If
    If ilComp(1) > 1 Then
        slComp(1) = lbcComp(1).List(ilComp(1))
    End If
    If lbcComp(0).ListCount <> lbcComp(1).ListCount Then
        lbcComp(0).Clear
    End If
    'ilRet = gIMoveListBox(Advt, lbcComp(0), lbcCompCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Advt, lbcComp(0), tgCompCode(), sgCompCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mCompPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcComp(0).AddItem "[None]", 0
        lbcComp(0).AddItem "[New]", 0  'Force as first item on list
        lbcComp(1).Clear
        For ilLoop = lbcComp(0).ListCount - 1 To 0 Step -1
            lbcComp(1).AddItem lbcComp(0).List(ilLoop), 0
        Next ilLoop
        imChgMode = True
        If ilComp(0) > 1 Then
            gFindMatch slComp(0), 2, lbcComp(0)
            If gLastFound(lbcComp(0)) > 1 Then
                lbcComp(0).ListIndex = gLastFound(lbcComp(0))
            Else
                lbcComp(0).ListIndex = -1
            End If
        Else
            lbcComp(0).ListIndex = ilComp(0)
        End If
        If ilComp(1) > 1 Then
            gFindMatch slComp(1), 2, lbcComp(1)
            If gLastFound(lbcComp(1)) > 1 Then
                lbcComp(1).ListIndex = gLastFound(lbcComp(1))
            Else
                lbcComp(1).ListIndex = -1
            End If
        Else
            lbcComp(1).ListIndex = ilComp(1)
        End If
        imChgMode = False
    End If
    Exit Sub
mCompPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Demo list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    'ilRet = gPopMnfPlusFieldsBox(Advt, lbcDemo(0), lbcDemoCode, "D")
    ilRet = gPopMnfPlusFieldsBox(Advt, lbcDemo(0), tgDemoCode(), sgDemoCodeTag, "D")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox)", Advt
        On Error GoTo 0
        lbcDemo(0).AddItem "[N/A]", 0
        For ilLoop = 0 To lbcDemo(0).ListCount - 1 Step 1
            lbcDemo(1).AddItem lbcDemo(0).List(ilLoop)
            lbcDemo(2).AddItem lbcDemo(0).List(ilLoop)
            lbcDemo(3).AddItem lbcDemo(0).List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDmEnableBox                    *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mDmEnableBox(ilBoxNo As Integer)
'
'   mDmEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilIndex As Integer
    If ilBoxNo < imLBDmCtrls Or (ilBoxNo > UBound(tmDmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DMPRICETYPEINDEX
            If imPriceType < 0 Then
                imPriceType = 0    'N/A
                tmCtrls(DEMOINDEX).iChg = True
                tmDmCtrls(DMPRICETYPEINDEX).iChg = True
                mSetCommands
            End If
            pbcDmPriceType.Width = tmDmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDm, pbcDmPriceType, tmDmCtrls(ilBoxNo).fBoxX, tmDmCtrls(ilBoxNo).fBoxY
            pbcDmPriceType_Paint
            pbcDmPriceType.Visible = True
            pbcDmPriceType.SetFocus
        Case DMDEMOINDEX, DMDEMOINDEX + 2, DMDEMOINDEX + 4, DMDEMOINDEX + 6
            ilIndex = (ilBoxNo - DMDEMOINDEX) \ 2
            lbcDemo(ilIndex).height = gListBoxHeight(lbcDemo(ilIndex).ListCount, 6)
            edcDmDropDown.Width = tmDmCtrls(ilBoxNo).fBoxW - cmcDmDropDown.Width
            edcDmDropDown.MaxLength = 6
            gMoveFormCtrl pbcDm, edcDmDropDown, tmDmCtrls(ilBoxNo).fBoxX, tmDmCtrls(ilBoxNo).fBoxY
            cmcDmDropDown.Move edcDmDropDown.Left + edcDmDropDown.Width, edcDmDropDown.Top
            lbcDemo(ilIndex).Move edcDmDropDown.Left, edcDmDropDown.Top + edcDmDropDown.height
            imChgMode = True
            If lbcDemo(ilIndex).ListIndex < 0 Then
                lbcDemo(ilIndex).ListIndex = 0   'First one within list
            End If
            If lbcDemo(ilIndex).ListIndex < 0 Then
                edcDmDropDown.Text = ""
            Else
                edcDmDropDown.Text = lbcDemo(ilIndex).List(lbcDemo(ilIndex).ListIndex)
            End If
            imChgMode = False
            edcDmDropDown.SelStart = 0
            edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
            edcDmDropDown.Visible = True
            cmcDmDropDown.Visible = True
            edcDmDropDown.SetFocus
        Case DMVALUEINDEX, DMVALUEINDEX + 2, DMVALUEINDEX + 4, DMVALUEINDEX + 6
            ilIndex = (ilBoxNo - DMVALUEINDEX) \ 2
            edcDmDropDown.Width = tmDmCtrls(ilBoxNo).fBoxW
            edcDmDropDown.MaxLength = 8
            gMoveFormCtrl pbcDm, edcDmDropDown, tmDmCtrls(ilBoxNo).fBoxX, tmDmCtrls(ilBoxNo).fBoxY
            edcDmDropDown.Text = smTarget(ilIndex)
            edcDmDropDown.SelStart = 0
            edcDmDropDown.SelLength = Len(edcDmDropDown.Text)
            edcDmDropDown.Visible = True  'Set visibility
            edcDmDropDown.SetFocus
    End Select
    mDmSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDmSetChg                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mDmSetChg(ilBoxNo As Integer)
'
'   mDmSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    If ilBoxNo < imLBDmCtrls Or ilBoxNo > UBound(tmDmCtrls) Then
        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DMPRICETYPEINDEX
        Case DMDEMOINDEX, DMDEMOINDEX + 2, DMDEMOINDEX + 4, DMDEMOINDEX + 6
            ilIndex = (ilBoxNo - DMDEMOINDEX) \ 2
            gSetChgFlag smSvDemo(ilIndex), lbcDemo(ilIndex), tmDmCtrls(ilBoxNo)
        Case DMVALUEINDEX, DMVALUEINDEX + 2, DMVALUEINDEX + 4, DMVALUEINDEX + 6
            ilIndex = (ilBoxNo - DMVALUEINDEX) \ 2
            gSetChgFlag smSvTarget(ilIndex), edcDmDropDown, tmDmCtrls(ilBoxNo)
    End Select
    tmCtrls(DEMOINDEX).iChg = False
    For ilLoop = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        If tmDmCtrls(ilLoop).iChg Then
            tmCtrls(DEMOINDEX).iChg = True
        End If
    Next ilLoop
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDmSetFocus                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mDmSetFocus(ilBoxNo As Integer)
'
'   mDmSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBDmCtrls Or (ilBoxNo > UBound(tmDmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DMPRICETYPEINDEX
            pbcDmPriceType.SetFocus
        Case DMDEMOINDEX, DMDEMOINDEX + 2, DMDEMOINDEX + 4, DMDEMOINDEX + 6
            edcDmDropDown.SetFocus
        Case DMVALUEINDEX, DMVALUEINDEX + 2, DMVALUEINDEX + 4, DMVALUEINDEX + 6
            edcDmDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDmSetShow                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mDmSetShow(ilBoxNo As Integer)
'
'   mDmSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilIndex As Integer
    Dim slStr As String
    If ilBoxNo < imLBDmCtrls Or (ilBoxNo > UBound(tmDmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DMPRICETYPEINDEX
            pbcDmPriceType.Visible = False  'Set visibility
            If imPriceType = 0 Then
                slStr = "N/A"
            ElseIf imPriceType = 1 Then
                slStr = "CPM"
            ElseIf imPriceType = 2 Then
                slStr = "CPP"
            Else
                slStr = ""
            End If
            gSetShow pbcDm, slStr, tmDmCtrls(ilBoxNo)
        Case DMDEMOINDEX, DMDEMOINDEX + 2, DMDEMOINDEX + 4, DMDEMOINDEX + 6
            ilIndex = (ilBoxNo - DMDEMOINDEX) \ 2
            lbcDemo(ilIndex).Visible = False
            edcDmDropDown.Visible = False
            cmcDmDropDown.Visible = False
            If lbcDemo(ilIndex).ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcDemo(ilIndex).List(lbcDemo(ilIndex).ListIndex)
            End If
            gSetShow pbcDm, slStr, tmDmCtrls(ilBoxNo)
        Case DMVALUEINDEX, DMVALUEINDEX + 2, DMVALUEINDEX + 4, DMVALUEINDEX + 6
            ilIndex = (ilBoxNo - DMVALUEINDEX) \ 2
            edcDmDropDown.Visible = False  'Set visibility
            smTarget(ilIndex) = edcDmDropDown.Text
            If imPriceType = 2 Then
                gFormatStr smTarget(ilIndex), FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            Else
                gFormatStr smTarget(ilIndex), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            End If
            gSetShow pbcDm, slStr, tmDmCtrls(ilBoxNo)
    End Select
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
Private Function mDmTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String

    If (ilCtrlNo = DMPRICETYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imPriceType = 0 Then
            slStr = "N/A"
        ElseIf imPriceType = 1 Then
            slStr = "CPM"
        ElseIf imPriceType = 2 Then
            slStr = "CPP"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Price type must be specified", tmDmCtrls(DMPRICETYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMPRICETYPEINDEX
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMDEMOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcDemo(0), "", "Demo must be specified", tmDmCtrls(DMDEMOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMDEMOINDEX
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMVALUEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smTarget(0), "", "CPM or CPP must be specified", tmDmCtrls(DMVALUEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMVALUEINDEX
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMDEMOINDEX + 2) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcDemo(1), "", "Demo must be specified", tmDmCtrls(DMDEMOINDEX + 2).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMDEMOINDEX + 2
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMVALUEINDEX + 2) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smTarget(1), "", "CPM or CPP must be specified", tmDmCtrls(DMVALUEINDEX + 2).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMVALUEINDEX + 2
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMDEMOINDEX + 4) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcDemo(2), "", "Demo must be specified", tmDmCtrls(DMDEMOINDEX + 4).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMDEMOINDEX + 4
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMVALUEINDEX + 4) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smTarget(2), "", "CPM or CPP must be specified", tmDmCtrls(DMVALUEINDEX + 4).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMVALUEINDEX + 4
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMDEMOINDEX + 6) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcDemo(3), "", "Demo must be specified", tmDmCtrls(DMDEMOINDEX + 6).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMDEMOINDEX + 6
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DMVALUEINDEX + 6) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smTarget(3), "", "CPM or CPP must be specified", tmDmCtrls(DMVALUEINDEX + 6).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imDMBoxNo = DMVALUEINDEX + 6
            End If
            mDmTestFields = NO
            Exit Function
        End If
    End If
    mDmTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEDIBranch                      *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to EDI    *
'*                      service and process            *
'*                      communication back from EDI    *
'*                      Service                        *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mEDIBranch(ilIndex As Integer)
'
'   ilRet = mEDIBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim ilPdf As Integer
    
    ilRet = gOptionalLookAhead(edcDropDown, lbcEDI(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mEDIBranch = False
        Exit Function
    End If
    If igWinStatus(EDISERVICESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mEDIBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(EDISERVICESLIST)) Then
    '    imDoubleClickName = False
    '    mEDIBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgArfCallType = "A"
    igArfCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    If sgArfName = "PDF EMail" Then
        ReDim tgPdf(0 To UBound(tmPdf)) As PDF
        For ilPdf = 0 To UBound(tmPdf) - 1 Step 1
            tgPdf(ilPdf) = tmPdf(ilPdf)
        Next ilPdf
        sgCommandStr = edcName.Text
        PDFEMailPersonnel.Show vbModal
    Else
        sgCommandStr = slStr
        NmAddr.Show vbModal
    End If
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mEDIBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
        If sgArfName = "PDF EMail" Then
            ReDim tmPdf(0 To UBound(tgPdf)) As PDF
            For ilPdf = 0 To UBound(tgPdf) - 1 Step 1
                tmPdf(ilPdf) = tgPdf(ilPdf)
            Next ilPdf
            bmPDFEMailChgd = True
        End If
'        gSetMenuState True
        lbcEDI(ilIndex).Clear
        smEDICodeTag = ""
        mEDIPop
        If imTerminate Then
            mEDIBranch = False
            Exit Function
        End If
        gFindMatch sgArfName, 1, lbcEDI(ilIndex)
        sgArfName = ""
        If gLastFound(lbcEDI(ilIndex)) > 0 Then
            imChgMode = True
            lbcEDI(ilIndex).ListIndex = gLastFound(lbcEDI(ilIndex))
            edcDropDown.Text = lbcEDI(ilIndex).List(lbcEDI(ilIndex).ListIndex)
            imChgMode = False
            mEDIBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcEDI(ilIndex).ListIndex = 0
            edcDropDown.Text = lbcEDI(ilIndex).List(0)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEDIPop                         *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate EDI service list      *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mEDIPop()
'
'   mEDIPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slEDIC As String
    Dim slEDII As String
    Dim ilEDIC As Integer
    Dim ilEDII As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "A"
    ilOffSet(0) = gFieldOffset("Arf", "ArfType") '2
    ilEDIC = lbcEDI(0).ListIndex
    ilEDII = lbcEDI(1).ListIndex
    If ilEDIC >= 1 Then
        slEDIC = lbcEDI(0).List(ilEDIC)
    End If
    If ilEDII >= 1 Then
        slEDII = lbcEDI(1).List(ilEDII)
    End If
    If lbcEDI(0).ListCount <> lbcEDI(1).ListCount Then
        lbcEDI(0).Clear
    End If
    'ilRet = gIMoveListBox(Advt, lbcEDI(0), lbcEDICode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Advt, lbcEDI(0), tmEDICode(), smEDICodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEDIPopErr
        gCPErrorMsg ilRet, "mEDIPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcEDI(0).AddItem "[None]", 0
        'lbcEDI(0).AddItem "[New]", 0  'Force as first item on list
        lbcEDI(1).Clear
        For ilLoop = lbcEDI(0).ListCount - 1 To 0 Step -1
            lbcEDI(1).AddItem lbcEDI(0).List(ilLoop), 0
        Next ilLoop
        imChgMode = True
        If ilEDIC >= 1 Then
            gFindMatch slEDIC, 1, lbcEDI(0)
            If gLastFound(lbcEDI(0)) >= 1 Then
                lbcEDI(0).ListIndex = gLastFound(lbcEDI(0))
            Else
                lbcEDI(0).ListIndex = -1
            End If
        Else
            lbcEDI(0).ListIndex = ilEDIC
        End If
        If ilEDII >= 1 Then
            gFindMatch slEDII, 1, lbcEDI(1)
            If gLastFound(lbcEDI(1)) >= 1 Then
                lbcEDI(1).ListIndex = gLastFound(lbcEDI(1))
            Else
                lbcEDI(1).ListIndex = -1
            End If
        Else
            lbcEDI(1).ListIndex = ilEDII
        End If
        imChgMode = False
    End If
    Exit Sub
mEDIPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxNoCtrls Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            lbcName.height = gListBoxHeight(lbcName.ListCount, 12)
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 30
            gMoveFormCtrl pbcAdvt, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcName.Left + edcName.Width, edcName.Top
            lbcName.Move edcName.Left, edcName.Top + edcName.height
            gFindMatch edcName.Text, 0, lbcName
            If gLastFound(lbcName) >= 0 Then
                imChgMode = True
                lbcName.ListIndex = gLastFound(lbcName)
                edcName.Text = lbcName.List(lbcName.ListIndex)
                imChgMode = False
            Else
                imChgMode = True
                lbcName.ListIndex = -1
                imChgMode = False
            End If
            edcName.SelStart = 0
            edcName.SelLength = Len(edcName.Text)
            edcName.Visible = True
            cmcDropDown.Visible = True
            edcName.SetFocus
        Case ABBRINDEX 'Abbreviation
            edcAbbr.Width = tmCtrls(ilBoxNo).fBoxW
            edcAbbr.MaxLength = 7
            gMoveFormCtrl pbcAdvt, edcAbbr, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAbbr.Visible = True  'Set visibility
            edcAbbr.SetFocus
        Case POLITICALINDEX   'Political
            If imPolitical < 0 Then
                imPolitical = 1  'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcPolitical.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcPolitical, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcPolitical_Paint
            pbcPolitical.Visible = True
            pbcPolitical.SetFocus
        Case STATEINDEX   'Active/Dormant
            If imState < 0 Then
                imState = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
        Case PRODINDEX 'Product
            mProdPop tmAdf.iCode
            If imTerminate Then
                Exit Sub
            End If
            lbcProd.height = gListBoxHeight(lbcProd.ListCount, 14)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 35  'tgSpf.iAProd
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProd.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            gFindMatch smProduct, 1, lbcProd
            If gLastFound(lbcProd) >= 1 Then
                imChgMode = True
                lbcProd.ListIndex = gLastFound(lbcProd)
                edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                imChgMode = False
            Else
                If smProduct <> "" Then
                    imChgMode = True
                    lbcProd.ListIndex = -1
                    edcDropDown.Text = smProduct
                    imChgMode = False
                Else
                    imChgMode = True
                    If imSelectedIndex > 0 Then
'                        lbcProd.ListIndex = 1   '[None]
                        lbcProd.ListIndex = 0   '[None]
                    Else
                        lbcProd.ListIndex = 0   '[None]
                    End If
                    edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SPERSONINDEX   'Salesperson
            mSPersonPop
            If imTerminate Then
                Exit Sub
            End If
            lbcSPerson.height = gListBoxHeight(lbcSPerson.ListCount, 13)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSPerson.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcSPerson.ListIndex < 0 Then
                lbcSPerson.ListIndex = 1   '[None]
            End If
            If lbcSPerson.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcSPerson.List(lbcSPerson.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AGENCYINDEX   'Agency
            mAgencyPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAgency.height = gListBoxHeight(lbcAgency.ListCount, 13)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 50
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAgency.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcAgency.ListIndex < 0 Then
                lbcAgency.ListIndex = 1   '[None]
            End If
            If lbcAgency.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcAgency.List(lbcAgency.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case REPCODEINDEX 'Rep Agency Code
            If Trim$(edcRepCode.Text) = "" Then
                edcRepCode.Text = gGetNextGPNo(hmAdf, hmAgf, hmSaf)
            End If
            edcRepCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcRepCode.MaxLength = 10
            gMoveFormCtrl pbcAdvt, edcRepCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRepCode.Visible = True  'Set visibility
            edcRepCode.SetFocus
        Case AGYCODEINDEX 'Agency Code
            edcAgyCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcAgyCode.MaxLength = 10
            gMoveFormCtrl pbcAdvt, edcAgyCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAgyCode.Visible = True  'Set visibility
            edcAgyCode.SetFocus
        Case STNCODEINDEX 'Station Agency Code
            edcStnCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcStnCode.MaxLength = 10
            gMoveFormCtrl pbcAdvt, edcStnCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcStnCode.Visible = True  'Set visibility
            edcStnCode.SetFocus
        Case COMPINDEX   'Competitive/Exclusion
            plcCE.Visible = True
            pbcCE.Visible = True
            imCEBoxNo = -1
            pbcCESTab.SetFocus
        Case DEMOINDEX 'Demo Code
            plcDemo.Visible = True  'Set visibility
            pbcDm.Visible = True
            imDMBoxNo = -1
            pbcDmSTab.SetFocus
        Case CREDITAPPROVALINDEX 'Credit Approval
            lbcCreditApproval.height = gListBoxHeight(lbcCreditApproval.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 8
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCreditApproval.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcCreditApproval.Width, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcCreditApproval.ListIndex < 0 Then
                If (tgUrf(0).sChgCrRt <> "I") Then
                    lbcCreditApproval.ListIndex = 0 'Requires Checking
                Else
                    lbcCreditApproval.ListIndex = 1 'Approved
                End If
            End If
            imComboBoxIndex = lbcCreditApproval.ListIndex
            If lbcCreditApproval.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcCreditApproval.List(lbcCreditApproval.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CREDITRESTRINDEX 'Credit restrictions
            lbcCreditRestr.height = gListBoxHeight(lbcCreditRestr.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + tmCtrls(ilBoxNo + 1).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 24
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCreditRestr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcCreditRestr.ListIndex < 0 Then
                lbcCreditRestr.ListIndex = 0   'No Limit
            End If
            imComboBoxIndex = lbcCreditRestr.ListIndex
            If lbcCreditRestr.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcCreditRestr.List(lbcCreditRestr.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CREDITRESTRINDEX + 1 'Limit
            edcCreditLimit.Width = tmCtrls(ilBoxNo).fBoxW
            edcCreditLimit.MaxLength = 10
            gMoveFormCtrl pbcAdvt, edcCreditLimit, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCreditLimit.Visible = True  'Set visibility
            edcCreditLimit.SetFocus
        Case PAYMRATINGINDEX 'Payment rating
            lbcPaymRating.height = gListBoxHeight(lbcPaymRating.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 13
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcPaymRating.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcPaymRating.ListIndex < 0 Then
                lbcPaymRating.ListIndex = 0
            End If
            imComboBoxIndex = lbcPaymRating.ListIndex
            If lbcPaymRating.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcPaymRating.List(lbcPaymRating.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CREDITRATINGINDEX 'Credit Rating
            edcRating.Width = tmCtrls(ilBoxNo).fBoxW
            edcRating.MaxLength = 5
            gMoveFormCtrl pbcAdvt, edcRating, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRating.Visible = True  'Set visibility
            edcRating.SetFocus
        Case RATEONINVINDEX   'Rates on Invoice
            If imRateOnInv < 0 Then
                imRateOnInv = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcRateOnInv.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcRateOnInv, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcRateOnInv_Paint
            pbcRateOnInv.Visible = True
            pbcRateOnInv.SetFocus
        Case ISCIINDEX   'ISCI on Invoices
            If imISCI < 0 Then
                If tgSpf.sAISCI = "Y" Then
                    imISCI = 0    'Yes
                Else
                    imISCI = 1  'No
                End If
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcISCI.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcISCI, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcISCI_Paint
            pbcISCI.Visible = True
            pbcISCI.SetFocus
        Case REPINVINDEX   'Rep Inv
            If imRepInv < 0 Then
                'imRepInv = 0    'Internal
                If (igInternalAdfCount <> 0) And (igInternalAdfCount < UBound(tgCommAdf) / 2) Then
                    imRepInv = 1    'External
                Else
                    imRepInv = 0    'internal
                End If
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcRepInv.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2
            gMoveFormCtrl pbcAdvt, pbcRepInv, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcRepInv_Paint
            pbcRepInv.Visible = True
            pbcRepInv.SetFocus
        Case INVSORTINDEX   'Invoice sorting
            mInvSortPop
            If imTerminate Then
                Exit Sub
            End If
            lbcInvSort.height = gListBoxHeight(lbcInvSort.ListCount, 6)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAdvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX - tmCtrls(ilBoxNo).fBoxW, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcInvSort.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcInvSort.Width, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcInvSort.ListIndex < 0 Then
                lbcInvSort.ListIndex = 1   '[None]
            End If
            If lbcInvSort.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcInvSort.List(lbcInvSort.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PACKAGEINDEX   'Package Invoice Show
            If imPackage < 0 Then
                imPackage = 1   'Time Jim 6/20/97 was 0    'Daypart
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcPackage.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcPackage, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcPackage_Paint
            pbcPackage.Visible = True
            pbcPackage.SetFocus
        Case REPMGINDEX   'Rep MG
            If imRepMG < 0 Then
                imRepMG = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcRepMG.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcRepMG, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcRepMG_Paint
            pbcRepMG.Visible = True
            pbcRepMG.SetFocus
        Case BONUSONINVINDEX   'Rates on Invoice
            If imBonusOnInv < 0 Then
                imBonusOnInv = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcBonusOnInv.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAdvt, pbcBonusOnInv, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcBonusOnInv_Paint
            pbcBonusOnInv.Visible = True
            pbcBonusOnInv.SetFocus
        Case REFIDINDEX 'Ref id L.Bianchi 05/26/2021
            edcRefId.Width = tmCtrls(ilBoxNo).fBoxW
            edcRefId.MaxLength = 36
            gMoveFormCtrl pbcAdvt, edcRefId, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRefId.Visible = True
            edcRefId.SetFocus
        Case DIRECTREFIDINDEX 'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
            edcDirectRefID.Width = tmCtrls(ilBoxNo).fBoxW
            edcDirectRefID.MaxLength = 36
            gMoveFormCtrl pbcAdvt, edcDirectRefID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDirectRefID.Visible = True
            edcDirectRefID.SetFocus
        Case MEGAPHONEADVID
            edcMegaphoneAdvID.Width = tmCtrls(ilBoxNo).fBoxW
            edcMegaphoneAdvID.MaxLength = 36
            gMoveFormCtrl pbcAdvt, edcMegaphoneAdvID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcMegaphoneAdvID.Visible = True
            edcMegaphoneAdvID.SetFocus
        Case CRMIDINDEX ' JD 09-22-22
'            edcCRMID.Text = ""
'            If tmAdfx.lCrmId <> 0 Then
'                edcCRMID.Text = CStr(tmAdfx.lCrmId)
'            End If
            edcCRMID.Width = tmCtrls(CRMIDINDEX).fBoxW
            edcCRMID.MaxLength = 10
            
            ' gMoveFormCtrl pbcAdvt, edcCRMID, tmCtrls(CRMIDINDEX).fBoxX, tmCtrls(CRMIDINDEX).fBoxY
            
            edcCRMID.Move _
                tmCtrls(CRMIDINDEX).fBoxX + 260, _
                tmCtrls(CRMIDINDEX).fBoxY + 720, _
                tmCtrls(CRMIDINDEX).fBoxW - 30, _
                fgBoxStH - 135
                
'            edcCRMID.Move _
'                tmCtrls(CRMIDINDEX).fBoxX + 220, _
'                tmCtrls(CRMIDINDEX).fBoxY + 720, _
'                tmCtrls(CRMIDINDEX).fBoxW - 1830, _
'                fgBoxStH - 135
            edcCRMID.Visible = True  'Set visibility
            edcCRMID.SetFocus
            
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcDirect, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case ADDRIDINDEX 'Address ID
            edcAddrID.Width = tmCtrls(ilBoxNo).fBoxW
            edcAddrID.MaxLength = 9
            gMoveFormCtrl pbcDirect, edcAddrID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddrID.Visible = True  'Set visibility
            edcAddrID.SetFocus
        Case BUYERINDEX 'Product
            mBuyerPop tmAdf.iCode, "", -1
            If imTerminate Then
                Exit Sub
            End If
            lbcBuyer.height = gListBoxHeight(lbcBuyer.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 64
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcBuyer.Move edcDropDown.Left, edcDropDown.Top - lbcBuyer.height
            gFindMatch smBuyer, 1, lbcBuyer
            If gLastFound(lbcBuyer) >= 1 Then
                imChgMode = True
                lbcBuyer.ListIndex = gLastFound(lbcBuyer)
                edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
                imChgMode = False
            Else
                If smBuyer <> "" Then
                    imChgMode = True
                    lbcBuyer.ListIndex = -1
                    edcDropDown.Text = smBuyer
                    imChgMode = False
                Else
                    imChgMode = True
                    If imSelectedIndex > 0 Then
'                        lbcProd.ListIndex = 1   '[None]
                        lbcBuyer.ListIndex = 1   '[None]
                    Else
                        lbcBuyer.ListIndex = 1   '[None]
                    End If
                    edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PAYABLEINDEX 'Product
            mPayablePop tmAdf.iCode, "", -1
            If imTerminate Then
                Exit Sub
            End If
            lbcPayable.height = gListBoxHeight(lbcPayable.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 64
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcPayable.Move edcDropDown.Left, edcDropDown.Top - lbcPayable.height
            gFindMatch smPayable, 1, lbcPayable
            If gLastFound(lbcPayable) >= 1 Then
                imChgMode = True
                lbcPayable.ListIndex = gLastFound(lbcPayable)
                edcDropDown.Text = lbcPayable.List(lbcPayable.ListIndex)
                imChgMode = False
            Else
                If smPayable <> "" Then
                    imChgMode = True
                    lbcPayable.ListIndex = -1
                    edcDropDown.Text = smPayable
                    imChgMode = False
                Else
                    imChgMode = True
                    If imSelectedIndex > 0 Then
'                        lbcProd.ListIndex = 1   '[None]
                        lbcPayable.ListIndex = 1   '[None]
                    Else
                        lbcPayable.ListIndex = 1   '[None]
                    End If
                    edcDropDown.Text = lbcPayable.List(lbcPayable.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case LKBOXINDEX 'Lock box
            mLkBoxPop
            If imTerminate Then
                Exit Sub
            End If
            lbcLkBox.height = gListBoxHeight(lbcLkBox.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLkBox.Move edcDropDown.Left, edcDropDown.Top - lbcLkBox.height
            imChgMode = True
            If lbcLkBox.ListIndex < 0 Then
                If lbcLkBox.ListCount > 2 Then
                    lbcLkBox.ListIndex = 2  'First one
                Else
                    lbcLkBox.ListIndex = 1   '[None]
                End If
            End If
            If lbcLkBox.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcLkBox.List(lbcLkBox.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case EDICINDEX   'EDI service for Contracts
            mEDIPop
            If imTerminate Then
                Exit Sub
            End If
            lbcEDI(0).height = gListBoxHeight(lbcEDI(0).ListCount, 7)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcEDI(0).Move edcDropDown.Left, edcDropDown.Top - lbcEDI(0).height
            imChgMode = True
            If lbcEDI(0).ListIndex < 0 Then
                lbcEDI(0).ListIndex = 0   '[None]
            End If
            If lbcEDI(0).ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcEDI(0).List(lbcEDI(0).ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case EDIIINDEX   'EDI service for Contracts
            mEDIPop
            If imTerminate Then
                Exit Sub
            End If
            lbcEDI(1).height = gListBoxHeight(lbcEDI(1).ListCount, 7)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcEDI(1).Move edcDropDown.Left, edcDropDown.Top - lbcEDI(1).height
            imChgMode = True
            If lbcEDI(1).ListIndex < 0 Then
                lbcEDI(1).ListIndex = 0   '[None]
            End If
            If lbcEDI(1).ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcEDI(1).List(lbcEDI(1).ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
'        Case PRTSTYLEINDEX   'Print style
'            If imPrtStyle < 0 Then
'                imPrtStyle = 1    'Narrow
'                tmCtrls(ilBoxNo).iChg = True
'                mSetCommands
'            End If
'            pbcPrtStyle.Width = tmCtrls(ilBoxNo).fBoxW
'            gMoveFormCtrl pbcDirect, pbcPrtStyle, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
'            pbcPrtStyle_Paint
'            pbcPrtStyle.Visible = True
'            pbcPrtStyle.SetFocus
        Case TERMSINDEX   'Terms
            mTermsPop
            If imTerminate Then
                Exit Sub
            End If
            lbcTerms.height = gListBoxHeight(lbcTerms.ListCount, 6)
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2 - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTerms.Move edcDropDown.Left, edcDropDown.Top - lbcTerms.height, edcDropDown.Width + cmcDropDown.Width
            imChgMode = True
            If lbcTerms.ListIndex < 0 Then
                lbcTerms.ListIndex = 1   'Default
            End If
            If lbcTerms.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTerms.List(lbcTerms.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TAXINDEX   'Tax
            lbcTax.height = gListBoxHeight(lbcTax.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 0
            gMoveFormCtrl pbcDirect, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTax.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcTax.Width, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcTax.ListIndex < 0 Then
                lbcTax.ListIndex = 0   'No
            End If
            imComboBoxIndex = lbcTax.ListIndex
            If lbcTax.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
            
         Case SUPPRESSNETINDEX   'Suppress Net Amount for Trade Invoices  ' TTP 10622 - 2023-03-08 JJB
            If imSuppressNet < 0 Then
                imSuppressNet = -1
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            
            pbcSuppressNet.Width = 2700 'tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDirect, pbcSuppressNet, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcSuppressNet_Paint
            pbcSuppressNet.Visible = True
            pbcSuppressNet.SetFocus
    End Select
    
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExclBranch                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      exclusion and process          *
'*                      communication back from        *
'*                      exclusion                      *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mExclBranch(ilIndex As Integer) As Integer
'
'   ilRet = mExclBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcCEDropDown, lbcExcl(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcCEDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mExclBranch = False
        Exit Function
    End If
    If igWinStatus(EXCLUSIONSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mExclBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(EXCLUSIONSLIST)) Then
    '    imDoubleClickName = False
    '    mExclBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "X"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcCEDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mExclBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcExcl(ilIndex).Clear
        sgExclCodeTag = ""
        sgExclMnfStamp = ""
        mExclPop
        If imTerminate Then
            mExclBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcExcl(ilIndex)
        sgMNmName = ""
        If gLastFound(lbcExcl(ilIndex)) > 0 Then
            imChgMode = True
            lbcExcl(ilIndex).ListIndex = gLastFound(lbcExcl(ilIndex))
            edcCEDropDown.Text = lbcExcl(ilIndex).List(lbcExcl(ilIndex).ListIndex)
            imChgMode = False
            mExclBranch = False
            mCESetChg imCEBoxNo
        Else
            imChgMode = True
            lbcExcl(ilIndex).ListIndex = 1
            edcCEDropDown.Text = lbcExcl(ilIndex).List(1)
            imChgMode = False
            mCESetChg imCEBoxNo
            edcCEDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCEEnableBox imCEBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCEEnableBox imCEBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExclPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate exclusion list        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mExclPop()
'
'   mExclPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slExcl(0 To 1) As String      'Exclusion name, saved to determine if changed
    ReDim ilExcl(0 To 1) As Integer      'Exclusion name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "X"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilExcl(0) = lbcExcl(0).ListIndex
    ilExcl(1) = lbcExcl(1).ListIndex
    If ilExcl(0) > 1 Then
        slExcl(0) = lbcExcl(0).List(ilExcl(0))
    End If
    If ilExcl(1) > 1 Then
        slExcl(1) = lbcExcl(1).List(ilExcl(1))
    End If
    If lbcExcl(0).ListCount <> lbcExcl(1).ListCount Then
        lbcExcl(0).Clear
    End If
    'ilRet = gIMoveListBox(Advt, lbcExcl(0), lbcExclCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Advt, lbcExcl(0), tgExclCode(), sgExclCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mExclPopErr
        gCPErrorMsg ilRet, "mExclPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcExcl(0).AddItem "[None]", 0
        lbcExcl(0).AddItem "[New]", 0  'Force as first item on list
        lbcExcl(1).Clear
        For ilLoop = lbcExcl(0).ListCount - 1 To 0 Step -1
            lbcExcl(1).AddItem lbcExcl(0).List(ilLoop), 0
        Next ilLoop
        imChgMode = True
        If ilExcl(0) > 1 Then
            gFindMatch slExcl(0), 2, lbcExcl(0)
            If gLastFound(lbcExcl(0)) > 1 Then
                lbcExcl(0).ListIndex = gLastFound(lbcExcl(0))
            Else
                lbcExcl(0).ListIndex = -1
            End If
        Else
            lbcExcl(0).ListIndex = ilExcl(0)
        End If
        If ilExcl(1) > 1 Then
            gFindMatch slExcl(1), 2, lbcExcl(1)
            If gLastFound(lbcExcl(1)) > 1 Then
                lbcExcl(1).ListIndex = gLastFound(lbcExcl(1))
            Else
                lbcExcl(1).ListIndex = -1
            End If
        Else
            lbcExcl(1).ListIndex = ilExcl(1)
        End If
        imChgMode = False
    End If
    Exit Sub
mExclPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
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
    Dim ilRet As Integer    'Return Status
    Dim ilSlfCode As Integer
    Dim ilSlf As Integer

'    ilRet = GetOffSetForInt(tmAdf, tmAdf.iSlfCode)
'    ilRet = GetOffSetForStr(tmAdf, tmAdf.sName)
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBDmCtrls = 1
    imLBDmCtrls = 1
    'DoEvents
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    ilSlfCode = tgUrf(0).iSlfCode
    If (tgUrf(0).iSlfCode > 0) Then
        For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgMSlf(ilSlf).iCode = tgUrf(0).iSlfCode Then
                If StrComp(tgMSlf(ilSlf).sJobTitle, "S", 1) <> 0 Then
                    ilSlfCode = 0
                End If
                Exit For
            End If
        Next ilSlf
    End If
    If (ilSlfCode > 0) Or (igAdvtCallSource = CALLSOURCECONTRACT) Or (tgUrf(0).iRemoteUserID > 0) Then
        imShortForm = True
    Else
        imShortForm = False
    End If
    mInitBox
    Advt.height = cmcCancel.Top + 5 * cmcCancel.height / 3
    gCenterStdAlone Advt
    'Advt.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture

    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imAdfRecLen = Len(tmAdf)  'Get and save ADF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDMBoxNo = -1
    imCEBoxNo = -1
    imDirProcess = -1
    imBillTabDir = 0
    imEdcChgMode = False
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) And (tgSpf.iBTax(1) = 0) Then
    'If ((Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA) Or ((Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA) Then
    '    imTaxDefined = True
    'Else
    '    imTaxDefined = False
    'End If
    lbcCreditRestr.AddItem "No Restrictions"
    lbcCreditRestr.AddItem "Credit Limit"
    lbcCreditRestr.AddItem "Cash in Advance- Weekly"
    lbcCreditRestr.AddItem "Cash in Advance- Monthly"
    lbcCreditRestr.AddItem "Cash in Advance- Contract"
    lbcCreditRestr.AddItem "Prohibit New Orders"
    lbcPaymRating.AddItem "Quick Pay"
    lbcPaymRating.AddItem "Normal Pay"
    lbcPaymRating.AddItem "Slow Pay"
    lbcPaymRating.AddItem "Difficult"
    lbcPaymRating.AddItem "In Collection"
    lbcCreditApproval.AddItem "Requires Checking"
    lbcCreditApproval.AddItem "Approved"
    lbcCreditApproval.AddItem "Denied"
    'lbcTax.AddItem "Tax 1-No   Tax 2-No"
    'lbcTax.AddItem "Tax 1-Yes  Tax 2-No"
    'lbcTax.AddItem "Tax 1-No   Tax 2-Yes"
    'lbcTax.AddItem "Tax 1-Yes  Tax 2-Yes"
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Then
        If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
            ilRet = gPopTaxRateBox(True, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
            imTaxDefined = True
        ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
            ilRet = gPopTaxRateBox(False, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
            imTaxDefined = True
        Else
            imTaxDefined = False
        End If
    Else
        imTaxDefined = False
    End If
    If Not imTaxDefined Then
        ReDim tmTaxSortCode(0 To 0) As SORTCODE
    End If
    lbcTax.AddItem "[None]", 0

    mDemoPop
'    cbcPriceType.AddItem "N/A"
'    cbcPriceType.AddItem "CPM"
'    cbcPriceType.AddItem "CPP"
    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "PRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PRF.Btr)", Advt
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmPnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPnf, "", sgDBPath & "PNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PNF.Btr)", Advt
    On Error GoTo 0
    imPnfRecLen = Len(tmBPnf)
    
    hmPDF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPDF, "", sgDBPath & "PDF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PDF.Btr)", Advt
    On Error GoTo 0
    ReDim tmPdf(0 To 0) As PDF
    imPdfRecLen = Len(tmPdf(0))
    
    hmAxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAxf, "", sgDBPath & "AXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: AXF.Btr)", Advt
    On Error GoTo 0
    imAxfRecLen = Len(tmAxf)
    
    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "SAF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SAF.Btr)", Advt
    On Error GoTo 0
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "AGF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: AGF.Btr)", Advt
    On Error GoTo 0
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "DSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: DSF.Btr)", Advt
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    hmAdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", Advt
    On Error GoTo 0
    lbcSPerson.Clear 'Force list box to be populated
    mSPersonPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcAgency.Clear 'Force list box to be populated
    mAgencyPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcComp(0).Clear 'Force list box to be populated
    lbcComp(1).Clear 'Force list box to be populated
    mCompPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcExcl(0).Clear 'Force list box to be populated
    lbcExcl(1).Clear 'Force list box to be populated
    mExclPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcLkBox.Clear 'Force list box to be populated
    mLkBoxPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcEDI(0).Clear 'Force list box to be populated
    lbcEDI(1).Clear
    mEDIPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcInvSort.Clear 'Force list box to be populated
    mInvSortPop
    If imTerminate Then
        Exit Sub
    End If
    lbcTerms.Clear 'Force list box to be populated
    mTermsPop
    If imTerminate Then
        Exit Sub
    End If
    mEnableSplitCue
    Screen.MousePointer = vbHourglass  'Wait
'    gCenterModalForm Advt
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    Screen.MousePointer = vbHourglass  'Wait
    rbcBill(0).Value = True
    imCreditRestrFirst = True
    imPaymRatingFirst = True
'    imPrtStyleFirst = True
    imScriptFirst = True
    imPriceTypeFirst = True
    imTaxFirst = True
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
'*             Created:6/1/93        By:D. LeVine      *
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
    Dim llMax As Long
    Dim llShortMax As Long
    Dim ilLoop As Integer
    Dim ilGap As Integer
    
    flTextHeight = pbcAdvt.TextHeight("1") - 35
    If imShortForm Then
        pbcAdvt.height = 375
        pbcAdvt.Width = 4590
        pbcDirect.height = 1065
    Else
        pbcAdvt.height = 1790 'L.Bianchi 05/26/2021
        pbcAdvt.Width = 8490
        pbcDirect.height = 2900
    End If
    'Position panel and picture areas with panel
    plcAdvt.Move 210, 540, pbcAdvt.Width + fgPanelAdj, pbcAdvt.height + fgPanelAdj
    pbcAdvt.Move plcAdvt.Left + fgBevelX, plcAdvt.Top + fgBevelY
    'plcDirect.Move plcAdvt.Left, 2400, pbcDirect.Width + fgPanelAdj, pbcDirect.height + fgPanelAdj
    plcDirect.Move plcAdvt.Left, 2775, pbcDirect.Width + fgPanelAdj, pbcDirect.height + fgPanelAdj 'L.Bianchi 05/26/2021
    pbcDirect.Move plcDirect.Left + fgBevelX, plcDirect.Top + fgBevelY
    plcDirect.Visible = False
    pbcDirect.Visible = False
    plcNotDirect.Move plcAdvt.Left, plcDirect.Top + plcDirect.height - pbcNotDirect.height - fgPanelAdj, pbcDirect.Width + fgPanelAdj, pbcNotDirect.height + fgPanelAdj
    pbcNotDirect.Move plcNotDirect.Left + fgBevelX, plcNotDirect.Top + fgBevelY
    If imShortForm Then
        plcNotDirect.Visible = False
        pbcNotDirect.Visible = False
    Else
        plcNotDirect.Visible = True
        pbcNotDirect.Visible = True
    End If
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2985, fgBoxStH
    'Abbreviation
    gSetCtrl tmCtrls(ABBRINDEX), 3030, tmCtrls(NAMEINDEX).fBoxY, 930, fgBoxStH
    'Political
    gSetCtrl tmCtrls(POLITICALINDEX), 3975, tmCtrls(NAMEINDEX).fBoxY, 600, fgBoxStH
    'State
    gSetCtrl tmCtrls(STATEINDEX), 4590, tmCtrls(NAMEINDEX).fBoxY, 1065, fgBoxStH
    'Product
    gSetCtrl tmCtrls(PRODINDEX), 5670, tmCtrls(NAMEINDEX).fBoxY, 2800, fgBoxStH
    tmCtrls(PRODINDEX).iReq = False
    'Salesperson
    gSetCtrl tmCtrls(SPERSONINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2445, fgBoxStH
    tmCtrls(SPERSONINDEX).iReq = False
    'Agency
    gSetCtrl tmCtrls(AGENCYINDEX), 2490, tmCtrls(SPERSONINDEX).fBoxY, 2445, fgBoxStH
    tmCtrls(AGENCYINDEX).iReq = False
    'Rep Advertiser Code
    gSetCtrl tmCtrls(REPCODEINDEX), 4950, tmCtrls(SPERSONINDEX).fBoxY, 1020, fgBoxStH
    tmCtrls(REPCODEINDEX).iReq = False  'If changed to True, test tgSpf.sARepCodes and if not using leave as false
    'Agenct Advertiser Code
    gSetCtrl tmCtrls(AGYCODEINDEX), 5985, tmCtrls(SPERSONINDEX).fBoxY, 1230, fgBoxStH
    tmCtrls(AGYCODEINDEX).iReq = False  'If changed to True, test tgSpf.sAAgyCodes and if not using leave as false
    'Station Agency Code
    gSetCtrl tmCtrls(STNCODEINDEX), 7235, tmCtrls(SPERSONINDEX).fBoxY, 1230, fgBoxStH
    tmCtrls(STNCODEINDEX).iReq = False  'If changed to True, test tgSpf.sAStnCodes and if not using leave as false
    'Competitive/Exclusions
    gSetCtrl tmCtrls(COMPINDEX), 30, tmCtrls(SPERSONINDEX).fBoxY + fgStDeltaY, 4215, fgBoxStH   'was 2805
    tmCtrls(COMPINDEX).iReq = False
        'Competitive/Program Exclusion
        gMoveCtrl pbcAdvt, plcCE, tmCtrls(COMPINDEX).fBoxX, tmCtrls(COMPINDEX).fBoxY
        pbcCE.Move plcCE.Left + 150, plcCE.Top + 315
        gSetCtrl tmCECtrls(CECOMPINDEX), 30, 30, 2340, fgBoxStH
        tmCECtrls(CECOMPINDEX).iReq = False
        gSetCtrl tmCECtrls(CECOMPINDEX + 1), 2385, tmCECtrls(CECOMPINDEX).fBoxY, 2340, fgBoxStH
        tmCECtrls(CECOMPINDEX + 1).iReq = False
        gSetCtrl tmCECtrls(CEEXCLINDEX), 30, tmCECtrls(CECOMPINDEX).fBoxY + fgStDeltaY, 2340, fgBoxStH
        tmCECtrls(CEEXCLINDEX).iReq = False  'If changed to True, test tgSpf.sAEDIC and if not using leave as false
        gSetCtrl tmCECtrls(CEEXCLINDEX + 1), 2385, tmCECtrls(CEEXCLINDEX).fBoxY, 2340, fgBoxStH
        tmCECtrls(CEEXCLINDEX + 1).iReq = False  'If changed to True, test tgSpf.sAEDII and if not using leave as false
'    gSetCtrl tmCtrls(COMPINDEX + 1), 2130, tmCtrls(COMPINDEX).fBoxY, 2115, fgBoxStH
'    gSetCtrl tmCtrls(COMPINDEX + 1), tmCtrls(COMPINDEX).fBoxX + lbcComp(0).Width - 15, tmCtrls(COMPINDEX).fBoxY, 2115, fgBoxStH
'    tmCtrls(COMPINDEX + 1).iReq = False
    'Demo
    gSetCtrl tmCtrls(DEMOINDEX), 4260, tmCtrls(COMPINDEX).fBoxY, 2670, fgBoxStH     '2970, fgBoxStH
    tmCtrls(DEMOINDEX).iReq = False
    gSetCtrl tmCtrls(BKOUTPOOLINDEX), 6945, tmCtrls(COMPINDEX).fBoxY, 280, fgBoxStH
    tmCtrls(BKOUTPOOLINDEX).iReq = False
        'Price type/Demo/Target
        gMoveCtrl pbcAdvt, plcDemo, tmCtrls(DEMOINDEX).fBoxX, tmCtrls(DEMOINDEX).fBoxY
        pbcDm.Move plcDemo.Left + 150, plcDemo.Top + 285
        gSetCtrl tmDmCtrls(DMPRICETYPEINDEX), 30, 30, 2835, fgBoxStH
        tmDmCtrls(DMPRICETYPEINDEX).iReq = False
        gSetCtrl tmDmCtrls(DMDEMOINDEX), 30, tmDmCtrls(DMPRICETYPEINDEX).fBoxY + fgStDeltaY, 1410, fgBoxStH
        tmDmCtrls(DMDEMOINDEX).iReq = False
        gSetCtrl tmDmCtrls(DMVALUEINDEX), 1455, tmDmCtrls(DMDEMOINDEX).fBoxY, 1410, fgBoxStH
        tmDmCtrls(DMVALUEINDEX).iReq = False
        gSetCtrl tmDmCtrls(DMDEMOINDEX + 2), 30, tmDmCtrls(DMDEMOINDEX).fBoxY + fgStDeltaY, 1410, fgBoxStH
        tmDmCtrls(DMDEMOINDEX + 2).iReq = False
        gSetCtrl tmDmCtrls(DMVALUEINDEX + 2), 1455, tmDmCtrls(DMDEMOINDEX + 2).fBoxY, 1410, fgBoxStH
        tmDmCtrls(DMVALUEINDEX + 2).iReq = False
        gSetCtrl tmDmCtrls(DMDEMOINDEX + 4), 30, tmDmCtrls(DMDEMOINDEX + 2).fBoxY + fgStDeltaY, 1410, fgBoxStH
        tmDmCtrls(DMDEMOINDEX + 4).iReq = False
        gSetCtrl tmDmCtrls(DMVALUEINDEX + 4), 1455, tmDmCtrls(DMDEMOINDEX + 4).fBoxY, 1410, fgBoxStH
        tmDmCtrls(DMVALUEINDEX + 4).iReq = False
        gSetCtrl tmDmCtrls(DMDEMOINDEX + 6), 30, tmDmCtrls(DMDEMOINDEX + 4).fBoxY + fgStDeltaY, 1410, fgBoxStH
        tmDmCtrls(DMDEMOINDEX + 6).iReq = False
        gSetCtrl tmDmCtrls(DMVALUEINDEX + 6), 1455, tmDmCtrls(DMDEMOINDEX + 6).fBoxY, 1410, fgBoxStH
        tmDmCtrls(DMVALUEINDEX + 6).iReq = False
    'Credit Approval
    gSetCtrl tmCtrls(CREDITAPPROVALINDEX), 7245, tmCtrls(COMPINDEX).fBoxY, 1215, fgBoxStH
    If tgUrf(0).sChgCrRt <> "I" Then
        tmCtrls(CREDITAPPROVALINDEX).iReq = False
    End If
    'Credit restriction
    gSetCtrl tmCtrls(CREDITRESTRINDEX), 30, tmCtrls(COMPINDEX).fBoxY + fgStDeltaY, 950, fgBoxStH
    If tgUrf(0).sCredit <> "I" Then
        tmCtrls(CREDITRESTRINDEX).iReq = False
    End If
    'gSetCtrl tmCtrls(CREDITRESTRINDEX + 1), 965, tmCtrls(CREDITRESTRINDEX).fBoxY, 950, fgBoxStH
    gSetCtrl tmCtrls(CREDITRESTRINDEX + 1), 995, tmCtrls(CREDITRESTRINDEX).fBoxY, 950, fgBoxStH
'    gSetCtrl tmCtrls(CREDITRESTRINDEX + 1), tmCtrls(CREDITRESTRINDEX).fBoxX + lbcCreditRestr.Width - 15, tmCtrls(CREDITRESTRINDEX).fBoxY, 1065, fgBoxStH
    tmCtrls(CREDITRESTRINDEX + 1).iReq = False
    'Payment Rating
    gSetCtrl tmCtrls(PAYMRATINGINDEX), 1950, tmCtrls(CREDITRESTRINDEX).fBoxY, 1080, fgBoxStH
    If tgUrf(0).sPayRate <> "I" Then
        tmCtrls(PAYMRATINGINDEX).iReq = False
    End If
    'Credit Rating
    gSetCtrl tmCtrls(CREDITRATINGINDEX), 3045, tmCtrls(CREDITRESTRINDEX).fBoxY, 870, fgBoxStH
    tmCtrls(CREDITRATINGINDEX).iReq = False
    'Rep Inv
    gSetCtrl tmCtrls(REPINVINDEX), 3930, tmCtrls(CREDITRESTRINDEX).fBoxY, 495, fgBoxStH
    tmCtrls(REPINVINDEX).iReq = False
    'Invoice sorting
    gSetCtrl tmCtrls(INVSORTINDEX), 4440, tmCtrls(CREDITRESTRINDEX).fBoxY, 750, fgBoxStH
    tmCtrls(INVSORTINDEX).iReq = False
    'Rate on Invoice
    gSetCtrl tmCtrls(RATEONINVINDEX), 5865, tmCtrls(CREDITRESTRINDEX).fBoxY, 375, fgBoxStH
    tmCtrls(RATEONINVINDEX).iReq = False
    'ISCI on invoices
    gSetCtrl tmCtrls(ISCIINDEX), 6255, tmCtrls(CREDITRESTRINDEX).fBoxY, 375, fgBoxStH
    If (tgSpf.sAISCI = "A") Or (tgSpf.sAISCI = "X") Then
        tmCtrls(ISCIINDEX).iReq = False
    End If
    'Package Invoice Show
    'gSetCtrl tmCtrls(PACKAGEINDEX), 6645, tmCtrls(CREDITRESTRINDEX).fBoxY, 795, fgBoxStH
    gSetCtrl tmCtrls(PACKAGEINDEX), 6660, tmCtrls(CREDITRESTRINDEX).fBoxY, 795, fgBoxStH
    tmCtrls(PACKAGEINDEX).iReq = False
    'Rep MG
    'gSetCtrl tmCtrls(REPMGINDEX), 7455, tmCtrls(CREDITRESTRINDEX).fBoxY, 555, fgBoxStH
    gSetCtrl tmCtrls(REPMGINDEX), 7470, tmCtrls(CREDITRESTRINDEX).fBoxY, 565, fgBoxStH
    'If (sgMktClusterDef <> "Y") And (sgMktRepDef <> "Y") Then
        tmCtrls(REPMGINDEX).iReq = False
    'End If
    'Rate on Invoice
    'gSetCtrl tmCtrls(BONUSONINVINDEX), 8025, tmCtrls(CREDITRESTRINDEX).fBoxY, 450, fgBoxStH
    gSetCtrl tmCtrls(BONUSONINVINDEX), 8040, tmCtrls(CREDITRESTRINDEX).fBoxY, 435, fgBoxStH
    tmCtrls(BONUSONINVINDEX).iReq = False

' line 5
    'RefId L.Bianchi 05/26/2021
    gSetCtrl tmCtrls(REFIDINDEX), 30, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 2100, fgBoxStH
    tmCtrls(REFIDINDEX).iReq = False

    'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    gSetCtrl tmCtrls(DIRECTREFIDINDEX), 2145, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 2000, fgBoxStH
    tmCtrls(DIRECTREFIDINDEX).iReq = False
    
    'JJB - 04/15/24 - SOW Megaphone Phase I
    gSetCtrl tmCtrls(MEGAPHONEADVID), 4155, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 2200, fgBoxStH
    tmCtrls(MEGAPHONEADVID).iReq = False

    ' JD 09-19-22
    gSetCtrl tmCtrls(CRMIDINDEX), 6370, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 2090, fgBoxStH
    tmCtrls(CRMIDINDEX).iReq = False
   
   
'
'    'NSF Checkes
'    gSetCtrl tmNDTCtrls(NSFCHKSINDEX), 30, 30 + 350, 1515, fgBoxStH
'    'Date Last Billed
'    gSetCtrl tmNDTCtrls(DATELSTINVINDEX), 1560, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
'    'Date Last Payment
'    gSetCtrl tmNDTCtrls(DATELSTPAYMINDEX), 3090, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
'    'Avg # days to pay
'    gSetCtrl tmNDTCtrls(AVGTOPAYINDEX), 4620, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
'    '# Days to pay lst payment
'    gSetCtrl tmNDTCtrls(LSTTOPAYINDEX), 6150, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 2325, fgBoxStH
    
    
' line 1
    'Contract Address
'    gSetCtrl tmCtrls(CADDRINDEX), 30, tmCtrls(CREDITRESTRINDEX).fBoxY + 2 * fgStDeltaY, 4215, fgBoxStH
    gSetCtrl tmCtrls(CADDRINDEX), 30, 30, 4215, fgBoxStH
    tmCtrls(CADDRINDEX).iReq = False
    gSetCtrl tmCtrls(CADDRINDEX + 1), 30, tmCtrls(CADDRINDEX).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(CADDRINDEX + 2), 30, tmCtrls(CADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 2).iReq = False
    'Billing Address
'    gSetCtrl tmCtrls(BADDRINDEX), 4260, tmCtrls(CREDITRESTRINDEX).fBoxY + 2 * fgStDeltaY, 4215, fgBoxStH
    'gSetCtrl tmCtrls(BADDRINDEX), 4260, 30, 4215, fgBoxStH
    gSetCtrl tmCtrls(BADDRINDEX), 4260, 30, 4230, fgBoxStH
    tmCtrls(BADDRINDEX).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 1), 4260, tmCtrls(BADDRINDEX).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 2), 4260, tmCtrls(BADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 2).iReq = False
    'Address ID
    gSetCtrl tmCtrls(ADDRIDINDEX), 30, tmCtrls(CADDRINDEX).fBoxY + fgAddDeltaY, 1515, fgBoxStH
    tmCtrls(ADDRIDINDEX).iReq = False
    'Buyer Name
    gSetCtrl tmCtrls(BUYERINDEX), 1560, tmCtrls(ADDRIDINDEX).fBoxY, 6915, fgBoxStH
    tmCtrls(BUYERINDEX).iReq = False
    ''Phone
    'gSetCtrl tmCtrls(BUYERPHONEINDEX), 2235, tmCtrls(BUYERINDEX).fBoxY, 2010, fgBoxStH
    'tmCtrls(BUYERPHONEINDEX).iReq = False
    ''Fax
    'gSetCtrl tmCtrls(BUYERFAXINDEX), 4260, tmCtrls(BUYERINDEX).fBoxY, 1500, fgBoxStH
    'tmCtrls(BUYERFAXINDEX).iReq = False
    'Payable Contact Name
    gSetCtrl tmCtrls(PAYABLEINDEX), 30, tmCtrls(BUYERINDEX).fBoxY + fgStDeltaY, 5730, fgBoxStH
    tmCtrls(PAYABLEINDEX).iReq = False
    ''Phone
    'gSetCtrl tmCtrls(PAYABLEPHONEINDEX), 2235, tmCtrls(PAYABLEINDEX).fBoxY, 2010, fgBoxStH
    'tmCtrls(PAYABLEPHONEINDEX).iReq = False
    ''Fax
    'gSetCtrl tmCtrls(PAYABLEFAXINDEX), 4260, tmCtrls(PAYABLEINDEX).fBoxY, 1500, fgBoxStH
    'tmCtrls(PAYABLEFAXINDEX).iReq = False
    'LkBox
    gSetCtrl tmCtrls(LKBOXINDEX), 5775, tmCtrls(PAYABLEINDEX).fBoxY, 2700, fgBoxStH
    tmCtrls(LKBOXINDEX).iReq = False
    'EDI Service Contract
    gSetCtrl tmCtrls(EDICINDEX), 30, tmCtrls(PAYABLEINDEX).fBoxY + fgStDeltaY, 2190, fgBoxStH
    tmCtrls(EDICINDEX).iReq = False
    'EDI Service Contract
    gSetCtrl tmCtrls(EDIIINDEX), 2235, tmCtrls(EDICINDEX).fBoxY, 2010, fgBoxStH
    tmCtrls(EDIIINDEX).iReq = False
'    'Contract Print Style
'    gSetCtrl tmCtrls(PRTSTYLEINDEX), 4260, tmCtrls(EDICINDEX).fBoxY, 1500, fgBoxStH
'    If (tgSpf.sAPrtStyle = "W") Or (tgSpf.sAPrtStyle = "N") Then
'        tmCtrls(PRTSTYLEINDEX).iReq = False
'    Else
'        tmCtrls(PRTSTYLEINDEX).iReq = False
'    End If
    gSetCtrl tmCtrls(TERMSINDEX), 4260, tmCtrls(EDICINDEX).fBoxY, 1500, fgBoxStH
    tmCtrls(TERMSINDEX).iReq = False
    'Sales Tax
    gSetCtrl tmCtrls(TAXINDEX), 5775, tmCtrls(EDICINDEX).fBoxY, 2700, fgBoxStH
    If Not imTaxDefined Then
        tmCtrls(TAXINDEX).iReq = False
    Else
        tmCtrls(TAXINDEX).iReq = False
    End If
    
     'Suppress Net Amount For Trade Invoices  ' TTP 10622 - 2023-03-08 JJB
    gSetCtrl tmCtrls(SUPPRESSNETINDEX), 30, tmCtrls(EDICINDEX).fBoxY + 350, 3825, fgBoxStH
    tmCtrls(SUPPRESSNETINDEX).iReq = False
    'Unused
    gSetCtrl tmCtrls(UNUSEDINDEX), 50, tmCtrls(EDICINDEX).fBoxY + 350, 1400, fgBoxStH
    tmCtrls(UNUSEDINDEX).iReq = False
    
ex:
    If imShortForm Then
        tmCtrls(CREDITAPPROVALINDEX).iReq = False
        tmCtrls(CREDITRESTRINDEX).iReq = False
        tmCtrls(PAYMRATINGINDEX).iReq = False
        tmCtrls(ISCIINDEX).iReq = False
    End If
    
    'Pct 90
    'gSetCtrl tmNDTCtrls(PCT90INDEX), 30, 30, 1515, fgBoxStH
    gSetCtrl tmNDTCtrls(PCT90INDEX), 30, 30, 1515, fgBoxStH
    'Currect A/R
    gSetCtrl tmNDTCtrls(CURRARINDEX), 1560, tmNDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Unbilled
    gSetCtrl tmNDTCtrls(UNBILLEDINDEX), 3090, tmNDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Hi Credit
    gSetCtrl tmNDTCtrls(HICREDITINDEX), 4620, tmNDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Total Gross
    gSetCtrl tmNDTCtrls(TOTALGROSSINDEX), 6150, tmNDTCtrls(PCT90INDEX).fBoxY, 1280, fgBoxStH
    'Date
    gSetCtrl tmNDTCtrls(DATEENTRDINDEX), 7420, tmNDTCtrls(PCT90INDEX).fBoxY, 1035, fgBoxStH
    
    'NSF Checkes
    gSetCtrl tmNDTCtrls(NSFCHKSINDEX), 30, 30 + 350, 1515, fgBoxStH
    'Date Last Billed
    gSetCtrl tmNDTCtrls(DATELSTINVINDEX), 1560, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Date Last Payment
    gSetCtrl tmNDTCtrls(DATELSTPAYMINDEX), 3090, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Avg # days to pay
    gSetCtrl tmNDTCtrls(AVGTOPAYINDEX), 4620, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    '# Days to pay lst payment
    gSetCtrl tmNDTCtrls(LSTTOPAYINDEX), 6150, tmNDTCtrls(NSFCHKSINDEX).fBoxY, 2325, fgBoxStH
    
    'Pct 90
    gSetCtrl tmDTCtrls(PCT90INDEX), 30, tmCtrls(EDICINDEX).fBoxY + 350 + 400, 1515, fgBoxStH
    'Currect A/R
    gSetCtrl tmDTCtrls(CURRARINDEX), 1560, tmDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Unbilled
    gSetCtrl tmDTCtrls(UNBILLEDINDEX), 3090, tmDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Hi Credit
    gSetCtrl tmDTCtrls(HICREDITINDEX), 4625, tmDTCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    
    'Total Gross
    gSetCtrl tmDTCtrls(TOTALGROSSINDEX), 6150, tmDTCtrls(PCT90INDEX).fBoxY, 1280, fgBoxStH
    'Date
    gSetCtrl tmDTCtrls(DATEENTRDINDEX), 7420, tmDTCtrls(PCT90INDEX).fBoxY, 1035, fgBoxStH
    
    'NSF Checkes
    gSetCtrl tmDTCtrls(NSFCHKSINDEX), 30, tmCtrls(EDICINDEX).fBoxY + 350 + 750, 1515, fgBoxStH
    'Date Last Billed
    gSetCtrl tmDTCtrls(DATELSTINVINDEX), 1560, tmDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Date Last Payment
    gSetCtrl tmDTCtrls(DATELSTPAYMINDEX), 3090, tmDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Avg # days to pay
    gSetCtrl tmDTCtrls(AVGTOPAYINDEX), 4620, tmDTCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Avg # days to pay
    gSetCtrl tmDTCtrls(LSTTOPAYINDEX), 6150, tmDTCtrls(NSFCHKSINDEX).fBoxY, 2325, fgBoxStH
    
    llMax = 0
    llShortMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If tmCtrls(ilLoop).fBoxX >= 0 Then
            tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
            Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
            Loop
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmCtrls(ilLoop).fBoxX > 90 Then
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
        If ilLoop <= POLITICALINDEX Then
            If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llShortMax Then
                llShortMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
            End If
        End If

    Next ilLoop
    
    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmNDTCtrls) Step 1
        If tmNDTCtrls(ilLoop).fBoxX >= 0 Then
            tmNDTCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmNDTCtrls(ilLoop).fBoxW)
            Do While (tmNDTCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmNDTCtrls(ilLoop).fBoxW = tmNDTCtrls(ilLoop).fBoxW + 1
            Loop
            tmNDTCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmNDTCtrls(ilLoop).fBoxX)
            Do While (tmNDTCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmNDTCtrls(ilLoop).fBoxX = tmNDTCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmNDTCtrls(ilLoop).fBoxX > 90 Then
                    If tmNDTCtrls(ilLoop - 1).fBoxX + tmNDTCtrls(ilLoop - 1).fBoxW + 15 < tmNDTCtrls(ilLoop).fBoxX Then
                        tmNDTCtrls(ilLoop - 1).fBoxW = tmNDTCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmNDTCtrls(ilLoop - 1).fBoxX + tmNDTCtrls(ilLoop - 1).fBoxW + 15 > tmNDTCtrls(ilLoop).fBoxX Then
                        tmNDTCtrls(ilLoop - 1).fBoxW = tmNDTCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmNDTCtrls(ilLoop).fBoxX + tmNDTCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmNDTCtrls(ilLoop).fBoxX + tmNDTCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmDTCtrls) Step 1
        If tmDTCtrls(ilLoop).fBoxX >= 0 Then
            tmDTCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmDTCtrls(ilLoop).fBoxW)
            Do While (tmDTCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDTCtrls(ilLoop).fBoxW = tmDTCtrls(ilLoop).fBoxW + 1
            Loop
            tmDTCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmDTCtrls(ilLoop).fBoxX)
            Do While (tmDTCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDTCtrls(ilLoop).fBoxX = tmDTCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDTCtrls(ilLoop).fBoxX > 90 Then
                    If tmDTCtrls(ilLoop - 1).fBoxX + tmDTCtrls(ilLoop - 1).fBoxW + 15 < tmDTCtrls(ilLoop).fBoxX Then
                        tmDTCtrls(ilLoop - 1).fBoxW = tmDTCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDTCtrls(ilLoop - 1).fBoxX + tmDTCtrls(ilLoop - 1).fBoxW + 15 > tmDTCtrls(ilLoop).fBoxX Then
                        tmDTCtrls(ilLoop - 1).fBoxW = tmDTCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDTCtrls(ilLoop).fBoxX + tmDTCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDTCtrls(ilLoop).fBoxX + tmDTCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    


    pbcAdvt.Picture = LoadPicture("")
    If imShortForm Then
        pbcAdvt.Width = llShortMax + 15
        plcAdvt.Width = llShortMax + 2 * fgBevelX + 30
    Else
        pbcAdvt.Width = llMax
        plcAdvt.Width = llMax + 2 * fgBevelX + 15
    End If
    pbcDirect.Picture = LoadPicture("")
    pbcDirect.Width = llMax
    plcDirect.Width = llMax + 2 * fgBevelX + 15
    
    pbcNotDirect.Picture = LoadPicture("")
    pbcNotDirect.Width = llMax
    plcNotDirect.Width = llMax + 2 * fgBevelX + 15
    If imShortForm Then
        plcAdvt.Left = plcDirect.Left + plcDirect.Width - plcAdvt.Width
        pbcAdvt.Left = plcAdvt.Left + fgBevelX + 15
    End If
    cbcSelect.Left = plcAdvt.Left + plcAdvt.Width - cbcSelect.Width
    lacCode.Left = plcAdvt.Left + plcAdvt.Width - lacCode.Width
    
    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    
    cmcErase.Left = Advt.Width / 2 - cmcErase.Width / 2
    cmcUpdate.Left = cmcErase.Left - cmcUpdate.Width - ilGap
    cmcCancel.Left = cmcUpdate.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    
    cmcUndo.Left = cmcErase.Left + cmcErase.Width + ilGap
    cmcMerge.Left = cmcUndo.Left + cmcUndo.Width + ilGap
    cmcSplitCue.Left = cmcMerge.Left + cmcMerge.Width + ilGap
    
    frcBill.Left = Advt.Width / 2 - frcBill.Width / 2
    
    gMoveCtrl pbcAdvt, plcDemo, tmCtrls(DEMOINDEX).fBoxX, tmCtrls(DEMOINDEX).fBoxY
    pbcDm.Move plcDemo.Left + 150, plcDemo.Top + 285
   
    pbcAdvt.BackColor = vbWhite
    pbcDirect.BackColor = vbWhite
    
    frcBill.BackColor = Advt.BackColor
    frcBill.Top = plcAdvt.Top + plcAdvt.height + fgBevelX * 2 'L.Bianchi 05/26/2021
    
    rbcBill(0).BackColor = Advt.BackColor
    rbcBill(1).BackColor = Advt.BackColor
    lacBill.BackColor = Advt.BackColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitDmShow                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitDmShow(ilBoxNo As Integer)
'
'   mInitDmShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilIndex As Integer
    Dim slStr As String
    If ilBoxNo < imLBDmCtrls Or (ilBoxNo > UBound(tmDmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DMPRICETYPEINDEX
            If imPriceType = 0 Then
                slStr = "N/A"
            ElseIf imPriceType = 1 Then
                slStr = "CPM"
            ElseIf imPriceType = 2 Then
                slStr = "CPP"
            Else
                slStr = ""
            End If
            gSetShow pbcDm, slStr, tmDmCtrls(ilBoxNo)
        Case DMDEMOINDEX, DMDEMOINDEX + 2, DMDEMOINDEX + 4, DMDEMOINDEX + 6
            ilIndex = (ilBoxNo - DMDEMOINDEX) \ 2
            If lbcDemo(ilIndex).ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcDemo(ilIndex).List(lbcDemo(ilIndex).ListIndex)
            End If
            gSetShow pbcDm, slStr, tmDmCtrls(ilBoxNo)
        Case DMVALUEINDEX, DMVALUEINDEX + 2, DMVALUEINDEX + 4, DMVALUEINDEX + 6
            ilIndex = (ilBoxNo - DMVALUEINDEX) \ 2
            If imPriceType = 2 Then
                gFormatStr smTarget(ilIndex), FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            Else
                gFormatStr smTarget(ilIndex), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            End If
            gSetShow pbcDm, smTarget(ilIndex), tmDmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvSortBranch                  *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to invoice*
'*                      sorting and process            *
'*                      communication back from        *
'*                      invoice sorting                *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mInvSortBranch() As Integer
'
'   ilRet = mInvSortBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcInvSort, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mInvSortBranch = False
        Exit Function
    End If
    If igWinStatus(INVOICESORTLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mInvSortBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mInvSortBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "V"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mInvSortBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcInvSort.Clear
        smInvSortCodeTag = ""
        mInvSortPop
        If imTerminate Then
            mInvSortBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcInvSort
        sgMNmName = ""
        If gLastFound(lbcInvSort) > 0 Then
            imChgMode = True
            lbcInvSort.ListIndex = gLastFound(lbcInvSort)
            edcDropDown.Text = lbcInvSort.List(lbcInvSort.ListIndex)
            imChgMode = False
            mInvSortBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcInvSort.ListIndex = 1
            edcDropDown.Text = lbcInvSort.List(1)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvSortPop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Invoice sort list     *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mInvSortPop()
'
'   mCompPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcInvSort.ListIndex
    If ilIndex > 1 Then
        slName = lbcInvSort.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "V"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(Advt, lbcInvSort, lbcInvSortCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Advt, lbcInvSort, tmInvSortCode(), smInvSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvSortPopErr
        gCPErrorMsg ilRet, "mInvSortPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcInvSort.AddItem "[None]", 0  'Force as first item on list
        lbcInvSort.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcInvSort
            If gLastFound(lbcInvSort) > 1 Then
                lbcInvSort.ListIndex = gLastFound(lbcInvSort)
            Else
                lbcInvSort.ListIndex = -1
            End If
        Else
            lbcInvSort.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mInvSortPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLkBoxBranch                    *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Lock   *
'*                      Box and process                *
'*                      communication back from Lock   *
'*                      Box                            *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mLkBoxBranch() As Integer
'
'   ilRet = mLkBoxBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcLkBox, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Or (edcDropDown.Text = "[None]") Then
        mLkBoxBranch = False
        imDoubleClickName = False
        Exit Function
    End If
    If igWinStatus(LOCKBOXESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mLkBoxBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(LOCKBOXESLIST)) Then
    '    imDoubleClickName = False
    '    mLkBoxBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
'            gSetMenuState False
    'Screen.MousePointer = vbHourGlass  'Wait
    sgArfCallType = "L"
    igArfCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mLkBoxBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
'        gSetMenuState True
        lbcLkBox.Clear
        smLkBoxCodeTag = ""
        mLkBoxPop
        If imTerminate Then
            mLkBoxBranch = False
            Exit Function
        End If
        gFindMatch sgArfName, 1, lbcLkBox
        sgArfName = ""
        If gLastFound(lbcLkBox) > 0 Then
            imChgMode = True
            lbcLkBox.ListIndex = gLastFound(lbcLkBox)
            edcDropDown.Text = lbcLkBox.List(lbcLkBox.ListIndex)
            imChgMode = False
            mLkBoxBranch = False
            mSetChg LKBOXINDEX
        Else
            imChgMode = True
            lbcLkBox.ListIndex = 1
            edcDropDown.Text = lbcLkBox.List(1)
            imChgMode = False
            mSetChg LKBOXINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mLkBoxPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Lock Box list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mLkBoxPop()
'
'   mLkBoxPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcLkBox.ListIndex
    If ilIndex > 1 Then
        slName = lbcLkBox.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "L"
    ilOffSet(0) = gFieldOffset("Arf", "ArfType") '2
    'ilRet = gIMoveListBox(Advt, lbcLkBox, lbcLkBoxCode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Advt, lbcLkBox, tmLkBoxCode(), smLkBoxCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLkBoxPopErr
        gCPErrorMsg ilRet, "mLkBoxPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcLkBox.AddItem "[None]", 0  'Force as first item on list
        lbcLkBox.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcLkBox
            If gLastFound(lbcLkBox) > 1 Then
                lbcLkBox.ListIndex = gLastFound(lbcLkBox)
            Else
                lbcLkBox.ListIndex = -1
            End If
        Else
            lbcLkBox.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mLkBoxPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilLoop As Integer
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Code number
    Dim slStr As String

    slStr = edcCRMID.Text
    tmAdfx.lCrmId = gStrDecToLong(slStr, 0)

    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmAdf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(ABBRINDEX).iChg Then
        tmAdf.sAbbr = edcAbbr.Text
    End If
    If Not ilTestChg Or tmCtrls(POLITICALINDEX).iChg Then
        Select Case imPolitical
            Case 0  'Yes
                tmAdf.sPolitical = "Y"
            Case 1  'No
                tmAdf.sPolitical = "N"
            Case Else
                tmAdf.sPolitical = "N"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmAdf.sState = "A"
            Case 1  'Dormant
                tmAdf.sState = "D"
            Case Else
                tmAdf.sState = "A"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(PRODINDEX).iChg Then
        If smProduct = "[None]" Then
            tmAdf.sProduct = ""
        Else
            tmAdf.sProduct = smProduct
        End If
    End If
    If Not ilTestChg Or tmCtrls(SPERSONINDEX).iChg Then
        If lbcSPerson.ListIndex >= 2 Then
            'If lbcSPerson.ListIndex <= Traffic!lbcSalesperson.ListCount + 1 Then    '+2 - 1
                slNameCode = tgSalesperson(lbcSPerson.ListIndex - 2).sKey  'Traffic!lbcSalesperson.List(lbcSPerson.ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmAdf.iSlfCode = CInt(slCode)
            'Else
            '    slNameCode = Traffic!lbcSPersonCombo.List(lbcSPerson.ListIndex - 2 - Traffic!lbcSalesperson.ListCount)
            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    On Error GoTo mMoveCtrlToRecErr
            '    gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            '    On Error GoTo 0
            '    slCode = Trim$(slCode)
            '    tmAdf.iSlfCode = -CInt(slCode)
            'End If
        Else
            tmAdf.iSlfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(AGENCYINDEX).iChg Then
        If lbcAgency.ListIndex >= 2 Then
            slNameCode = tgAgency(lbcAgency.ListIndex - 2).sKey    'Traffic!lbcAgency.List(lbcAgency.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iAgfCode = CInt(slCode)
        Else
            tmAdf.iAgfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(REPCODEINDEX).iChg Then
        If tgSpf.sARepCodes = "N" Then
            tmAdf.sCodeRep = ""
        Else
            tmAdf.sCodeRep = edcRepCode.Text
        End If
    End If
    If Not ilTestChg Or tmCtrls(AGYCODEINDEX).iChg Then
        If tgSpf.sAAgyCodes = "N" Then
            tmAdf.sCodeAgy = ""
        Else
            tmAdf.sCodeAgy = edcAgyCode.Text
        End If
    End If
    If Not ilTestChg Or tmCtrls(STNCODEINDEX).iChg Then
        If tgSpf.sAStnCodes = "N" Then
            tmAdf.sCodeStn = ""
        Else
            tmAdf.sCodeStn = edcStnCode.Text
        End If
    End If
    For ilLoop = 0 To 1 Step 1
        If Not ilTestChg Or tmCECtrls(CECOMPINDEX + ilLoop).iChg Then
            If lbcComp(ilLoop).ListIndex >= 2 Then
                slNameCode = tgCompCode(lbcComp(ilLoop).ListIndex - 2).sKey    'lbcCompCode.List(lbcComp(ilLoop).ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmAdf.iMnfComp(ilLoop) = CInt(slCode)
            Else
                tmAdf.iMnfComp(ilLoop) = 0
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To 1 Step 1
        If Not ilTestChg Or tmCECtrls(CEEXCLINDEX + ilLoop).iChg Then
            If lbcExcl(ilLoop).ListIndex >= 2 Then
                slNameCode = tgExclCode(lbcExcl(ilLoop).ListIndex - 2).sKey    'lbcExclCode.List(lbcExcl(ilLoop).ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmAdf.iMnfExcl(ilLoop) = CInt(slCode)
            Else
                tmAdf.iMnfExcl(ilLoop) = 0
            End If
        End If
    Next ilLoop
    If Not ilTestChg Or tmCtrls(DEMOINDEX).iChg Then
        Select Case imPriceType 'cbcPriceType.ListIndex
            Case 1  'CPM
                tmAdf.sCppCpm = "M"
                For ilLoop = 0 To 3 Step 1
                    If lbcDemo(ilLoop).ListIndex >= 1 Then
                        slNameCode = tgDemoCode(lbcDemo(ilLoop).ListIndex - 1).sKey    'lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        tmAdf.iMnfDemo(ilLoop) = Val(slCode)
                        'gStrToPDN smTarget(ilLoop), 2, 4, tmAdf.sTarget(ilLoop)
                        tmAdf.lTarget(ilLoop) = gStrDecToLong(smTarget(ilLoop), 2)
                    Else
                        tmAdf.iMnfDemo(ilLoop) = 0
                        'slStr = ""
                        'gStrToPDN slStr, 2, 4, tmAdf.sTarget(ilLoop)
                        tmAdf.lTarget(ilLoop) = 0
                    End If
                Next ilLoop
            Case 2  'CPP
                tmAdf.sCppCpm = "P"
                For ilLoop = 0 To 3 Step 1
                    If lbcDemo(ilLoop).ListIndex >= 1 Then
                        slNameCode = tgDemoCode(lbcDemo(ilLoop).ListIndex - 1).sKey    'lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        tmAdf.iMnfDemo(ilLoop) = Val(slCode)
                        'gStrToPDN smTarget(ilLoop), 2, 4, tmAdf.sTarget(ilLoop)
                        tmAdf.lTarget(ilLoop) = gStrDecToLong(smTarget(ilLoop), 2)
                    Else
                        tmAdf.iMnfDemo(ilLoop) = 0
                        'slStr = ""
                        'gStrToPDN slStr, 2, 4, tmAdf.sTarget(ilLoop)
                        tmAdf.lTarget(ilLoop) = 0
                    End If
                Next ilLoop
            Case Else   'N/A
                tmAdf.sCppCpm = "N"
                For ilLoop = 0 To 3 Step 1
                    tmAdf.iMnfDemo(ilLoop) = 0
                    'slStr = ""
                    'gStrToPDN slStr, 2, 4, tmAdf.sTarget(ilLoop)
                    tmAdf.lTarget(ilLoop) = 0
                Next ilLoop
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CREDITRESTRINDEX).iChg Then
        Select Case lbcCreditRestr.ListIndex
            Case 0  'No restrictions
                tmAdf.sCreditRestr = "N"
            Case 1  'Credit Limit
                tmAdf.sCreditRestr = "L"
            Case 2  'cash in advance weekly
                tmAdf.sCreditRestr = "W"
            Case 3  'Cash in advance monthly
                tmAdf.sCreditRestr = "M"
            Case 4  'Cash in advance quarterly
                tmAdf.sCreditRestr = "T"
            Case 5  'Prohibit
                tmAdf.sCreditRestr = "P"
            Case Else
                If (tgUrf(0).sCredit <> "I") Or (imShortForm) Then
                    tmAdf.sCreditRestr = "N"
                Else
                    tmAdf.sCreditRestr = ""
                End If
        End Select
    End If
    If lbcCreditRestr.ListIndex <> 1 Then
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmAdf.sCreditLimit
        tmAdf.lCreditLimit = 0
    Else
        If Not ilTestChg Or tmCtrls(CREDITRESTRINDEX + 1).iChg Then
            slStr = edcCreditLimit.Text
            'gStrToPDN slStr, 2, 5, tmAdf.sCreditLimit
            tmAdf.lCreditLimit = gStrDecToLong(slStr, 2)
        End If
    End If
    If Not ilTestChg Or tmCtrls(PAYMRATINGINDEX).iChg Then
        Select Case lbcPaymRating.ListIndex
            Case 0  'Quick
                tmAdf.sPaymRating = "0"
            Case 1  'Normal
                tmAdf.sPaymRating = "1"
            Case 2  'Slow
                tmAdf.sPaymRating = "2"
            Case 3  'Difficult
                tmAdf.sPaymRating = "3"
            Case 4  'in Collection
                tmAdf.sPaymRating = "4"
            Case Else
                If (tgUrf(0).sPayRate <> "I") Or (imShortForm) Then
                    tmAdf.sPaymRating = "1"
                Else
                    tmAdf.sPaymRating = ""
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CREDITAPPROVALINDEX).iChg Then
        Select Case lbcCreditApproval.ListIndex
            Case 0  'Requires Checking
                tmAdf.sCrdApp = "R"
            Case 1  'Approved
                tmAdf.sCrdApp = "A"
            Case 2  'Denied
                tmAdf.sCrdApp = "D"
            Case Else
                'If (tgUrf(0).sChgCrRt <> "I") Then
'                    tmAdf.sCrdApp = "R"
                'Else
                '    tmAdf.sCrdApp = "A"
                'End If
                'Restored 8/7/03 since all user only see the short form
                If (tgUrf(0).sChgCrRt <> "I") Then
                    tmAdf.sCrdApp = "R"   'Requires Checking
                Else
                    tmAdf.sCrdApp = "A"   'Approved
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CREDITRATINGINDEX).iChg Then
        tmAdf.sCrdRtg = edcRating.Text
    End If
    If Not ilTestChg Or tmCtrls(RATEONINVINDEX).iChg Then
        Select Case imRateOnInv
            Case 0  'Yes
                tmAdf.sRateOnInv = "Y"
            Case 1  'No
                tmAdf.sRateOnInv = "N"
            Case Else
                tmAdf.sRateOnInv = "Y"
        End Select
    End If

    If Not ilTestChg Or tmCtrls(ISCIINDEX).iChg Then
        Select Case imISCI
            Case 0  'Yes
                tmAdf.sShowISCI = "Y"
            Case 1  'No
                tmAdf.sShowISCI = "N"
            Case 2      'yes and remove W/O Leader
                tmAdf.sShowISCI = "W"
            Case Else
                If tgSpf.sAISCI = "A" Then
                    tmAdf.sShowISCI = "Y"
                ElseIf tgSpf.sAISCI = "X" Then
                    tmAdf.sShowISCI = "N"
                Else
                    tmAdf.sShowISCI = ""
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(REPINVINDEX).iChg Then
        Select Case imRepInv
            Case 0  'Internal
                tmAdf.sRepInvGen = "I"
            Case 1  'External
                tmAdf.sRepInvGen = "E"
            Case Else
                tmAdf.sRepInvGen = "I"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(INVSORTINDEX).iChg Then
        If lbcInvSort.ListIndex >= 2 Then
            slNameCode = tmInvSortCode(lbcInvSort.ListIndex - 2).sKey 'lbcInvSortCode.List(lbcInvSort.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iMnfSort = CInt(slCode)
        Else
            tmAdf.iMnfSort = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(PACKAGEINDEX).iChg Then
        Select Case imPackage
            Case 0  'Daypart
                tmAdf.sPkInvShow = "D"
            Case 1  'Times
                tmAdf.sPkInvShow = "T"
            Case Else
                '6/27/12:  To match change when tabbing into field
                tmAdf.sPkInvShow = "T"  'D
        End Select
    End If
    If Not ilTestChg Or tmCtrls(REPMGINDEX).iChg Then
        Select Case imRepMG
            Case 0  'Yes
                tmAdf.sAllowRepMG = "Y"
            Case 1  'No
                tmAdf.sAllowRepMG = "N"
            Case Else
                tmAdf.sAllowRepMG = "Y"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(BONUSONINVINDEX).iChg Then
        Select Case imBonusOnInv
            Case 0  'Yes
                tmAdf.sBonusOnInv = "Y"
            Case 1  'No
                tmAdf.sBonusOnInv = "N"
            Case Else
                tmAdf.sBonusOnInv = "Y"
        End Select
    End If
    'L.Bianchi 05/31/2021
    If Not ilTestChg Or tmCtrls(REFIDINDEX).iChg Then
        tmAdfx.sRefID = Trim$(edcRefId.Text)
    End If
    'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    If Not ilTestChg Or tmCtrls(DIRECTREFIDINDEX).iChg Then
        tmAdfx.sDirectRefId = Trim$(edcDirectRefID.Text)
    End If
       
    If rbcBill(1).Value Then
        tmAdf.sBillAgyDir = "D"
    Else
        tmAdf.sBillAgyDir = "A"
    End If
    For ilLoop = 0 To 2 Step 1
        If Not ilTestChg Or tmCtrls(CADDRINDEX + ilLoop).iChg Then
            tmAdf.sCntrAddr(ilLoop) = edcCAddr(ilLoop).Text
        End If
    Next ilLoop
    For ilLoop = 0 To 2 Step 1
        If Not ilTestChg Or tmCtrls(BADDRINDEX + ilLoop).iChg Then
            tmAdf.sBillAddr(ilLoop) = edcBAddr(ilLoop).Text
        End If
    Next ilLoop
    If Not ilTestChg Or tmCtrls(ADDRIDINDEX).iChg Then
        tmAdf.sAddrID = edcAddrID.Text
    End If
    If Not ilTestChg Or tmCtrls(BUYERINDEX).iChg Then
        If lbcBuyer.ListIndex >= 2 Then
            slNameCode = tgBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iPnfBuyer = CInt(slCode)
        Else
            tmAdf.iPnfBuyer = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(PAYABLEINDEX).iChg Then
        If lbcPayable.ListIndex >= 2 Then
            slNameCode = tgPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iPnfPay = CInt(slCode)
        Else
            tmAdf.iPnfPay = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(LKBOXINDEX).iChg Then
        If lbcLkBox.ListIndex >= 2 Then
            slNameCode = tmLkBoxCode(lbcLkBox.ListIndex - 2).sKey  'lbcLkBoxCode.List(lbcLkBox.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iArfLkCode = CInt(slCode)
        Else
            tmAdf.iArfLkCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(EDICINDEX).iChg Then
        If lbcEDI(0).ListIndex >= 1 Then
            slNameCode = tmEDICode(lbcEDI(0).ListIndex - 1).sKey   'lbcEDICode.List(lbcEDI(0).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iArfCntrCode = CInt(slCode)
        Else
            tmAdf.iArfCntrCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(EDIIINDEX).iChg Then
        If lbcEDI(1).ListIndex >= 1 Then
            slNameCode = tmEDICode(lbcEDI(1).ListIndex - 1).sKey   'lbcEDICode.List(lbcEDI(1).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iArfInvCode = CInt(slCode)
        Else
            tmAdf.iArfInvCode = 0
        End If
    End If
     
    If Not ilTestChg Or tmCtrls(SUPPRESSNETINDEX).iChg Then ' TTP 10622 - 2023-03-08 JJB
        Select Case imSuppressNet
            Case 0  'No
                tmAdfx.iInvFeatures = 0
            Case 1  'Yes
                tmAdfx.iInvFeatures = 1
        End Select
    End If
    
'    If Not ilTestChg Or tmCtrls(PRTSTYLEINDEX).iChg Then
'        Select Case imPrtStyle
'            Case 0  'Wide
'                tmAdf.sCntrPrtSz = "W"
'            Case 1  'Narrow
'                tmAdf.sCntrPrtSz = "N"
'            Case Else
                If tgSpf.sAPrtStyle = "W" Then
                    tmAdf.sCntrPrtSz = "W"
                ElseIf tgSpf.sAPrtStyle = "N" Then
                    tmAdf.sCntrPrtSz = "N"
                Else
                    tmAdf.sCntrPrtSz = ""
                End If
'        End Select
'    End If
    If Not ilTestChg Or tmCtrls(TERMSINDEX).iChg Then
        If lbcTerms.ListIndex >= 2 Then
            slNameCode = tmTermsCode(lbcTerms.ListIndex - 2).sKey 'lbcInvSortCode.List(lbcInvSort.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Advt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAdf.iMnfInvTerms = CInt(slCode)
        Else
            tmAdf.iMnfInvTerms = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(TAXINDEX).iChg Then
        '12/17/06-Change to tax by agency or vehicle
        'Select Case lbcTax.ListIndex
        '    Case 0  'No; No
        '        tmAdf.sSlsTax(0) = "N"
        '        tmAdf.sSlsTax(1) = "N"
        '    Case 1  'Yes; No
        '        tmAdf.sSlsTax(0) = "Y"
        '        tmAdf.sSlsTax(1) = "N"
        '    Case 2  'No; Yes
        '        tmAdf.sSlsTax(0) = "N"
        '        tmAdf.sSlsTax(1) = "Y"
        '    Case 3  'Yes; Yes
        '        tmAdf.sSlsTax(0) = "Y"
        '        tmAdf.sSlsTax(1) = "Y"
        '    Case Else
        '        If Not imTaxDefined Then
        '            tmAdf.sSlsTax(0) = "N"
        '            tmAdf.sSlsTax(1) = "N"
        '        Else
        '            tmAdf.sSlsTax(0) = ""
        '            tmAdf.sSlsTax(1) = ""
        '        End If
        'End Select
        If lbcTax.ListIndex >= 1 Then
            tmAdf.iTrfCode = lbcTax.ItemData(lbcTax.ListIndex)
        Else
            tmAdf.iTrfCode = 0
        End If
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Function mGetAvfCode() As Integer
    
   '--------------------------------------------------------------------------------
    ' Procedure  :       mGetAvfCode
    ' Description:       Loads the AvfCode
    ' Created by :       J. Butner
    ' Date-Time  :       3/15/2024
    ' Parameters :       NONE
    '--------------------------------------------------------------------------------
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    
    slSQLQuery = "SELECT TOP 1 avfCode FROM avf_AdVendor WHERE avfName = 'Megaphone'"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    
    If Not rst_Temp.EOF Then
        mGetAvfCode = Trim(rst_Temp!avfCode)
    End If
    
End Function

Function mGetVendorAdvertiserID(ilVendorAvfCode As Integer, ilAdfCode As Integer) As String
    
   '--------------------------------------------------------------------------------
    ' Procedure  :       mGetVendorAdvertiserID
    ' Description:       Loads the External Vendor's Advertiser ID for the specified Vendor and Counterpoint Advertiser
    ' Created by :       J. White
    ' Machine    :       CSI-FM
    ' Date-Time  :       3/8/2024-10:07:38
    ' Parameters :       ilVendorAvfCode (Integer)
    '                    ilAdfCode (Integer)
    ' Notes      :       Function borrowed from CNTRSCHD module - JJB 3/15/24
    '--------------------------------------------------------------------------------
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    
    slSQLQuery = ""
    slSQLQuery = slSQLQuery & " Select "
    slSQLQuery = slSQLQuery & "     vacExternalID "
    slSQLQuery = slSQLQuery & " from "
    slSQLQuery = slSQLQuery & "     VAC_Vendor_Adf "
    slSQLQuery = slSQLQuery & " where "
    slSQLQuery = slSQLQuery & "     vacAdfCode = " & ilAdfCode & " and "
    slSQLQuery = slSQLQuery & "     vacAvfcode = " & ilVendorAvfCode
    
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    
    If Not rst_Temp.EOF Then
        mGetVendorAdvertiserID = Trim(rst_Temp!vacExternalID)
    End If
    
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
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
    Dim slRecCode As String
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    Dim slStr As String
    Dim ilComp As Integer
    Dim ilExcl As Integer
    Dim ilDemo As Integer
    bmAxfChg = False
    
    edcCRMID.Text = ""
    If tmAdfx.lCrmId <> 0 Then
        edcCRMID.Text = gLongToStrDec(tmAdfx.lCrmId, 0)
    End If
    
    imAvfCode = mGetAvfCode
    
    edcName.Text = Trim$(tmAdf.sName)
    edcAbbr.Text = Trim$(tmAdf.sAbbr)
    edcRefId.Text = Trim$(tmAdfx.sRefID)
    edcDirectRefID.Text = Trim$(tmAdfx.sDirectRefId)
    edcMegaphoneAdvID.Text = mGetVendorAdvertiserID(imAvfCode, tmAdf.iCode)
    smMegaphoneAdvID_Original = edcMegaphoneAdvID.Text
    
    Select Case tmAdf.sPolitical
        Case "Y"
            imPolitical = 0
        Case "N"
            imPolitical = 1
        Case Else
            imPolitical = 1
    End Select
    If tmAdf.sState = "D" Then
        imState = 1 'Dormant
    Else
        imState = 0 'Active
    End If
    smProduct = Trim$(tmAdf.sProduct)
    If smProduct = "" Then
        smProduct = "[None]"
    End If
    mProdPop tmAdf.iCode
    'look up salesperson name from code number
    lbcSPerson.ListIndex = 1
    smSPerson = ""
    'If tmAdf.iSlfCode >= 0 Then
        slRecCode = Trim$(str$(tmAdf.iSlfCode))
        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2:Salesperson)", Advt
            On Error GoTo 0
            If slRecCode = slCode Then
                lbcSPerson.ListIndex = ilLoop + 2
                smSPerson = lbcSPerson.List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    'Else
    '    slRecCode = Trim$(Str$(-tmAdf.iSlfCode))
    '    For ilLoop = 0 To Traffic!lbcSPersonCombo.ListCount - 1 Step 1
    '        slNameCode = Traffic!lbcSPersonCombo.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '        On Error GoTo mMoveRecToCtrlErr
    '        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2:Combo Salesperson)", Advt
    '        On Error GoTo 0
    '        If slRecCode = slCode Then
    '            lbcSPerson.ListIndex = ilLoop + 2 + Traffic!lbcSalesperson.ListCount
    '            smSPerson = lbcSPerson.List(ilLoop + 2 + Traffic!lbcSalesperson.ListCount)
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If
    'look up agency name from code number
    lbcAgency.ListIndex = 1
    smAgency = ""
    slRecCode = Trim$(str$(tmAdf.iAgfCode))
    For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        slNameCode = tgAgency(ilLoop).sKey 'Traffic!lbcAgency.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcAgency.ListIndex = ilLoop + 2
            smAgency = lbcAgency.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    edcRepCode.Text = Trim$(tmAdf.sCodeRep)
    edcAgyCode.Text = Trim$(tmAdf.sCodeAgy)
    edcStnCode.Text = Trim$(tmAdf.sCodeStn)
    'look up competitive name from code number
    For ilComp = 0 To 1 Step 1
        lbcComp(ilComp).ListIndex = 1
        smComp(ilComp) = ""
        slRecCode = Trim$(str$(tmAdf.iMnfComp(ilComp)))
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey   'lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
            On Error GoTo 0
            If slRecCode = slCode Then
                lbcComp(ilComp).ListIndex = ilLoop + 2
                smComp(ilComp) = lbcComp(ilComp).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    Next ilComp
    For ilExcl = 0 To 1 Step 1
        lbcExcl(ilExcl).ListIndex = 1
        smExcl(ilExcl) = ""
        slRecCode = Trim$(str$(tmAdf.iMnfExcl(ilExcl)))
        For ilLoop = 0 To UBound(tgExclCode) - 1 Step 1 'lbcExclCode.ListCount - 1 Step 1
            slNameCode = tgExclCode(ilLoop).sKey   'lbcExclCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
            On Error GoTo 0
            If slRecCode = slCode Then
                lbcExcl(ilExcl).ListIndex = ilLoop + 2
                smExcl(ilExcl) = lbcExcl(ilExcl).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    Next ilExcl
    Select Case tmAdf.sCppCpm
        Case "N"
            imPriceType = 0 'cbcPriceType.ListIndex = 0
            For ilLoop = 0 To 3 Step 1
                lbcDemo(ilLoop).ListIndex = -1
                smSvTarget(ilLoop) = ""
                smTarget(ilLoop) = ""
            Next ilLoop
        Case "M"
            imPriceType = 1 'cbcPriceType.ListIndex = 1
            For ilLoop = 0 To 3 Step 1
                lbcDemo(ilLoop).ListIndex = -1
                smSvDemo(ilLoop) = ""
                smSvTarget(ilLoop) = ""
                smTarget(ilLoop) = ""
                slRecCode = Trim$(str$(tmAdf.iMnfDemo(ilLoop)))
                For ilDemo = 0 To UBound(tgDemoCode) - 1 Step 1 'lbcDemoCode.ListCount - 1 Step 1
                    slNameCode = tgDemoCode(ilDemo).sKey   'lbcDemoCode.List(ilDemo)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
                    On Error GoTo 0
                    If slRecCode = slCode Then
                        lbcDemo(ilLoop).ListIndex = ilDemo + 1
                        smSvDemo(ilLoop) = lbcDemo(ilLoop).List(ilDemo + 1)
                        slStr = gLongToStrDec(tmAdf.lTarget(ilLoop), 2)
                        smSvTarget(ilLoop) = slStr
                        smTarget(ilLoop) = slStr
                        Exit For
                    End If
                Next ilDemo
            Next ilLoop
        Case "P"
            imPriceType = 2 'cbcPriceType.ListIndex = 2
            For ilLoop = 0 To 3 Step 1
                lbcDemo(ilLoop).ListIndex = -1
                smSvDemo(ilLoop) = ""
                smSvTarget(ilLoop) = ""
                smTarget(ilLoop) = ""
                slRecCode = Trim$(str$(tmAdf.iMnfDemo(ilLoop)))
                For ilDemo = 0 To UBound(tgDemoCode) - 1 Step 1 'lbcDemoCode.ListCount - 1 Step 1
                    slNameCode = tgDemoCode(ilDemo).sKey   'lbcDemoCode.List(ilDemo)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
                    On Error GoTo 0
                    If slRecCode = slCode Then
                        lbcDemo(ilLoop).ListIndex = ilDemo + 1
                        smSvDemo(ilLoop) = lbcDemo(ilLoop).List(ilDemo + 1)
                        slStr = gLongToStrDec(tmAdf.lTarget(ilLoop), 2)
                        'Remove decimal part of number
                        slStr = Left$(slStr, Len(slStr) - 3)
                        smSvTarget(ilLoop) = slStr
                        smTarget(ilLoop) = slStr
                        Exit For
                    End If
                Next ilDemo
            Next ilLoop
        Case Else
            imPriceType = -1 'cbcPriceType.ListIndex = 0
            For ilLoop = 0 To 3 Step 1
                lbcDemo(ilLoop).ListIndex = -1
                smSvDemo(ilLoop) = ""
                smSvTarget(ilLoop) = ""
                smTarget(ilLoop) = ""
            Next ilLoop
    End Select
    smBkoutPool = Trim$(tmAdf.sBkoutPoolStatus)
    If smBkoutPool = "" Then
        smBkoutPool = "N"
    End If
    Select Case tmAdf.sCreditRestr
        Case "N"
            lbcCreditRestr.ListIndex = 0
        Case "L"
            lbcCreditRestr.ListIndex = 1
        Case "W"
            lbcCreditRestr.ListIndex = 2
        Case "M"
            lbcCreditRestr.ListIndex = 3
        Case "T"
            lbcCreditRestr.ListIndex = 4
        Case "P"
            lbcCreditRestr.ListIndex = 5
        Case Else
            lbcCreditRestr.ListIndex = -1
    End Select
    If tmAdf.sCreditRestr = "L" Then
        'gPDNToStr tmAdf.sCreditLimit, 2, slStr
        slStr = gLongToStrDec(tmAdf.lCreditLimit, 2)
        edcCreditLimit.Text = slStr
    Else
        edcCreditLimit.Text = ""
    End If
    Select Case tmAdf.sPaymRating
        Case "0"
            lbcPaymRating.ListIndex = 0
        Case "1"
            lbcPaymRating.ListIndex = 1
        Case "2"
            lbcPaymRating.ListIndex = 2
        Case "3"
            lbcPaymRating.ListIndex = 3
        Case "4"
            lbcPaymRating.ListIndex = 4
        Case Else
            lbcPaymRating.ListIndex = -1
    End Select
    Select Case tmAdf.sCrdApp
        Case "R"
            lbcCreditApproval.ListIndex = 0
        Case "A"
            lbcCreditApproval.ListIndex = 1
        Case "D"
            lbcCreditApproval.ListIndex = 2
        Case Else
            lbcCreditApproval.ListIndex = -1
    End Select
    edcRating.Text = Trim$(tmAdf.sCrdRtg)
    Select Case tmAdf.sRepInvGen
        Case "I"
            imRepInv = 0
        Case "E"
            imRepInv = 1
        Case Else
            imRepInv = 0    '-1
    End Select
    Select Case tmAdf.sRateOnInv
        Case "Y"
            imRateOnInv = 0
        Case "N"
            imRateOnInv = 1
        Case Else
            imRateOnInv = 0 '-1
    End Select
    Select Case tmAdf.sShowISCI
        Case "Y"
            imISCI = 0
        Case "N"
            imISCI = 1
        Case "W"
            imISCI = 2
        Case Else
            'imISCI = -1
            If tgSpf.sAISCI = "A" Then
                imISCI = 0
            ElseIf tgSpf.sAISCI = "X" Then
                imISCI = 1
            Else
                imISCI = -1
            End If
    End Select
    Select Case tmAdf.sAllowRepMG
        Case "Y"
            imRepMG = 0
        Case "N"
            imRepMG = 1
        Case Else
            imRepMG = 0 '-1
    End Select
    Select Case tmAdf.sBonusOnInv
        Case "Y"
            imBonusOnInv = 0
        Case "N"
            imBonusOnInv = 1
        Case Else
            imBonusOnInv = 0    '-1
    End Select
    lbcInvSort.ListIndex = 1
    smInvSort = ""
    slRecCode = Trim$(str$(tmAdf.iMnfSort))
    For ilLoop = 0 To UBound(tmInvSortCode) - 1 Step 1 'lbcInvSortCode.ListCount - 1 Step 1
        slNameCode = tmInvSortCode(ilLoop).sKey   'lbcInvSortCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcInvSort.ListIndex = ilLoop + 2
            smInvSort = lbcInvSort.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    Select Case tmAdf.sPkInvShow
        Case "D"
            imPackage = 0
        Case "T"
            imPackage = 1
        Case Else
            imPackage = -1
    End Select
    If tmAdf.sBillAgyDir = "D" Then
        rbcBill(1).Value = True
    Else
        rbcBill(0).Value = True
    End If
    For ilLoop = 0 To 2 Step 1
        edcCAddr(ilLoop).Text = Trim$(tmAdf.sCntrAddr(ilLoop))
    Next ilLoop
    For ilLoop = 0 To 2 Step 1
        edcBAddr(ilLoop).Text = Trim$(tmAdf.sBillAddr(ilLoop))
    Next ilLoop
    edcAddrID.Text = Trim$(tmAdf.sAddrID)
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    mBuyerPop tmAdf.iCode, "", -1
    lbcBuyer.ListIndex = 1
    smBuyer = ""
    smOrigBuyer = ""
    slRecCode = Trim$(str$(tmAdf.iPnfBuyer))
    For ilLoop = 0 To UBound(tgBuyerCode) - 1 Step 1 'lbcBuyerCode.ListCount - 1 Step 1
        slNameCode = tgBuyerCode(ilLoop).sKey  'lbcBuyerCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcBuyer.ListIndex = ilLoop + 2
            smBuyer = lbcBuyer.List(ilLoop + 2)
            smOrigBuyer = smBuyer
            Exit For
        End If
    Next ilLoop
    mPayablePop tmAdf.iCode, "", -1
    lbcPayable.ListIndex = 1
    smPayable = ""
    smOrigPayable = ""
    slRecCode = Trim$(str$(tmAdf.iPnfPay))
    For ilLoop = 0 To UBound(tgPayableCode) - 1 Step 1  'lbcPayableCode.ListCount - 1 Step 1
        slNameCode = tgPayableCode(ilLoop).sKey    'lbcPayableCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcPayable.ListIndex = ilLoop + 2
            smPayable = lbcPayable.List(ilLoop + 2)
            smOrigPayable = smPayable
            Exit For
        End If
    Next ilLoop
    'look up LkBox from code number
    lbcLkBox.ListIndex = 1
    smLkBox = ""
    slRecCode = Trim$(str$(tmAdf.iArfLkCode))
    For ilLoop = 0 To UBound(tmLkBoxCode) - 1 Step 1 'lbcLkBoxCode.ListCount - 1 Step 1
        slNameCode = tmLkBoxCode(ilLoop).sKey  'lbcLkBoxCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcLkBox.ListIndex = ilLoop + 2
            smLkBox = lbcLkBox.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    lbcEDI(0).ListIndex = 0
    smEDIC = ""
    slRecCode = Trim$(str$(tmAdf.iArfCntrCode))
    For ilLoop = 0 To UBound(tmEDICode) - 1 Step 1  'lbcEDICode.ListCount - 1 Step 1
        slNameCode = tmEDICode(ilLoop).sKey    'lbcEDICode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcEDI(0).ListIndex = ilLoop + 1
            smEDIC = lbcEDI(0).List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    lbcEDI(1).ListIndex = 0
    smEDII = ""
    slRecCode = Trim$(str$(tmAdf.iArfInvCode))
    For ilLoop = 0 To UBound(tmEDICode) - 1 Step 1  'lbcEDICode.ListCount - 1 Step 1
        slNameCode = tmEDICode(ilLoop).sKey 'lbcEDICode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcEDI(1).ListIndex = ilLoop + 1
            smEDII = lbcEDI(1).List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    ReDim tmPdf(0 To 0) As PDF
    If Trim$(smEDII) = "PDF EMail" Then
        mPopPDFEMail
    End If
'    Select Case tmAdf.sCntrPrtSz
'        Case "W"
'            imPrtStyle = 0
'        Case "N"
'            imPrtStyle = 1
'        Case Else
'            imPrtStyle = -1
'    End Select
    lbcTerms.ListIndex = 1
    smTerms = lbcTerms.List(1)
    slRecCode = Trim$(str$(tmAdf.iMnfInvTerms))
    For ilLoop = 0 To UBound(tmTermsCode) - 1 Step 1 'lbcInvSortCode.ListCount - 1 Step 1
        slNameCode = tmTermsCode(ilLoop).sKey   'lbcInvSortCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcTerms.ListIndex = ilLoop + 2
            smTerms = lbcTerms.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    '12/17/06-Change to tax by agency or vehicle
    'If (tmAdf.sSlsTax(0) = "N") And (tmAdf.sSlsTax(1) = "N") Then
    '    lbcTax.ListIndex = 0
    'ElseIf (tmAdf.sSlsTax(0) = "Y") And (tmAdf.sSlsTax(1) = "N") Then
    '    lbcTax.ListIndex = 1
    'ElseIf (tmAdf.sSlsTax(0) = "N") And (tmAdf.sSlsTax(1) = "Y") Then
    '    lbcTax.ListIndex = 2
    'ElseIf (tmAdf.sSlsTax(0) = "Y") And (tmAdf.sSlsTax(1) = "Y") Then
    '    lbcTax.ListIndex = 3
    'Else
    '    lbcTax.ListIndex = -1
    'End If
    'look up Tax from code number
    lbcTax.ListIndex = 0
    smTax = lbcTax.List(lbcTax.ListIndex)
    slRecCode = Trim$(str$(tmAdf.iTrfCode))
    For ilLoop = 0 To UBound(tmTaxSortCode) - 1 Step 1
        slNameCode = tmTaxSortCode(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Advt
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcTax.ListIndex = ilLoop + 1
            smTax = lbcTax.List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    
    If tmAdfx.iInvFeatures = 1 Then ' TTP 10622 - 2023-03-08 JJB
        imSuppressNet = 1 'Suppress
        edcSuppressNet.Text = "Yes"
    Else
        imSuppressNet = 0 'No Suppression
        edcSuppressNet.Text = "No"
    End If
     
    
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    For ilLoop = imLBCECtrls To UBound(tmCECtrls) Step 1
        tmCECtrls(ilLoop).iChg = False
    Next ilLoop
    For ilLoop = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        tmDmCtrls(ilLoop).iChg = False
    Next ilLoop
    'gPDNToStr tmAdf.sPct90, 0, slStr
    slStr = gIntToStrDec(tmAdf.iPct90, 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 0, smPct90
    gPDNToStr tmAdf.sCurrAR, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smCurrAR
    gPDNToStr tmAdf.sUnbilled, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smUnbilled
    gPDNToStr tmAdf.sHiCredit, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smHiCredit
    gPDNToStr tmAdf.sTotalGross, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smTotalGross
    gUnpackDate tmAdf.iDateEntrd(0), tmAdf.iDateEntrd(1), smDateEntrd
    smNSFChks = Trim$(str$(tmAdf.iNSFChks))
    gUnpackDate tmAdf.iDateLstInv(0), tmAdf.iDateLstInv(1), smDateLstInv
    gUnpackDate tmAdf.iDateLstPaym(0), tmAdf.iDateLstPaym(1), smDateLstPaym
    smAvgToPay = Trim$(str$(tmAdf.iAvgToPay))
    smLstToPay = Trim$(str$(tmAdf.iLstToPay))
    smNoInvPd = Trim$(str$(tmAdf.iNoInvPd))
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilLoop As Integer
    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        'gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        'If gLastFound(cbcSelect) <> -1 Then   'Name found
        '    If gLastFound(cbcSelect) <> imSelectedIndex Then
        '        slStr = Trim$(edcName.Text)
        '        If slStr = cbcSelect.List(gLastFound(cbcSelect)) Then
        '            Beep
        '            MsgBox "Advertiser already defined, enter a different name", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
        '            edcName.Text = Trim$(tmAdf.sName) 'Reset text
        '            mSetShow imBoxNo
        '            mSetChg imBoxNo
        '            imBoxNo = NAMEINDEX
        '            mEnableBox imBoxNo
        '            mOKName = False
        '            Exit Function
        '        End If
        '    End If
        'End If
        For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            If tgCommAdf(ilLoop).iCode <> tmAdf.iCode Then
                If StrComp(Trim$(tgCommAdf(ilLoop).sName), slStr, vbTextCompare) = 0 Then
                    If rbcBill(1).Value Then
                        If tgCommAdf(ilLoop).sBillAgyDir = "D" Then
                            If StrComp(Trim$(tgCommAdf(ilLoop).sAddrID), Trim$(edcAddrID.Text), vbTextCompare) = 0 Then
                                Beep
                                MsgBox "Advertiser already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                                edcName.Text = Trim$(tmAdf.sName) 'Reset text
                                mSetShow imBoxNo
                                mSetChg imBoxNo
                                imBoxNo = NAMEINDEX
                                mEnableBox imBoxNo
                                mOKName = False
                                Exit Function
                            End If
                        End If
                    Else
                        If tgCommAdf(ilLoop).sBillAgyDir <> "D" Then
                            Beep
                            MsgBox "Advertiser already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                            edcName.Text = Trim$(tmAdf.sName) 'Reset text
                            mSetShow imBoxNo
                            mSetChg imBoxNo
                            imBoxNo = NAMEINDEX
                            mEnableBox imBoxNo
                            mOKName = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next ilLoop
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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone Advt, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igAdvtCallSource = Val(slStr)
    If igStdAloneMode Then
        igAdvtCallSource = CALLNONE
    End If
    If igAdvtCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgAdvtName = slStr
        Else
            sgAdvtName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPayablePop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Payable Personnel     *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mPayablePop(ilAdvtCode As Integer, slRetainName As String, ilReturnCode As Integer)
'
'   mPayablePop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilPnf As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    ilIndex = lbcPayable.ListIndex
    If ilIndex > 0 Then
        slName = lbcPayable.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'If imSelectedIndex > 0 Then 'Change mode
    '    'If imSelectedIndex = 0 Then
    '    '    ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "P", False, lbcPayable, lbcPayableCode)
    '    'Else
    '        'ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "P", True, 2, lbcPayable, lbcPayableCode)
            ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "P", True, 2, lbcPayable, tgPayableCode(), sgPayableCodeTag)
    '    'End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPayablePopErr
            gCPErrorMsg ilRet, "mPayablePop (gPopPersonnelBox)", Advt
            On Error GoTo 0
            'Filter out any contact not associated with this agency
            If imSelectedIndex = 0 Then
                For ilLoop = UBound(tgPayableCode) - 1 To 0 Step -1 'lbcPayableCode.ListCount - 1 To 0 Step -1
                    ilFound = False
                    slNameCode = tgPayableCode(ilLoop).sKey    'lbcPayableCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If (StrComp(Trim$(slRetainName), Trim$(Left$(slName, 30)), 1) = 0) And (slRetainName <> "") Or (Val(slCode) = ilReturnCode) And (slRetainName <> "") Then
                        ilFound = True
                    Else
                        ilCode = Val(slCode)
                        'For ilPnf = 1 To UBound(imNewPnfCode) - 1 Step 1
                        For ilPnf = 0 To UBound(imNewPnfCode) - 1 Step 1
                            If imNewPnfCode(ilPnf) = ilCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilPnf
                    End If
                    If Not ilFound Then
                        lbcPayable.RemoveItem ilLoop
                        'lbcPayableCode.RemoveItem ilLoop
                        gRemoveItemFromSortCode ilLoop, tgPayableCode()
                    End If
                Next ilLoop
            End If
            lbcPayable.AddItem "[None]", 0  'Force as first item on list
            lbcPayable.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 0 Then
                gFindMatch slName, 1, lbcPayable
                If gLastFound(lbcPayable) > 0 Then
                    lbcPayable.ListIndex = gLastFound(lbcPayable)
                Else
                    lbcPayable.ListIndex = -1
                End If
            Else
                lbcPayable.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    'Else
    '    If lbcPayable.ListCount = 0 Then
    '        lbcPayable.AddItem "[None]", 0  'Force as first item on list
    '        lbcPayable.AddItem "[New]", 0  'Force as first item on list
    '    End If
    'End If
    Exit Sub
mPayablePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPersonnelBranch                *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      personnel and process          *
'*                      communication back from        *
'*                      personnel                      *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mPersonnelBranch() As Integer
'
'   ilRet = mPersonnelBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim slBuyerOrPayable As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName30 As String * 30
    Dim ilBoxNo As Integer
    Dim slNewFlag As String
    Dim slReturnCode As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slName As String

    If imBoxNo = BUYERINDEX Then
        slBuyerOrPayable = "B"
        ilRet = gOptionalLookAhead(edcDropDown, lbcBuyer, imBSMode, slStr)
    Else
        slBuyerOrPayable = "P"
        ilRet = gOptionalLookAhead(edcDropDown, lbcPayable, imBSMode, slStr)
    End If
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mPersonnelBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mPersonnelBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
'        ilRet = gOptionalLookAhead(edcDropDown, lbcPersonnel, imBSMode, slStr)
    igPersonnelCallSource = CALLSOURCEADVERTISER
    'Screen.MousePointer = vbHourGlass  'Wait
    'sgPersonnelName = cbcSelect.List(imSelectedIndex)
    If edcDropDown.Text = "[New]" Then
        'sgPersonnelName = sgPersonnelName & "\" & " "
        sgPersonnelName = " "
    Else
        'sgPersonnelName = sgPersonnelName & "\" & Trim$(edcDropDown.Text)
        sgPersonnelName = Trim$(Left$(edcDropDown.Text, 30))    'Remove phone numbers
        If imBoxNo = BUYERINDEX Then
            If lbcBuyer.ListIndex >= 2 Then
                slNameCode = tgBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmPnfSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmPnf, tmBPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        sgPersonnelName = Trim$(tmBPnf.sName)
                    End If
                End If
            End If
        Else
            If lbcPayable.ListIndex >= 2 Then
                slNameCode = tgPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmPnfSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmPnf, tmPPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        sgPersonnelName = Trim$(tmPPnf.sName)
                    End If
                End If
            End If
        End If
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    If imSelectedIndex = 0 Then
        tmAdf.iCode = 0
    End If
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & Trim$(str$(igPersonnelCallSource)) & "\Advt" & "\" & Trim$(str$(tmAdf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & Trim$(str$(igPersonnelCallSource)) & "\Advt" & "\" & Trim$(str$(tmAdf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igPersonnelCallSource)) & "\Advt" & "\" & Trim$(Str$(tmAdf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igPersonnelCallSource)) & "\Advt" & "\" & Trim$(Str$(tmAdf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
    '    End If
    'End If
    ilBoxNo = imBoxNo
    'lgShellRet = Shell(sgExePath & "Persnnel.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Persnnel.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgPersonnelName)
    igPersonnelCallSource = Val(sgPersonnelName)
    ilParse = gParseItem(slStr, 2, "\", sgPersonnelName)
    ilParse = gParseItem(slStr, 3, "\", slNewFlag)
    ilParse = gParseItem(slStr, 4, "\", slReturnCode)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imBoxNo = ilBoxNo
    imDoubleClickName = False
    mPersonnelBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igPersonnelCallSource = CALLDONE Then  'Done
        igPersonnelCallSource = CALLNONE
'        gSetMenuState True
        If imBoxNo = BUYERINDEX Then
            lbcBuyer.Clear
            sgBuyerCodeTag = ""
            mBuyerPop tmAdf.iCode, sgPersonnelName, Val(slReturnCode)
            If imTerminate Then
                mPersonnelBranch = False
                Exit Function
            End If
            slName30 = sgPersonnelName  'Don't test phone number
            'gFindPartialMatch slName30, 2, 30, lbcBuyer
            'sgPersonnelName = ""
            'If gLastFound(lbcBuyer) > 1 Then
            '    imChgMode = True
            '    lbcBuyer.ListIndex = gLastFound(lbcBuyer)
            '    edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
            '    imChgMode = False
            '    'slNameCode = tgBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            '    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    'imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            ''    'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '    ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
            '    If slNewFlag = "Y" Then
            '        slNameCode = tgBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '        'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '        ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
            '    End If
            '    mPersonnelBranch = False
            '    mSetChg BUYERINDEX
            'Else
            '    imChgMode = True
            '    lbcBuyer.ListIndex = 1
            '    edcDropDown.Text = lbcBuyer.List(1)
            '    imChgMode = False
            '    mSetChg BUYERINDEX
            '    If edcDropDown.Visible Then
            '        edcDropDown.SetFocus
            '    Else
            '        pbcClickFocus.SetFocus
            '    End If
            '    Exit Function
            'End If
            ilFound = False
            For ilLoop = LBound(tgBuyerCode) To UBound(tgBuyerCode) - 1 Step 1
                slNameCode = tgBuyerCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = Val(slReturnCode) Then
                    ilFound = True
                    imChgMode = True
                    lbcBuyer.ListIndex = ilLoop + 2
                    edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
                    imChgMode = False
                    If slNewFlag = "Y" Then
                        slNameCode = tgBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
                        'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
                        ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
                    End If
                    mPersonnelBranch = False
                    mSetChg BUYERINDEX
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imChgMode = True
                lbcBuyer.ListIndex = 1
                edcDropDown.Text = lbcBuyer.List(1)
                imChgMode = False
                mSetChg BUYERINDEX
                If edcDropDown.Visible Then
                    edcDropDown.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
                Exit Function
            End If
        Else
            lbcPayable.Clear
            sgPayableCodeTag = ""
            mPayablePop tmAdf.iCode, sgPersonnelName, Val(slReturnCode)
            If imTerminate Then
                mPersonnelBranch = False
                Exit Function
            End If
            slName30 = sgPersonnelName  'Don't test phone number
            'gFindPartialMatch slName30, 2, 30, lbcPayable
            'sgPersonnelName = ""
            'If gLastFound(lbcPayable) > 1 Then
            '    imChgMode = True
            '    lbcPayable.ListIndex = gLastFound(lbcPayable)
            '    edcDropDown.Text = lbcPayable.List(lbcPayable.ListIndex)
            '    imChgMode = False
            '    'slNameCode = tgPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            '    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    'imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '    ''ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '    'ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
            '    If slNewFlag = "Y" Then
            '        slNameCode = tgPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '        'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '        ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
            '    End If
            '    mPersonnelBranch = False
            '    mSetChg PAYABLEINDEX
            'Else
            '    imChgMode = True
            '    lbcPayable.ListIndex = 1
            '    edcDropDown.Text = lbcPayable.List(1)
            '    imChgMode = False
            '    mSetChg PAYABLEINDEX
            '    If edcDropDown.Visible Then
            '        edcDropDown.SetFocus
            '    Else
            '        pbcClickFocus.SetFocus
            '    End If
            '    Exit Function
            'End If
            ilFound = False
            For ilLoop = LBound(tgPayableCode) To UBound(tgPayableCode) - 1 Step 1
                slNameCode = tgPayableCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = Val(slReturnCode) Then
                    ilFound = True
                    imChgMode = True
                    lbcPayable.ListIndex = ilLoop + 2
                    edcDropDown.Text = lbcPayable.List(lbcPayable.ListIndex)
                    imChgMode = False
                    If slNewFlag = "Y" Then
                        slNameCode = tgPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
                        'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
                        ReDim Preserve imNewPnfCode(0 To UBound(imNewPnfCode) + 1) As Integer
                    End If
                    mPersonnelBranch = False
                    mSetChg PAYABLEINDEX
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imChgMode = True
                lbcPayable.ListIndex = 1
                edcDropDown.Text = lbcPayable.List(1)
                imChgMode = False
                mSetChg PAYABLEINDEX
                If edcDropDown.Visible Then
                    edcDropDown.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
                Exit Function
            End If
        End If
    End If
    If igPersonnelCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igPersonnelCallSource = CALLNONE
        sgPersonnelName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igPersonnelCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igPersonnelCallSource = CALLNONE
        sgPersonnelName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
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

    imPopReqd = False
    'ilRet = gPopAdvtBox(Advt, cbcSelect, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(Advt, cbcSelect, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopAdvtBox)", Advt
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    mPopLbcName
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProdBranch                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      advertiser product and process *
'*                      communication back from        *
'*                      advertiser product             *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mProdBranch() As Integer
'
'   ilRet = mProdBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    If imSelectedIndex = 0 Then 'New selected
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mProdBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
'        ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
    If (Not imDoubleClickName) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igAdvtProdCallSource = CALLSOURCEADVERTISER
    sgAdvtProdName = cbcSelect.List(imSelectedIndex)
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    Else
    '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AdvtProd.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AdvtProd.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'Advt.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProdBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igAdvtProdCallSource = CALLDONE Then  'Done
        igAdvtProdCallSource = CALLNONE
'        gSetMenuState True
        lbcProd.Clear
        sgProdCodeTag = ""
        mProdPop tmAdf.iCode
        If imTerminate Then
            mProdBranch = False
            Exit Function
        End If
        gFindMatch sgAdvtProdName, 1, lbcProd
        sgAdvtProdName = ""
        If gLastFound(lbcProd) > 0 Then
            imChgMode = True
            lbcProd.ListIndex = gLastFound(lbcProd)
            edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
            imChgMode = False
            mProdBranch = False
            mSetChg PRODINDEX
        Else
            imChgMode = True
            lbcProd.ListIndex = -1
            edcDropDown.Text = sgAdvtProdName
            imChgMode = False
            mSetChg PRODINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igAdvtProdCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtProdCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mProdPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser product    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mProdPop(ilAdvtCode As Integer)
'
'   mProdPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcProd.ListIndex
    If ilIndex > 0 Then
        slName = lbcProd.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    If imSelectedIndex > 0 Then 'Change mode
        'ilRet = gPopAdvtProdBox(Advt, ilAdvtCode, lbcProd, lbcProdCode)
        ilRet = gPopAdvtProdBox(Advt, ilAdvtCode, lbcProd, tgProdCode(), sgProdCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mProdPopErr
            gCPErrorMsg ilRet, "mProdPop (gIMoveListBox)", Advt
            On Error GoTo 0
            lbcProd.AddItem "[None]", 0  'Force as first item on list
'            lbcProd.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 0 Then
                gFindMatch slName, 1, lbcProd
                If gLastFound(lbcProd) > 0 Then
                    lbcProd.ListIndex = gLastFound(lbcProd)
                Else
                    lbcProd.ListIndex = -1
                End If
            Else
                lbcProd.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    Else
        If lbcProd.ListCount = 0 Then
            lbcProd.AddItem "[None]", 0  'Force as first item on list
        End If
    End If
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer, ilShowMissingMsg As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim rs As ADODB.Recordset
    slNameCode = tgAdvertiser(ilSelectIndex - 1).sKey  'Traffic!lbcAdvertiser.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", Advt
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmAdfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Advt
    On Error GoTo 0
    If tmAdf.iPnfBuyer > 0 Then
        tmPnfSrchKey.iCode = tmAdf.iPnfBuyer
        ilRet = btrGetEqual(hmPnf, tmBPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        'On Error GoTo mReadRecErr
        'gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Advt
        'On Error GoTo 0
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            If ilShowMissingMsg Then
                MsgBox "Buyer Name Missing", vbOKOnly + vbExclamation
            End If
            tmBPnf.iCode = 0
            tmAdf.iPnfBuyer = 0
        Else
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual- Buyer)", Advt
            On Error GoTo 0
        End If
    Else
        tmBPnf.iCode = 0
    End If
    If tmAdf.iPnfPay > 0 Then
        tmPnfSrchKey.iCode = tmAdf.iPnfPay
        ilRet = btrGetEqual(hmPnf, tmPPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        'On Error GoTo mReadRecErr
        'gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Advt
        'On Error GoTo 0
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            If ilShowMissingMsg Then
                MsgBox "Payables Contact Name Missing", vbOKOnly + vbExclamation
            End If
            tmPPnf.iCode = 0
            tmAdf.iPnfPay = 0
        Else
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual- Payable)", Advt
            On Error GoTo 0
        End If
    Else
        tmPPnf.iCode = 0
    End If
    
    mReadAxf tmAdf.iCode
    
    'L.Bianchi '05/31/2021' start
    tmAdfx.iCode = tmAdf.iCode
    SQLQuery = "select * from ADFX_Advertisers WHERE adfxCode =" & tmAdf.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        If Not IsNull(rs!adfxRefId) Then
            tmAdfx.sRefID = rs!adfxRefId
            tmAdfx.sDirectRefId = rs!adfxDirectRefId
        End If
        tmAdfx.lCrmId = rs!adfxCrmId
        
        If Not IsNull(rs!adfxInvFeatures) Then
            tmAdfx.iInvFeatures = rs!adfxInvFeatures
        Else
            tmAdfx.iInvFeatures = 0
        End If
    Else
        tmAdfx.sRefID = ""
        tmAdfx.sDirectRefId = ""
        tmAdfx.lCrmId = 0
        tmAdfx.iInvFeatures = -1
    End If
    'L.Bianchi '05/31/2021' End
    
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
Private Function mSaveRec_VAC() As Integer
   'JJB 03/15/2024
   
    Dim rs As ADODB.Recordset
    Dim llCount As Long
    
    On Error GoTo mSaveRecErr
    
    SQLQuery = "Select vacExternalID from VAC_Vendor_Adf where vacAvfCode = " & imAvfCode & " and vacAdfCode = " & tmAdf.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    tmAdfx.iCode = tmAdf.iCode
    If Not rs.EOF Then
        SQLQuery = "UPDATE VAC_Vendor_Adf SET "
        SQLQuery = SQLQuery & "     vacExternalID =  '" & edcMegaphoneAdvID.Text & "' "
        SQLQuery = SQLQuery & " Where "
        SQLQuery = SQLQuery & "     vacAvfCode    = " & imAvfCode & " and "
        SQLQuery = SQLQuery & "     vacAdfCode    = " & tmAdf.iCode
        
         If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", "VAC-mSaveRec"
        End If
    Else
        SQLQuery = "INSERT INTO VAC_Vendor_Adf"
        SQLQuery = SQLQuery & "( "
        SQLQuery = SQLQuery & " vacAvfCode, "
        SQLQuery = SQLQuery & " vacAdfCode, "
        SQLQuery = SQLQuery & " vacExternalID "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & " VALUES "
        SQLQuery = SQLQuery & "("
        SQLQuery = SQLQuery & imAvfCode & ","
        SQLQuery = SQLQuery & tmAdf.iCode & ","
        SQLQuery = SQLQuery & "'" & edcMegaphoneAdvID.Text & "'"
        SQLQuery = SQLQuery & ") "
            
        If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
               gHandleError "TrafficErrors.txt", "VAC-mSaveRec"
        End If
    End If
    smMegaphoneAdvID_Original = edcMegaphoneAdvID.Text
    
    mSaveRec_VAC = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec_VAC = False
    
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slNameCode As String  'Name and Code number
    Dim slCode As String    'Code number
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim slStr As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlPnf As PNF
    Dim slEDII As String
    Dim rs As ADODB.Recordset
    Dim llCount As Long
    Dim bSave_Vac As Boolean
    mSetShow imBoxNo
    If rbcBill(1).Value Then
        tmCtrls(CADDRINDEX).iReq = True
'        If tgSpf.sAPrtStyle = "A" Then
'            tmCtrls(PRTSTYLEINDEX).iReq = True
'        End If
        If imTaxDefined Then
            tmCtrls(TAXINDEX).iReq = True
        End If
    End If
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    'If mTestFields(REFIDINDEX, SHOWMSG) = NO Then
        'mSaveRec = False
        'Exit Function
    'End If
    
    tmAdfx.iInvFeatures = IIF(edcSuppressNet.Text = "Yes", 1, 0) ' TTP 10622 - 2023-03-08 JJB
    
    If Val(edcCRMID.Text) > 2147483646 Then
        MsgBox "CRM ID cannot be larger than 2147483646", vbOKOnly + vbExclamation
        mSaveRec = False
        Exit Function
    End If
    If Len(edcCRMID.Text) < 1 Then
        tmAdfx.lCrmId = 0
    Else
        tmAdfx.lCrmId = CLng(edcCRMID.Text)
    End If
    
    '4/3/15: check if address fields should be cleared
    mClearAddressIfRequired
    
    tmCtrls(CADDRINDEX).iReq = False
'    tmCtrls(PRTSTYLEINDEX).iReq = False
    tmCtrls(TAXINDEX).iReq = False
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "ADF.Btr")
        If imSelectedIndex <> 0 Then
            sgAVIndicatorID = smAVIndicatorID
            sgXDSCue = smXDSCue
            If Not mReadRec(imSelectedIndex, SETFORWRITE, False) Then
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            smAVIndicatorID = sgAVIndicatorID
            smXDSCue = sgXDSCue
        End If
        If (imSelectedIndex = 0) And (imShortForm) Then 'New selected
            mMoveCtrlToRec False
        Else
            mMoveCtrlToRec True
        End If
        tmAdf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
        If imSelectedIndex = 0 Then 'New selected
            tmAdf.iCode = 0  'Autoincrement
            'slStr = ""
            'gStrToPDN slStr, 0, 2, tmAdf.sPct90
            tmAdf.iPct90 = 0
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAdf.sCurrAR
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAdf.sUnbilled
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAdf.sHiCredit
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAdf.sTotalGross
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, tmAdf.iDateEntrd(0), tmAdf.iDateEntrd(1)
            tmAdf.iNSFChks = 0
            slStr = ""
            gPackDate slStr, tmAdf.iDateLstInv(0), tmAdf.iDateLstInv(1)
            slStr = ""
            gPackDate slStr, tmAdf.iDateLstPaym(0), tmAdf.iDateLstPaym(1)
            tmAdf.iAvgToPay = 0
            tmAdf.iLstToPay = 0
            tmAdf.iNoInvPd = 0
            tmAdf.sNewBus = "Y" 'New advertiser without contract reference
            tmAdf.lGuar = 0
            tmAdf.iEndDate(0) = tgChfCntr.iEndDate(0)
            tmAdf.iEndDate(1) = tgChfCntr.iEndDate(1)
            tmAdf.iMnfBus = 0
            tmAdf.iMerge = 0
            tmAdf.sBkoutPoolStatus = "N"
            tmAdf.sUnused2 = ""
            tmAdf.iLastMonthNew = 0
            tmAdf.iLastYearNew = 0
            ilRet = btrInsert(hmAdf, tmAdf, imAdfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Advertiser)"
        Else 'Old record-Update
            tmAdf.sBkoutPoolStatus = smBkoutPool
            ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
            slMsg = "mSaveRec (btrUpdate:Advertiser)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    
    ' tmAdfx.lCrmId = CLng(edcCRMID.Text) ' JD 09-27-22
    
     'L.Bianchi 05/28/2021 start, JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
    SQLQuery = "SELECT * FROM ADFX_Advertisers WHERE adfxCode =" & tmAdf.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    tmAdfx.iCode = tmAdf.iCode
    If Not rs.EOF Then
        SQLQuery = "UPDATE ADFX_Advertisers SET "
        SQLQuery = SQLQuery & "     adfxRefId       = '" & tmAdfx.sRefID & "', "
        SQLQuery = SQLQuery & "     adfxDirectRefId = '" & tmAdfx.sDirectRefId & "', "
        SQLQuery = SQLQuery & "     adfxCrmId       =  " & tmAdfx.lCrmId & ", "
        SQLQuery = SQLQuery & "     adfxInvFeatures =  " & tmAdfx.iInvFeatures
        SQLQuery = SQLQuery & " Where "
        SQLQuery = SQLQuery & "     adfxCode = " & tmAdf.iCode
        
         If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", "Advt-mSaveRec"
        End If
    Else
        SQLQuery = "INSERT INTO ADFX_Advertisers"
        SQLQuery = SQLQuery & "( "
        SQLQuery = SQLQuery & " adfxCode, "
        SQLQuery = SQLQuery & " adfxRefId, "
        SQLQuery = SQLQuery & " adfxDirectRefID, "
        SQLQuery = SQLQuery & " adfxCrmId, "
        SQLQuery = SQLQuery & " adfxInvFeatures "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & " VALUES "
        SQLQuery = SQLQuery & "("
        SQLQuery = SQLQuery & tmAdf.iCode & ","
        SQLQuery = SQLQuery & "'" & tmAdfx.sRefID & "',"
        SQLQuery = SQLQuery & "'" & tmAdfx.sDirectRefId & "',"
        SQLQuery = SQLQuery & tmAdfx.lCrmId & ","
        SQLQuery = SQLQuery & tmAdfx.iInvFeatures
        SQLQuery = SQLQuery & ") "
            
        If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
               gHandleError "TrafficErrors.txt", "Advt-mSaveRec"
        End If
    End If
    'L.Bianchi 05/28/2021 End
    
    bSave_Vac = mSaveRec_VAC()
    
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Advt
    On Error GoTo 0
    If (imSelectedIndex = 0) Then   'and (tgSpf.sRemoteUsers = "Y") Then  'New selected
        'Add to tgCommAdf
        tgCommAdf(UBound(tgCommAdf)).iCode = tmAdf.iCode
        tgCommAdf(UBound(tgCommAdf)).sName = tmAdf.sName
        tgCommAdf(UBound(tgCommAdf)).sAbbr = tmAdf.sAbbr
        tgCommAdf(UBound(tgCommAdf)).iMnfSort = tmAdf.iMnfSort
        tgCommAdf(UBound(tgCommAdf)).sBillAgyDir = tmAdf.sBillAgyDir
        tgCommAdf(UBound(tgCommAdf)).sState = tmAdf.sState
        tgCommAdf(UBound(tgCommAdf)).sAllowRepMG = tmAdf.sAllowRepMG
        tgCommAdf(UBound(tgCommAdf)).sBonusOnInv = tmAdf.sBonusOnInv
        tgCommAdf(UBound(tgCommAdf)).sRepInvGen = tmAdf.sRepInvGen
        tgCommAdf(UBound(tgCommAdf)).sFirstCntrAddr = tmAdf.sCntrAddr(0)
        tgCommAdf(UBound(tgCommAdf)).sPolitical = tmAdf.sPolitical  '5-26-06
        tgCommAdf(UBound(tgCommAdf)).sAddrID = tmAdf.sAddrID
        tgCommAdf(UBound(tgCommAdf)).iTrfCode = tmAdf.iTrfCode
        'ReDim Preserve tgCommAdf(1 To UBound(tgCommAdf) + 1) As ADFEXT
        ReDim Preserve tgCommAdf(0 To UBound(tgCommAdf) + 1) As ADFEXT
        'Sort by code so that binary search can be used
        'If UBound(tgCommAdf) - 1 > 1 Then
        If UBound(tgCommAdf) - 1 > 0 Then
            'ArraySortTyp fnAV(tgCommAdf(), 1), UBound(tgCommAdf) - 1, 0, LenB(tgCommAdf(1)), 0, -1, 0
            ArraySortTyp fnAV(tgCommAdf(), 0), UBound(tgCommAdf), 0, LenB(tgCommAdf(0)), 0, -1, 0
        End If
        If tmAdf.sRepInvGen <> "E" Then
            igInternalAdfCount = igInternalAdfCount + 1
        End If
    Else
        ilRet = gBinarySearchAdf(tmAdf.iCode)
        If ilRet <> -1 Then
            tgCommAdf(ilRet).iCode = tmAdf.iCode
            tgCommAdf(ilRet).sName = tmAdf.sName
            tgCommAdf(ilRet).sAbbr = tmAdf.sAbbr
            tgCommAdf(ilRet).iMnfSort = tmAdf.iMnfSort
            tgCommAdf(ilRet).sBillAgyDir = tmAdf.sBillAgyDir
            tgCommAdf(ilRet).sState = tmAdf.sState
            tgCommAdf(ilRet).sAllowRepMG = tmAdf.sAllowRepMG
            tgCommAdf(ilRet).sBonusOnInv = tmAdf.sBonusOnInv
            tgCommAdf(ilRet).sPolitical = tmAdf.sPolitical      '5-26-06
            tgCommAdf(ilRet).sAddrID = tmAdf.sAddrID
            tgCommAdf(ilRet).iTrfCode = tmAdf.iTrfCode
            '4/3/15: Set address to handle the case where it was cleared (Direct changed to Agency and no transaction exist)
            tgCommAdf(ilRet).sFirstCntrAddr = tmAdf.sCntrAddr(0)
            If (tgCommAdf(ilRet).sRepInvGen = "E") And (tmAdf.sRepInvGen <> "E") Then
                igInternalAdfCount = igInternalAdfCount + 1
            End If
            tgCommAdf(ilRet).sRepInvGen = tmAdf.sRepInvGen
        End If
    End If
    If Trim$(tmAdf.sProduct) <> "" Then
        gFindMatch tmAdf.sProduct, 0, lbcProd
        If gLastFound(lbcProd) < 0 Then
            imPrfRecLen = Len(tmPrf)
            Do  'Loop until record updated or added
                tmPrf.lCode = 0
                tmPrf.iAdfCode = tmAdf.iCode
                tmPrf.sName = tmAdf.sProduct
                tmPrf.iMnfComp(0) = tmAdf.iMnfComp(0)
                tmPrf.iMnfComp(1) = tmAdf.iMnfComp(1)
                tmPrf.iMnfExcl(0) = tmAdf.iMnfExcl(0)
                tmPrf.iMnfExcl(1) = tmAdf.iMnfExcl(1)
                tmPrf.iPnfBuyer = tmAdf.iPnfBuyer
                tmPrf.sCppCpm = tmAdf.sCppCpm
                For ilLoop = 0 To 3
                    tmPrf.iMnfDemo(ilLoop) = tmAdf.iMnfDemo(ilLoop)
                    tmPrf.lTarget(ilLoop) = tmAdf.lTarget(ilLoop)
                    tmPrf.lLastCPP(ilLoop) = 0
                    tmPrf.lLastCPM(ilLoop) = 0
                Next ilLoop
                tmPrf.sState = "A"
                tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = 0
                ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert:Product)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Advt
            On Error GoTo 0
            'If tgSpf.sRemoteUsers = "Y" Then
                Do
                    'tmPrfSrchKey0.lCode = tmPrf.lCode
                    'ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    'slMsg = "mSaveRec (btrGetEqual:Product)"
                    'On Error GoTo mSaveRecErr
                    'gBtrvErrorMsg ilRet, slMsg, Advt
                    'On Error GoTo 0
                    tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmPrf.lAutoCode = tmPrf.lCode
                    tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate slSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                    gPackTime slSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                    ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                    slMsg = "mSaveRec (btrUpdate:Product)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Advt
                On Error GoTo 0
            'End If
        Else
            slNameCode = tgProdCode(gLastFound(lbcProd) - 1).sKey    'lbcProdCode.List(gLastFound(lbcProd) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            slCode = Trim$(slCode)
            Do
                tmPrfSrchKey0.lCode = Val(slCode)
                ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                tmPrf.iMnfComp(0) = tmAdf.iMnfComp(0)
                tmPrf.iMnfComp(1) = tmAdf.iMnfComp(1)
                tmPrf.iMnfExcl(0) = tmAdf.iMnfExcl(0)
                tmPrf.iMnfExcl(1) = tmAdf.iMnfExcl(1)
                tmPrf.iPnfBuyer = tmAdf.iPnfBuyer
                tmPrf.sCppCpm = tmAdf.sCppCpm
                For ilLoop = 0 To 3
                    tmPrf.iMnfDemo(ilLoop) = tmAdf.iMnfDemo(ilLoop)
                    tmPrf.lTarget(ilLoop) = tmAdf.lTarget(ilLoop)
                Next ilLoop
                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                gPackTime slSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                slMsg = "mSaveRec (btrUpdate:Product)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Advt
            On Error GoTo 0
        End If
    End If
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            Do
                tlPnf.iAdfCode = tmAdf.iCode
                tlPnf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tlPnf.iSyncDate(0), tlPnf.iSyncDate(1)
                gPackTime slSyncTime, tlPnf.iSyncTime(0), tlPnf.iSyncTime(1)
                ilRet = btrUpdate(hmPnf, tlPnf, imPnfRecLen)
                slMsg = "mSaveRec (btrUpdate:Personnel)"
                If ilRet = BTRV_ERR_CONFLICT Then
                    tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
                    ilCRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Advt
            On Error GoTo 0
        End If
    Next ilLoop
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    slEDII = ""
    If lbcEDI(1).ListIndex >= 1 Then
        slEDII = Trim(lbcEDI(1).List(lbcEDI(1).ListIndex))
    End If
    If slEDII = "PDF EMail" Then
        If bmPDFEMailChgd Then
            ilRet = mRemovePDFEMail()
            ilRet = mAddPDFEMail()
        End If
    Else
        ilRet = mRemovePDFEMail()
    End If
    bmPDFEMailChgd = False
    
    ilRet = mAddOrUpdateAxf()
    
    'If sgAdvertiserTag <> "" Then
    '    If slStamp = sgAdvertiserTag Then
    '        sgAdvertiserTag = gFileDateTime(sgDBPath & "ADF.Btr")
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    'Traffic!lbcAdvertiser.RemoveItem imSelectedIndex - 1
    '    gRemoveItemFromSortCode imSelectedIndex - 1, tgAdvertiser()
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = Trim$(tmAdf.sName)
    'If Trim$(tmAdf.sBillAgyDir) = "D" Then
    '    slName = slName & "/" & "Direct"
    'End If
    'cbcSelect.AddItem slName
    'Do While Len(slName) < Len(tmAdf.sName) + 8
    '    slName = slName & " "
    'Loop
    'slName = slName + "\" + LTrim$(Str$(tmAdf.iCode))
    ''Traffic!lbcAdvertiser.AddItem slName
    'tgAdvertiser(UBound(tgAdvertiser)).sKey = slName
    'ReDim Preserve tgAdvertiser(0 To UBound(tgAdvertiser) + 1) As SORTCODE
    'If UBound(tgAdvertiser) - 1 > 0 Then
    '    ArraySortTyp fnAV(tgAdvertiser(),0), UBound(tgAdvertiser), 0, LenB(tgAdvertiser(0)), 0, Len(tgAdvertiser(0).sKey), 0
    'End If
    'cbcSelect.AddItem "[New]", 0
    'mProdPop tmAdf.iCode
    'mBuyerPop tmAdf.iCode, ""
    'mPayablePop tmAdf.iCode, ""
    imChgSaveFlag = True
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function
mSaveRecErr:
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
'*             Created:6/04/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If bmAxfChg Then
        ilAltered = YES
    End If
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcAdvt_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxNoCtrls) Then
        'mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmAdf.sName, edcName, tmCtrls(ilBoxNo)
        Case ABBRINDEX 'Abbreviation
            gSetChgFlag tmAdf.sAbbr, edcAbbr, tmCtrls(ilBoxNo)
        Case STATEINDEX
        Case PRODINDEX  'Product
            gSetChgFlagStr tmAdf.sProduct, smProduct, tmCtrls(ilBoxNo)
        Case SPERSONINDEX   'Salesperson
            gSetChgFlag smSPerson, lbcSPerson, tmCtrls(ilBoxNo)
        Case AGENCYINDEX   'Agency
            gSetChgFlag smAgency, lbcAgency, tmCtrls(ilBoxNo)
        Case REPCODEINDEX 'Rep Code
            gSetChgFlag tmAdf.sCodeRep, edcRepCode, tmCtrls(ilBoxNo)
        Case AGYCODEINDEX 'Agency Code
            gSetChgFlag tmAdf.sCodeAgy, edcAgyCode, tmCtrls(ilBoxNo)
        Case STNCODEINDEX 'Station Code
            gSetChgFlag tmAdf.sCodeStn, edcStnCode, tmCtrls(ilBoxNo)
'        Case COMPINDEX 'Competitive
'            gSetChgFlag smComp(0), lbcComp(0), tmCtrls(COMPINDEX)
'        Case COMPINDEX + 1'Competitive
'            gSetChgFlag smComp(1), lbcComp(1), tmCtrls(COMPINDEX + 1)
'        Case DEMOINDEX
'            Select Case tmAdf.sPriceType
'                Case "N"
'                    slStr = "N/A" 'cbcPriceType.List(0)
'                Case "M"
'                    slStr = "CPM" 'cbcPriceType.List(1)
'                Case "P"
'                    slStr = "CPP" 'cbcPriceType.List(2)
'            End Select
'            gSetChgFlag slStr, cbcPriceType, tmCtrls(ilBoxNo)
'            If Not tmCtrls(ilBoxNo).iChg Then
'                For ilLoop = 0 To 3 Step 1
'                    gSetChgFlag tmAdf.sDemo(ilLoop), lbcDemo(ilLoop), tmCtrls(ilBoxNo)
'                Next ilLoop
'            End If
'            If Not tmCtrls(ilBoxNo).iChg Then
'                For ilLoop = 0 To 3 Step 1
'                    gPDNToStr tmAdf.sTarget(ilLoop), 2, slStr
''                    gSetChgFlag slStr, edcTarget(ilLoop), tmCtrls(ilBoxNo)
'                Next ilLoop
'            End If
        Case CREDITRESTRINDEX
            Select Case tmAdf.sCreditRestr
                Case "N"
                    slStr = lbcCreditRestr.List(0)
                Case "L"
                    slStr = lbcCreditRestr.List(1)
                Case "W"
                    slStr = lbcCreditRestr.List(2)
                Case "M"
                    slStr = lbcCreditRestr.List(3)
                Case "T"
                    slStr = lbcCreditRestr.List(4)
                Case "P"
                    slStr = lbcCreditRestr.List(5)
            End Select
            gSetChgFlag slStr, lbcCreditRestr, tmCtrls(ilBoxNo)
        Case CREDITRESTRINDEX + 1
            'gPDNToStr tmAdf.sCreditLimit, 2, slStr
            slStr = gLongToStrDec(tmAdf.lCreditLimit, 2)
            gSetChgFlag slStr, edcCreditLimit, tmCtrls(ilBoxNo)
        Case PAYMRATINGINDEX
            Select Case tmAdf.sPaymRating
                Case "0"
                    slStr = lbcPaymRating.List(0)
                Case "1"
                    slStr = lbcPaymRating.List(1)
                Case "2"
                    slStr = lbcPaymRating.List(2)
                Case "3"
                    slStr = lbcPaymRating.List(3)
                Case "4"
                    slStr = lbcPaymRating.List(4)
            End Select
            gSetChgFlag slStr, lbcPaymRating, tmCtrls(ilBoxNo)
        Case CREDITAPPROVALINDEX
            Select Case tmAdf.sCrdApp
                Case "R"
                    slStr = lbcCreditApproval.List(0)
                Case "A"
                    slStr = lbcCreditApproval.List(1)
                Case "D"
                    slStr = lbcCreditApproval.List(2)
                Case Else
                    slStr = ""
            End Select
            gSetChgFlag slStr, lbcCreditApproval, tmCtrls(ilBoxNo)
        Case CREDITRATINGINDEX
            gSetChgFlag tmAdf.sCrdRtg, edcRating, tmCtrls(ilBoxNo)
        Case ISCIINDEX
        Case INVSORTINDEX   'Invoice sorting
            gSetChgFlag smInvSort, lbcInvSort, tmCtrls(ilBoxNo)
        Case PACKAGEINDEX
        Case REFIDINDEX 'Ref Id L.Bianchi 05/26/2021
            gSetChgFlag tmAdfx.sRefID, edcRefId, tmCtrls(ilBoxNo)
        Case DIRECTREFIDINDEX 'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
            gSetChgFlag tmAdfx.sDirectRefId, edcDirectRefID, tmCtrls(ilBoxNo)
        Case MEGAPHONEADVID 'JJB - 3/15/24
            gSetChgFlag smMegaphoneAdvID_Original, edcMegaphoneAdvID, tmCtrls(ilBoxNo) 'TTP 11060 JJB
        Case CRMIDINDEX ' JD 09-22-22
            If Len(edcCRMID.Text) > 0 And Not IsNumeric(edcCRMID.Text) Then
                edcCRMID.Text = ""  ' Don't allow anything but numbers
            End If
            If tmCtrls(CRMIDINDEX).iChg <> True Then
                slStr = ""
                If tmAdfx.lCrmId <> 0 Then
                    slStr = CStr(tmAdfx.lCrmId)
                End If
                gSetChgFlag slStr, edcCRMID, tmCtrls(ilBoxNo)
                If Len(edcCRMID.Text) > 0 And IsNumeric(edcCRMID.Text) Then
                    tmAdfx.lCrmId = CLng(edcCRMID.Text)
                Else
                    edcCRMID.Text = ""
                End If
            End If
            
        Case CADDRINDEX 'Contract address
            gSetChgFlag tmAdf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Contract address
            gSetChgFlag tmAdf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Contract address
            gSetChgFlag tmAdf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Billing address
            gSetChgFlag tmAdf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Billing address
            gSetChgFlag tmAdf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Billing address
            gSetChgFlag tmAdf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case ADDRIDINDEX 'Address ID
            gSetChgFlag tmAdf.sAddrID, edcAddrID, tmCtrls(ilBoxNo)
        Case BUYERINDEX 'Buyer Name
            gSetChgFlag smOrigBuyer, lbcBuyer, tmCtrls(ilBoxNo)
        Case PAYABLEINDEX 'Buyer Name
            gSetChgFlag smOrigPayable, lbcPayable, tmCtrls(ilBoxNo)
        Case LKBOXINDEX   'Lock box
            gSetChgFlag smLkBox, lbcLkBox, tmCtrls(ilBoxNo)
        Case EDICINDEX   'EDI Service for contract
            gSetChgFlag smEDIC, lbcEDI(0), tmCtrls(ilBoxNo)
        Case EDIIINDEX   'EDI Service for Invoice
            gSetChgFlag smEDII, lbcEDI(1), tmCtrls(ilBoxNo)
            If bmPDFEMailChgd Then
                tmCtrls(ilBoxNo).iChg = True
            End If
'        Case PRTSTYLEINDEX
        Case TERMSINDEX   'Terms
            gSetChgFlag smTerms, lbcTerms, tmCtrls(ilBoxNo)
        Case TAXINDEX
            '12/17/06-Change to tax by agency or vehicle
            'If (tmAdf.sSlsTax(0) = "N") And (tmAdf.sSlsTax(1) = "N") Then
            '    slStr = lbcTax.List(0)
            'ElseIf (tmAdf.sSlsTax(0) = "Y") And (tmAdf.sSlsTax(1) = "N") Then
            '    slStr = lbcTax.List(1)
            'ElseIf (tmAdf.sSlsTax(0) = "N") And (tmAdf.sSlsTax(1) = "Y") Then
            '    slStr = lbcTax.List(2)
            'ElseIf (tmAdf.sSlsTax(0) = "Y") And (tmAdf.sSlsTax(1) = "Y") Then
            '    slStr = lbcTax.List(3)
            'Else
            '    slStr = ""
            'End If
            gSetChgFlag smTax, lbcTax, tmCtrls(ilBoxNo)
        Case SUPPRESSNETINDEX
            'gSetChgFlag tmAdfx.iInvFeatures, edcSuppressNet, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If bmAxfChg Then
        ilAltered = YES
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And imUpdateAllowed Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    'If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
    If ((imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO)) And (tgSpf.sRemoteUsers <> "Y") And (imUpdateAllowed) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
    'Merge set only if change mode
    'If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") And (tgUrf(0).iRemoteID = 0) Then
    If (Not ilAltered) And (tgUrf(0).sMerge = "I") And (imUpdateAllowed) Then
        cmcMerge.Enabled = True
    Else
        cmcMerge.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set Focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxNoCtrls Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case ABBRINDEX 'Abbreviation
            edcAbbr.SetFocus
        Case STATEINDEX   'Active/Dormant
            pbcState.SetFocus
        Case PRODINDEX 'Product
            edcDropDown.SetFocus
        Case SPERSONINDEX   'Salesperson
            edcDropDown.SetFocus
        Case AGENCYINDEX   'Agency
            edcDropDown.SetFocus
        Case REPCODEINDEX 'Rep Agency Code
            edcRepCode.SetFocus
        Case AGYCODEINDEX 'Agency Code
            edcAgyCode.SetFocus
        Case STNCODEINDEX 'Station Agency Code
            edcStnCode.SetFocus
        Case COMPINDEX   'Competitive/Exclusion
        Case DEMOINDEX 'Demo Code
        Case CREDITAPPROVALINDEX 'Credit Approval
            edcDropDown.SetFocus
        Case CREDITRESTRINDEX 'Credit restrictions
            edcDropDown.SetFocus
        Case CREDITRESTRINDEX + 1 'Limit
            edcCreditLimit.SetFocus
        Case PAYMRATINGINDEX 'Payment rating
            edcDropDown.SetFocus
        Case CREDITRATINGINDEX 'Credit Rating
            edcRating.SetFocus
        Case RATEONINVINDEX   'ISCI on Invoices
            pbcRateOnInv.SetFocus
        Case ISCIINDEX   'ISCI on Invoices
            pbcISCI.SetFocus
        Case REPINVINDEX   'Rep Inv
            pbcRepInv.SetFocus
        Case INVSORTINDEX   'Invoice sorting
            edcDropDown.SetFocus
        Case PACKAGEINDEX   'Package Invoice Show
            pbcPackage.SetFocus
             
        Case MEGAPHONEADVID
            edcMegaphoneAdvID.SetFocus
        Case CRMIDINDEX
            edcCRMID.SetFocus
            
        Case REPMGINDEX   'Rep MG
            pbcRepMG.SetFocus
        Case BONUSONINVINDEX   'Bonus on Invoices
            pbcBonusOnInv.SetFocus
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case ADDRIDINDEX 'Address ID
            edcAddrID.SetFocus
        Case BUYERINDEX 'Buyer
            edcDropDown.SetFocus
        Case PAYABLEINDEX 'Product
            edcDropDown.SetFocus
        Case LKBOXINDEX 'Lock box
            edcDropDown.SetFocus
        Case EDICINDEX   'EDI service for Contracts
            edcDropDown.SetFocus
        Case EDIIINDEX   'EDI service for Contracts
            edcDropDown.SetFocus
'        Case PRTSTYLEINDEX   'Print style
'            pbcPrtStyle.SetFocus
        Case TERMSINDEX   'Invoice sorting
            edcDropDown.SetFocus
        Case TAXINDEX   'Tax
            edcDropDown.SetFocus
         Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
            'edcSuppressNet.SetFocus
            pbcSuppressNet.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    ReDim slTarg(0 To 3) As String
    Dim ilPos As Integer
    Dim slFirst As String
    Dim slLast As String
    Dim ilFirst As Integer
    Dim flWidth As Single
    Dim flSvWidth As Single
    Dim ilCount As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxNoCtrls) Then
        Exit Sub
    End If

    '2/4/16: Add filter to handle the case where the name has illegal characters and it was pasted into the field
    If (ilBoxNo = NAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcName.Text)
        edcName.Text = slStr
    End If
    If (ilBoxNo = PRODINDEX) Then
        '9760
        'slStr = gReplaceIllegalCharacters(edcDropDown.Text)
        slStr = gRemoveIllegalPastedChar(edcDropDown.Text)
        edcDropDown.Text = slStr
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            lbcName.Visible = False
            cmcDropDown.Visible = False
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case ABBRINDEX 'Name
            edcAbbr.Visible = False  'Set visibility
            slStr = edcAbbr.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case POLITICALINDEX   'Active/Dormant
            pbcPolitical.Visible = False  'Set visibility
            If imPolitical = 0 Then
                slStr = "Yes"
            ElseIf imPolitical = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case STATEINDEX   'Active/Dormant
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case PRODINDEX
            lbcProd.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
'            smProduct = edcDropDown.Text
            gSetShow pbcAdvt, smProduct, tmCtrls(ilBoxNo)
        Case SPERSONINDEX   'Salesperson
            lbcSPerson.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcSPerson.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcSPerson.List(lbcSPerson.ListIndex)
            End If
            If Not igSlfFirstNameFirst Then
                ilPos = InStr(slStr, ",")
                If ilPos > 0 Then
                    slLast = Left$(slStr, ilPos - 1)
                    slFirst = right$(slStr, Len(slStr) - ilPos - 1)
                    slStr = slFirst & " " & slLast
                End If
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case AGENCYINDEX   'Agency
            lbcAgency.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcAgency.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcAgency.List(lbcAgency.ListIndex)
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case REPCODEINDEX 'Rep Advertiser Code
            edcRepCode.Visible = False  'Set visibility
            If tgSpf.sARepCodes = "N" Then
                slStr = ""
            Else
                slStr = edcRepCode.Text
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case AGYCODEINDEX 'Agency Code
            edcAgyCode.Visible = False  'Set visibility
            If tgSpf.sAAgyCodes = "N" Then
                slStr = ""
            Else
                slStr = edcAgyCode.Text
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case STNCODEINDEX 'Station Advertiser Code
            edcStnCode.Visible = False  'Set visibility
            If tgSpf.sAStnCodes = "N" Then
                slStr = ""
            Else
                slStr = edcStnCode.Text
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case COMPINDEX   'Competitive
            mCESetShow imCEBoxNo
            imCEBoxNo = -1
            pbcCE.Visible = False
            plcCE.Visible = False
            flSvWidth = tmCtrls(ilBoxNo).fBoxW
            ilCount = 0
            For ilLoop = 0 To 1 Step 1
                If lbcComp(ilLoop).ListIndex > 1 Then
                    ilCount = ilCount + 1
                End If
                If lbcExcl(ilLoop).ListIndex > 1 Then
                    ilCount = ilCount + 1
                End If
            Next ilLoop
            If ilCount > 0 Then
                flWidth = tmCtrls(ilBoxNo).fBoxW / ilCount
            Else
                flWidth = tmCtrls(ilBoxNo).fBoxW
            End If
            ilCount = 0
            slStr = ""
            For ilLoop = 0 To 1 Step 1
                If lbcComp(ilLoop).ListIndex > 1 Then
                    If ilCount > 0 Then
                        slStr = slStr & "," ' "/"
                    End If
                    ilCount = ilCount + 1
                    slStr = slStr & lbcComp(ilLoop).List(lbcComp(ilLoop).ListIndex)
                    tmCtrls(ilBoxNo).fBoxW = ilCount * flWidth
                    gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
                End If
            Next ilLoop
            For ilLoop = 0 To 1 Step 1
                If lbcExcl(ilLoop).ListIndex > 1 Then
                    If ilCount > 0 Then
                        slStr = slStr & "," ' "/"
                    End If
                    ilCount = ilCount + 1
                    slStr = slStr & lbcExcl(ilLoop).List(lbcExcl(ilLoop).ListIndex)
                    tmCtrls(ilBoxNo).fBoxW = ilCount * flWidth
                    gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
                End If
            Next ilLoop
            tmCtrls(ilBoxNo).fBoxW = flSvWidth
        Case DEMOINDEX
            mDmSetShow imDMBoxNo
            imDMBoxNo = -1
            pbcDm.Visible = False
            plcDemo.Visible = False
            ilFirst = True
            If tgSpf.iATargets = 0 Then
                slStr = ""
            Else
                If imPriceType = 0 Then
                    slStr = "N/A"
                ElseIf imPriceType = 1 Then 'CPM
                    slStr = "CPM: "
                ElseIf imPriceType = 2 Then 'CPP
                    slStr = "CPP: "
                Else
                    slStr = ""
                End If
                If (imPriceType = 1) Or (imPriceType = 2) Then
                    For ilLoop = 0 To 3 Step 1
                        gFormatStr smTarget(ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slTarg(ilLoop)
                        If (lbcDemo(ilLoop).ListIndex > 0) Then
                            If ilFirst Then
                                slStr = slStr & lbcDemo(ilLoop).Text & " @ " & slTarg(ilLoop)
                                ilFirst = False
                            Else
                                slStr = slStr & ";  " & lbcDemo(ilLoop).Text & " @ " & slTarg(ilLoop)
                            End If
                        End If
                    Next ilLoop
                End If
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
            'Required since area is not covered
            gInvalidateArea pbcAdvt, Int(tmCtrls(ilBoxNo).fBoxX), Int(tmCtrls(ilBoxNo).fBoxY), Int(tmCtrls(ilBoxNo).fBoxW), Int(fgBoxStH)
            pbcAdvt_Paint
        Case CREDITAPPROVALINDEX   'Credit Approval
            lbcCreditApproval.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcCreditApproval.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCreditApproval.List(lbcCreditApproval.ListIndex)
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case CREDITRESTRINDEX   'Credit Restriction
            lbcCreditRestr.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcCreditRestr.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCreditRestr.List(lbcCreditRestr.ListIndex)
            End If
            If lbcCreditRestr.ListIndex <> 1 Then
                flWidth = tmCtrls(ilBoxNo).fBoxW
                tmCtrls(ilBoxNo).fBoxW = 2 * tmCtrls(ilBoxNo).fBoxW
                gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
                tmCtrls(ilBoxNo).fBoxW = flWidth
            Else
                gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
            End If
            If lbcCreditRestr.ListIndex <> 1 Then
                edcCreditLimit.Text = ""
                slStr = ""
                gSetShow pbcAdvt, slStr, tmCtrls(CREDITRESTRINDEX + 1)
                gPaintArea pbcAdvt, tmCtrls(CREDITRESTRINDEX + 1).fBoxX, tmCtrls(CREDITRESTRINDEX + 1).fBoxY + 120, tmCtrls(CREDITRESTRINDEX + 1).fBoxW - 15, tmCtrls(CREDITRESTRINDEX + 1).fBoxH - 135, pbcAdvt.BackColor     'WHITE
            End If
        Case CREDITRESTRINDEX + 1 'Credit Restriction
            edcCreditLimit.Visible = False
            If lbcCreditRestr.ListIndex = 1 Then
                slStr = edcCreditLimit.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case PAYMRATINGINDEX   'Payment Rating
            lbcPaymRating.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcPaymRating.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcPaymRating.List(lbcPaymRating.ListIndex)
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case CREDITRATINGINDEX 'Name
            edcRating.Visible = False  'Set visibility
            slStr = edcRating.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case RATEONINVINDEX   'Rates on Invoice
            pbcRateOnInv.Visible = False  'Set visibility
            If imRateOnInv = 0 Then
                slStr = "Yes"
            ElseIf imRateOnInv = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case ISCIINDEX   'ISCI on Invoices
            pbcISCI.Visible = False  'Set visibility
            If imISCI = 0 Then
                slStr = "Yes"
            ElseIf imISCI = 1 Then
                slStr = "No"
            ElseIf imISCI = 2 Then      'yes and truncate
                slStr = "W/O Leader"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case REPINVINDEX   'Rep Inv
            pbcRepInv.Visible = False  'Set visibility
            If imRepInv = 0 Then
                slStr = "Internal"
            ElseIf imRepInv = 1 Then
                slStr = "External"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case INVSORTINDEX   'Invoice sorting
            lbcInvSort.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcInvSort.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcInvSort.List(lbcInvSort.ListIndex)
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case PACKAGEINDEX   'Package Invoice Show
            pbcPackage.Visible = False  'Set visibility
            If imPackage = 0 Then
                slStr = "Daypart"
            ElseIf imPackage = 1 Then
                slStr = "Time"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case REPMGINDEX   'Rep MG
            pbcRepMG.Visible = False  'Set visibility
            If imRepMG = 0 Then
                slStr = "Yes"
            ElseIf imRepMG = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case BONUSONINVINDEX   'Rates on Invoice
            pbcBonusOnInv.Visible = False  'Set visibility
            If imBonusOnInv = 0 Then
                slStr = "Yes"
            ElseIf imBonusOnInv = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case REFIDINDEX 'L.Bianchi 05/26/2021
            edcRefId.Visible = False
            slStr = edcRefId.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case DIRECTREFIDINDEX 'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
            edcDirectRefID.Visible = False
            slStr = edcDirectRefID.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case MEGAPHONEADVID
            edcMegaphoneAdvID.Visible = False
            slStr = edcMegaphoneAdvID.Text
            tmCtrls(MEGAPHONEADVID).sShow = edcMegaphoneAdvID.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case CRMIDINDEX
            edcCRMID.Visible = False
            tmCtrls(CRMIDINDEX).sShow = edcCRMID.Text
        
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case ADDRIDINDEX 'Address ID
            edcAddrID.Visible = False  'Set visibility
            slStr = edcAddrID.Text
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case BUYERINDEX 'Buyer
            lbcBuyer.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcBuyer.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcBuyer.List(lbcBuyer.ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case PAYABLEINDEX 'Buyer
            lbcPayable.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcPayable.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcPayable.List(lbcPayable.ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case LKBOXINDEX   'Lock box
            lbcLkBox.Visible = False  'Set visibility
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcLkBox.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcLkBox.List(lbcLkBox.ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case EDICINDEX   'EDI service for contract
            lbcEDI(0).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcEDI(0).ListIndex <= 0) Or (tgSpf.sAEDIC = "N") Then
                slStr = ""
            Else
                slStr = lbcEDI(0).List(lbcEDI(0).ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case EDIIINDEX   'EDI service for Invoices
            lbcEDI(1).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcEDI(1).ListIndex <= 0) Or (tgSpf.sAEDII = "N") Then
                slStr = ""
            Else
                slStr = lbcEDI(1).List(lbcEDI(1).ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
'        Case PRTSTYLEINDEX   'Print style
'            pbcPrtStyle.Visible = False  'Set visibility
'            If imPrtStyle = 0 Then
'                slStr = "Wide"
'            ElseIf imPrtStyle = 1 Then
'                slStr = "Narrow"
'            Else
'                slStr = ""
'            End If
'            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        Case TERMSINDEX   'Terms
            lbcTerms.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTerms.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcTerms.List(lbcTerms.ListIndex)
            End If
            gSetShow pbcAdvt, slStr, tmCtrls(ilBoxNo)
        Case TAXINDEX   'Sales tax
            lbcTax.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcTax.ListIndex < 0) Or (Not imTaxDefined) Then
                slStr = ""
            Else
                slStr = lbcTax.List(lbcTax.ListIndex)
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
        
        Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
            pbcSuppressNet.Visible = False
            If imSuppressNet = 0 Then
                slStr = "No"
            ElseIf imSuppressNet = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcDirect, slStr, tmCtrls(ilBoxNo)
            
            edcSuppressNet.Text = ""
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonBranch                  *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      salesperson and process        *
'*                      communication back from        *
'*                      salesperson                    *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mSPersonBranch() As Integer
'
'   ilRet = mSPersonBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilCallReturn As Integer
    Dim slName As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcSPerson, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        mSPersonBranch = False
        Exit Function
    End If
    If igWinStatus(SALESPEOPLELIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mSPersonBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'Screen.MousePointer = vbHourGlass  'Wait
    If (imDoubleClickName) And (lbcSPerson.ListIndex > UBound(tgSalesperson)) Then
        imCombo = True
        'igVsfCallSource = CALLSOURCEADVERTISER
        'sgVsfName = slStr
        'sgVsfCallType = "S"
        'mSPersonBranch = True
'        Combo.Show vbModal
    Else
    'If Not gWinRoom(igNoLJWinRes(SALESPEOPLELIST)) Then
    '    imDoubleClickName = False
    '    mSPersonBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
        imCombo = False
        igSlfCallSource = CALLSOURCEADVERTISER
        If edcDropDown.Text = "[New]" Then
            sgSlfName = ""
        Else
            sgSlfName = slStr
        End If
        ilUpdateAllowed = imUpdateAllowed
'        Advt.Enabled = False
        'igChildDone = False
        'edcLinkSrceDoneMsg.Text = ""
        'If (Not igStdAloneMode) And (imShowHelpMsg) Then
            If igTestSystem Then
                slStr = "Advt^Test\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
            Else
                slStr = "Advt^Prod\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
            End If
        'Else
        '    If igTestSystem Then
        '        slStr = "Advt^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
        '    Else
        '        slStr = "Advt^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
        '    End If
        'End If
        'lgShellRet = Shell(sgExePath & "SPerson.Exe " & slStr, 1)
        'Advt.Enabled = False
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        sgCommandStr = slStr
        SPerson.Show vbModal
        slStr = sgDoneMsg
        ilParse = gParseItem(slStr, 1, "\", sgSlfName)
        igSlfCallSource = Val(sgSlfName)
        ilParse = gParseItem(slStr, 2, "\", sgSlfName)
        'Advt.Enabled = True
        'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
        'For ilLoop = 0 To 10
        '    DoEvents
        'Next ilLoop
    End If
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSPersonBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If imCombo Then
        'ilCallReturn = igVsfCallSource
        'slName = sgVsfName
    Else
        ilCallReturn = igSlfCallSource
        slName = sgSlfName
    End If
    'igVsfCallSource = CALLNONE
    igSlfCallSource = CALLNONE
    'sgVsfName = ""
    sgSlfName = ""
    mSPersonBranch = True
    If ilCallReturn = CALLDONE Then  'Done
        lbcSPerson.Clear
        sgMSlfStamp = ""
        sgSalespersonTag = ""
        mSPersonPop
        If imTerminate Then
            mSPersonBranch = False
            Exit Function
        End If
        gFindMatch slName, 1, lbcSPerson
        If gLastFound(lbcSPerson) > 0 Then
            imChgMode = True
            lbcSPerson.ListIndex = gLastFound(lbcSPerson)
            edcDropDown.Text = lbcSPerson.List(lbcSPerson.ListIndex)
            imChgMode = False
            mSPersonBranch = False
            mSetChg SPERSONINDEX
        Else
            imChgMode = True
            lbcSPerson.ListIndex = 1
            edcDropDown.Text = lbcSPerson.List(1)
            imChgMode = False
            mSetChg SPERSONINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If ilCallReturn = CALLCANCELLED Then  'Cancelled
        mEnableBox imBoxNo
        Exit Function
    End If
    If ilCallReturn = CALLTERMINATED Then
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonPop                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales office list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop()
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilType As Integer
    Dim ilDormant As Integer
    ilIndex = lbcSPerson.ListIndex
    If ilIndex > 1 Then
        slName = lbcSPerson.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSPersonComboBox(Advt, lbcSPerson, Traffic!lbcSalesperson, Traffic!lbcSPersonCombo, igSlfFirstNameFirst)
    ilType = 0  'All
    If imSelectedIndex = 0 Then 'New selected
        ilDormant = False
    Else
        ilDormant = True
    End If
    'ilRet = gPopSalespersonBox(Advt, ilType, False, True, lbcSPerson, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(Advt, ilType, False, ilDormant, lbcSPerson, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcSPerson.AddItem "[None]", 0
        lbcSPerson.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcSPerson
            If gLastFound(lbcSPerson) > 1 Then
                lbcSPerson.ListIndex = gLastFound(lbcSPerson)
            Else
                lbcSPerson.ListIndex = -1
            End If
        Else
            lbcSPerson.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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

    sgDoneMsg = Trim$(str$(igAdvtCallSource)) & "\" & sgAdvtName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload Advt
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTermsBranch                  *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to invoice*
'*                      sorting and process            *
'*                      communication back from        *
'*                      invoice sorting                *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mTermsBranch() As Integer
'
'   ilRet = mTermsBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcTerms, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (StrComp(edcDropDown.Text, sgDefaultTerms, vbTextCompare) = 0) Then
        imDoubleClickName = False
        mTermsBranch = False
        Exit Function
    End If
    If igWinStatus(ADVERTISERSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mTermsBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mTermsBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "J"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Advt^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Advt^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mTermsBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcTerms.Clear
        smTermsCodeTag = ""
        mTermsPop
        If imTerminate Then
            mTermsBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcTerms
        sgMNmName = ""
        If gLastFound(lbcTerms) > 0 Then
            imChgMode = True
            lbcTerms.ListIndex = gLastFound(lbcTerms)
            edcDropDown.Text = lbcTerms.List(lbcTerms.ListIndex)
            imChgMode = False
            mTermsBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcTerms.ListIndex = 1
            edcDropDown.Text = lbcTerms.List(1)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTermsPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Terms list            *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mTermsPop()
'
'   mCompPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    ilFilter(0) = CHARFILTER
    slFilter(0) = "J"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilFilter(1) = CHARFILTER
    slFilter(1) = "T"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfUnitType") '2
    ilIndex = lbcTerms.ListIndex
    If ilIndex > 1 Then
        slName = lbcTerms.List(ilIndex)
    End If
    ilRet = gIMoveListBox(Advt, lbcTerms, tmTermsCode(), smTermsCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTermsPopErr
        gCPErrorMsg ilRet, "mTermsPop (gIMoveListBox)", Advt
        On Error GoTo 0
        lbcTerms.AddItem sgDefaultTerms, 0  'Force as first item on list
        lbcTerms.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcTerms
            If gLastFound(lbcTerms) > 1 Then
                lbcTerms.ListIndex = gLastFound(lbcTerms)
            Else
                lbcTerms.ListIndex = -1
            End If
        Else
            lbcTerms.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mTermsPopErr:
    On Error GoTo 0
    imTerminate = True
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
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String
    Dim ilRet As Integer

    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ABBRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAbbr, "", "Abbreviation must be specified", tmCtrls(ABBRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ABBRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = POLITICALINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imPolitical = 0 Then
            slStr = "Yes"
        ElseIf imPolitical = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Political must be specified", tmCtrls(ISCIINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ISCIINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imShortForm Then
            If imState < 0 Then
                imState = 0
            End If
        End If
        If imState = 0 Then
            slStr = "Active"
        ElseIf imState = 1 Then
            slStr = "Dormant"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Active/Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PRODINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smProduct, "", "Product must be specified", tmCtrls(PRODINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PRODINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SPERSONINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcSPerson, "", "Salesperson must be specified", tmCtrls(SPERSONINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SPERSONINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = AGENCYINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcAgency, "", "Agency must be specified", tmCtrls(AGENCYINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = AGENCYINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REPCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRepCode, "", "Rep Advertiser Code must be specified", tmCtrls(REPCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REPCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    slStr = Trim$(edcRepCode.Text)
    If Not gGPNoOk(slStr, tmAdf.iCode, 0, hmAdf, hmAgf) Then
        If (rbcBill(1).Value) Or (slStr <> "") Then
            If (ilState And SHOWMSG) = SHOWMSG Then
                ilRet = MsgBox("Rep Advertiser Code must be specified", vbOKOnly + vbExclamation, "Incomplete")
            End If
        End If
    End If
    If (ilCtrlNo = AGYCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAgyCode, "", "Agency Advertiser Code must be specified", tmCtrls(AGYCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = AGYCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STNCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcStnCode, "", "Station Advertiser Code must be specified", tmCtrls(STNCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STNCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcComp(0), "", "Competitive must be specified", tmCtrls(COMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If gFieldDefinedCtrl(lbcComp(1), "", "Competitive Code must be specified", tmCtrls(COMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If gFieldDefinedCtrl(lbcExcl(0), "", "Program Exclusion must be specified", tmCtrls(COMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If gFieldDefinedCtrl(lbcExcl(1), "", "Program Exclusion Code must be specified", tmCtrls(COMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DEMOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imPriceType = 0 Then
            slStr = "N/A"
        ElseIf imPriceType = 1 Then
            slStr = "CPM"
        ElseIf imPriceType = 2 Then
            slStr = "CPP"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Price type must be specified", tmCtrls(DEMOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DEMOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CREDITAPPROVALINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcCreditApproval, "", "Credit Approval must be specified", tmCtrls(CREDITAPPROVALINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CREDITAPPROVALINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CREDITRESTRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcCreditRestr, "", "Credit restriction must be specified", tmCtrls(CREDITRESTRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CREDITRESTRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If lbcCreditRestr.ListIndex = 1 Then
            If gFieldDefinedCtrl(edcCreditLimit, "", "Credit limit must be specified", tmCtrls(CREDITRESTRINDEX + 1).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CREDITRESTRINDEX  'Use credit restriction not credit limit
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = PAYMRATINGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcPaymRating, "", "Payment Rating must be specified", tmCtrls(PAYMRATINGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PAYMRATINGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CREDITRATINGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRating, "", "Credit Rating must be specified", tmCtrls(CREDITRATINGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CREDITRATINGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = RATEONINVINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imRateOnInv = 0 Then
            slStr = "Yes"
        ElseIf imRateOnInv = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Rate on Invoice must be specified", tmCtrls(RATEONINVINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = RATEONINVINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ISCIINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imISCI = 0 Then
            slStr = "Yes"
        ElseIf imISCI = 1 Then
            slStr = "No"
        ElseIf imISCI = 2 Then
            slStr = "W/O Leader"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "ISCI on Invoices must be specified", tmCtrls(ISCIINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ISCIINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REPINVINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imRepInv = 0 Then
            slStr = "Internal"
        ElseIf imRepInv = 1 Then
            slStr = "External"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Rep Inv must be specified", tmCtrls(REPINVINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REPINVINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = INVSORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcInvSort, "", "Invoice sort must be specified", tmCtrls(INVSORTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INVSORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PACKAGEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imPackage = 0 Then
            slStr = "Daypart"
        ElseIf imPackage = 1 Then
            slStr = "Time"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Package Invoice Show must be specified", tmCtrls(PACKAGEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PACKAGEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REPMGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imRepMG = 0 Then
            slStr = "Yes"
        ElseIf imRepMG = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Rep MG must be specified", tmCtrls(REPMGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REPMGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = BONUSONINVINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imBonusOnInv = 0 Then
            slStr = "Yes"
        ElseIf imBonusOnInv = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Bonus on Invoice must be specified", tmCtrls(BONUSONINVINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = BONUSONINVINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TESTALLCTRLS) And (rbcBill(0).Value) Then
        mTestFields = YES
        Exit Function
    End If
    If (ilCtrlNo = CADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcCAddr(0), "", "Contract Address must be specified", tmCtrls(CADDRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CADDRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = BADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcBAddr(0), "", "Billing Address must be specified", tmCtrls(BADDRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = BADDRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ADDRIDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (rbcBill(1).Value) And (tgSpf.sSystemType = "R") Then
            tmCtrls(ADDRIDINDEX).iReq = True
        Else
            tmCtrls(ADDRIDINDEX).iReq = False
        End If
        If gFieldDefinedCtrl(edcAddrID, "", "Address ID must be specified", tmCtrls(ADDRIDINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ADDRIDINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        tmCtrls(ADDRIDINDEX).iReq = False
    End If
    If (ilCtrlNo = BUYERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smBuyer, "", "Buyer must be specified", tmCtrls(BUYERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = BUYERINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PAYABLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smPayable, "", "Payable Contact must be specified", tmCtrls(PAYABLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PAYABLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = LKBOXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcLkBox, "", "Lock Box must be specified", tmCtrls(LKBOXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LKBOXINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = EDICINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcEDI(0), "", "EDI Service for Contracts must be specified", tmCtrls(EDICINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = EDICINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = EDIIINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcEDI(1), "", "EDI Service for Invoices must be specified", tmCtrls(EDIIINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = EDIIINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
'    If (ilCtrlNo = PRTSTYLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
'        If imPrtStyle = 0 Then
'            slStr = "Wide"
'        ElseIf imPrtStyle = 1 Then
'            slStr = "Narrow"
'        Else
'            slStr = ""
'        End If
'        If gFieldDefinedStr(slStr, "", "Contract Print Style on Invoices must be specified", tmCtrls(PRTSTYLEINDEX).iReq, ilState) = NO Then
'            If ilState = (ALLMANDEFINED + SHOWMSG) Then
'                imBoxNo = PRTSTYLEINDEX
'            End If
'            mTestFields = NO
'            Exit Function
'        End If
'    End If
    If (ilCtrlNo = TERMSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTerms, "", "Terms must be specified", tmCtrls(TERMSINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TERMSINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TAXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTax, "", "Sales Tax must be specified", tmCtrls(TAXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TAXINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
     'L.Bianchi 06/02/2021
     'If (ilCtrlNo = REFIDINDEX) Then
        'If gFieldDefinedStr(edcRefId.Text, "", "Ref Id must be specified", tmCtrls(REFIDINDEX).iReq, ilState) = NO Then
            'If ilState = (ALLMANDEFINED + SHOWMSG) Or ilState = SHOWMSG Then
                'imBoxNo = REFIDINDEX
            'End If
            'mTestFields = NO
            'Exit Function
        'End If
        
         'If gFieldDefinedGuidStr(edcRefId.Text, "", "Valid Ref Id identifier must be specified", ilState) = NO Then
            'If ilState = (ALLMANDEFINED + SHOWMSG) Or ilState = SHOWMSG Then
                'imBoxNo = REFIDINDEX
            'End If
            'mTestFields = NO
            'Exit Function
        'End If
        
    'End If
    mTestFields = YES
End Function

Private Sub lbcTerms_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcTerms, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcTerms_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcTerms_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcTerms_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcTerms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcTerms, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub pbcAdvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer

    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To CRMIDINDEX Step 1 'L.Bianchi 05/31/2021, JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
            
                If ilBox = BKOUTPOOLINDEX Then
                    Beep
                    Exit Sub
                End If
                If (ilBox = REPCODEINDEX) And (tgSpf.sARepCodes = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = AGYCODEINDEX) And (tgSpf.sAAgyCodes = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = STNCODEINDEX) And (tgSpf.sAStnCodes = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = DEMOINDEX) And (tgSpf.iATargets = 0) Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITAPPROVALINDEX) And (tgUrf(0).sChgCrRt <> "I") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = ISCIINDEX) And (tgSpf.sAISCI = "A") Then
                    imISCI = 0
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = ISCIINDEX) And (tgSpf.sAISCI = "X") Then
                    imISCI = 1
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = PACKAGEINDEX) And (tgSpf.sCPkOrdered = "N") And (tgSpf.sCPkAired = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX) And (tgUrf(0).sCredit <> "I") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX + 1) And (tgUrf(0).sCredit <> "I") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = PAYMRATINGINDEX) And (tgUrf(0).sPayRate <> "I") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX + 1) Then
                    ilBox = CREDITRESTRINDEX
                End If
                mCESetShow imCEBoxNo
                imCEBoxNo = -1
                mDmSetShow imDMBoxNo
                imDMBoxNo = -1
                imBillTabDir = 0
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcAdvt_Paint()
    Dim ilBox As Integer
    mPaintAdvtTitle
'    Dim llColor As Long
    'For ilBox = imLBCtrls To PACKAGEINDEX Step 1
    For ilBox = imLBCtrls To CRMIDINDEX Step 1 'L.Bianchi 05/31/2021, JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
        If (ilBox <> CREDITAPPROVALINDEX) Or (tgUrf(0).sChgCrRt <> "H") Then
            If (ilBox <> PAYMRATINGINDEX) Or (tgUrf(0).sPayRate <> "H") Then
                If (ilBox <> CREDITRESTRINDEX) Or (tgUrf(0).sCredit <> "H") Then
                    If (ilBox <> CREDITRESTRINDEX + 1) Or (tgUrf(0).sCredit <> "H") Then
                        If ilBox = DEMOINDEX Then
                            gPaintArea pbcAdvt, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + fgOffset - 15, tmCtrls(ilBox).fBoxW - 15, fgBoxGridH, pbcAdvt.BackColor    '  WHITE
    '                        gInvalidateArea pbcAdvt, Int(tmCtrls(ilBox).fBoxX), Int(tmCtrls(ilBox).fBoxY + fgOffSet - 15), Int(tmCtrls(ilBox).fBoxW - 15), Int(fgBoxGridH)
                        End If
                        pbcAdvt.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                        pbcAdvt.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
                        pbcAdvt.Print tmCtrls(ilBox).sShow
                        
'                        Debug.Print ilBox & " - " & tmCtrls(ilBox).sShow & " ," & tmCtrls(ilBox).fBoxX + fgBoxInsetX & " , " & tmCtrls(ilBox).fBoxY + fgBoxInsetY
                    End If
                End If
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcBillTab_GotFocus(Index As Integer)
    If Index = 0 Then
        If imBillTabDir = 0 Then
            pbcTab.SetFocus
        Else
            pbcSTab.SetFocus
        End If
    Else
        If imBillTabDir = 0 Then
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub pbcBonusOnInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imBonusOnInv <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imBonusOnInv = 0
        pbcBonusOnInv_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imBonusOnInv <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imBonusOnInv = 1
        pbcBonusOnInv_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imBonusOnInv = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imBonusOnInv = 1
            pbcBonusOnInv_Paint
        ElseIf imBonusOnInv = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imBonusOnInv = 0
            pbcBonusOnInv_Paint
        End If
    End If
    mSetCommands

End Sub

Private Sub pbcBonusOnInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBonusOnInv = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imBonusOnInv = 1
    ElseIf imBonusOnInv = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imBonusOnInv = 0
    End If
    pbcBonusOnInv_Paint
    mSetCommands
End Sub

Private Sub pbcBonusOnInv_Paint()
    pbcBonusOnInv.Cls
    pbcBonusOnInv.CurrentX = fgBoxInsetX
    pbcBonusOnInv.CurrentY = 0 'fgBoxInsetY
    If imBonusOnInv = 0 Then
        pbcBonusOnInv.Print "Yes"
    ElseIf imBonusOnInv = 1 Then
        pbcBonusOnInv.Print "No"
    Else
        pbcBonusOnInv.Print "   "
    End If
End Sub

Private Sub pbcCE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBCECtrls To UBound(tmCECtrls) Step 1
        If (X >= tmCECtrls(ilBox).fBoxX) And (X <= tmCECtrls(ilBox).fBoxX + tmCECtrls(ilBox).fBoxW) Then
            If (Y >= tmCECtrls(ilBox).fBoxY) And (Y <= tmCECtrls(ilBox).fBoxY + tmCECtrls(ilBox).fBoxH) Then
                If (tgSpf.sAExcl <> "Y") And ((ilBox = CEEXCLINDEX) Or (ilBox = CEEXCLINDEX + 1)) Then
                    Beep
                    Exit Sub
                End If
                mCESetShow imCEBoxNo
                imCEBoxNo = ilBox
                mCEEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcCE_Paint()
    Dim ilBox As Integer
'    Dim llColor As Long

    For ilBox = imLBCECtrls To UBound(tmCECtrls) Step 1
        pbcCE.CurrentX = tmCECtrls(ilBox).fBoxX + fgBoxInsetX
        pbcCE.CurrentY = tmCECtrls(ilBox).fBoxY + fgBoxInsetY
        pbcCE.Print tmCECtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcCESTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcCESTab.HWnd Then
        Exit Sub
    End If
    If (imCEBoxNo = CECOMPINDEX) Or (imBoxNo = CECOMPINDEX + 1) Then
        If mCompBranch(imCEBoxNo - CECOMPINDEX) Then
            Exit Sub
        End If
    End If
    If (imCEBoxNo = CEEXCLINDEX) Or (imBoxNo = CEEXCLINDEX + 1) Then
        If mExclBranch(imCEBoxNo - CEEXCLINDEX) Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCECtrls) And (imCEBoxNo <= UBound(tmCECtrls)) Then
        If mCETestFields(imCEBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mCEEnableBox imCEBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imCEBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            ilBox = 1
        Case CECOMPINDEX ' (first control within header)
            mCESetShow imCEBoxNo
            imCEBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imCEBoxNo - 1
    End Select
    mCESetShow imCEBoxNo
    imCEBoxNo = ilBox
    mCEEnableBox ilBox
End Sub
Private Sub pbcCETab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcCETab.HWnd Then
        Exit Sub
    End If
    If (imCEBoxNo = CECOMPINDEX) Or (imCEBoxNo = CECOMPINDEX + 1) Then
        If mCompBranch(imCEBoxNo - CECOMPINDEX) Then
            Exit Sub
        End If
    End If
    If (imCEBoxNo = CEEXCLINDEX) Or (imCEBoxNo = CEEXCLINDEX + 1) Then
        If mExclBranch(imCEBoxNo - CEEXCLINDEX) Then
            Exit Sub
        End If
    End If
    If (imCEBoxNo >= imLBCECtrls) And (imCEBoxNo <= UBound(tmCECtrls)) Then
        If mCETestFields(imCEBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mCEEnableBox imCEBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imCEBoxNo
        Case -1 'Shift tab from button
            imTabDirection = -1  'Set-Right to left
            If tgSpf.sAExcl = "Y" Then
                ilBox = CEEXCLINDEX + 1
            Else
                ilBox = CECOMPINDEX + 1
            End If
        Case CECOMPINDEX + 1
            If tgSpf.sAExcl = "Y" Then
                ilBox = CEEXCLINDEX
            Else
                mCESetShow imCEBoxNo
                imCEBoxNo = -1
                pbcTab.SetFocus
                Exit Sub
            End If
        Case CEEXCLINDEX + 1
            mCESetShow imCEBoxNo
            imCEBoxNo = -1
            pbcTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imCEBoxNo + 1
    End Select
    mCESetShow imCEBoxNo
    imCEBoxNo = ilBox
    mCEEnableBox ilBox
End Sub
Private Sub pbcClickFocus_GotFocus()
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDirect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim flAdj As Single
    If Not rbcBill(1).Value Then
        Exit Sub
    End If
    For ilBox = CADDRINDEX To SUPPRESSNETINDEX Step 1 ' TTP 10622 - 2023-03-08 JJB
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (ilBox = CADDRINDEX + 1) Or (ilBox = CADDRINDEX + 2) Or (ilBox = BADDRINDEX + 1) Or (ilBox = BADDRINDEX + 2) Then
                flAdj = fgBoxInsetY
            Else
                flAdj = 0
            End If
            If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                If (ilBox = EDICINDEX) And (tgSpf.sAEDIC = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
               
                
                If (ilBox = EDIIINDEX) And (tgSpf.sAEDII = "N") Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
'                If (ilBox = PRTSTYLEINDEX) And (tgSpf.sAPrtStyle = "W") Then
'                    imPrtStyle = 0
'                    If imCEBoxNo > 0 Then
'                        mCESetFocus imCEBoxNo
'                    ElseIf imDMBoxNo > 0 Then
'                        mDmSetFocus imDMBoxNo
'                    Else
'                        mSetFocus imBoxNo
'                    End If
'                    Beep
'                    Exit Sub
'                End If
'                If (ilBox = PRTSTYLEINDEX) And (tgSpf.sAPrtStyle = "N") Then
'                    imPrtStyle = 1
'                    If imCEBoxNo > 0 Then
'                        mCESetFocus imCEBoxNo
'                    ElseIf imDMBoxNo > 0 Then
'                        mDmSetFocus imDMBoxNo
'                    Else
'                        mSetFocus imBoxNo
'                    End If
'                    Beep
'                    Exit Sub
'                End If
                If (ilBox = TAXINDEX) And (Not imTaxDefined) Then
                    If imCEBoxNo > 0 Then
                        mCESetFocus imCEBoxNo
                    ElseIf imDMBoxNo > 0 Then
                        mDmSetFocus imDMBoxNo
                    Else
                        mSetFocus imBoxNo
                    End If
                    Beep
                    Exit Sub
                End If
                mCESetShow imCEBoxNo
                imCEBoxNo = -1
                mDmSetShow imDMBoxNo
                imDMBoxNo = -1
                imBillTabDir = 0
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcDirect_Paint()
    Dim ilBox As Integer
'    Dim llColor As Long

    If Not rbcBill(1).Value Then
        Exit Sub
    End If
    mPaintDirectTitle
    
    For ilBox = CADDRINDEX To SUPPRESSNETINDEX Step 1 ' TTP 10622 - 2023-03-08 JJB
        pbcDirect.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDirect.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDirect.Print tmCtrls(ilBox).sShow
    Next ilBox
    
'    llColor = pbcDirect.ForeColor
'    pbcDirect.ForeColor = BLUE
    'pbcDirect.CurrentX = 30 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smPct90
    'pbcDirect.CurrentX = 1560 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smCurrAR
    'pbcDirect.CurrentX = 3090 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smUnbilled
    'pbcDirect.CurrentX = 4620 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smHiCredit
    'pbcDirect.CurrentX = 6150 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smTotalGross
    'pbcDirect.CurrentX = 7440 + fgBoxInsetX
    'pbcDirect.CurrentY = 1800 + fgBoxInsetY
    'pbcDirect.Print smDateEntrd
    'pbcDirect.CurrentX = 30 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
    'pbcDirect.Print smNSFChks
    'pbcDirect.CurrentX = 1560 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
    'pbcDirect.Print smDateLstInv
    'pbcDirect.CurrentX = 3090 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
    'pbcDirect.Print smDateLstPaym
    'pbcDirect.CurrentX = 4620 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
    'pbcDirect.Print smAvgToPay
    'pbcDirect.CurrentX = 6150 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
    'pbcDirect.Print smLstToPay
    'pbcDirect.CurrentX = 7440 + fgBoxInsetX
    'pbcDirect.CurrentY = 2145 + fgBoxInsetY
'    pbcDirect.ForeColor = llColor
    pbcDirect.CurrentX = tmDTCtrls(PCT90INDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(PCT90INDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smPct90
    pbcDirect.CurrentX = tmDTCtrls(CURRARINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(CURRARINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smCurrAR
    pbcDirect.CurrentX = tmDTCtrls(UNBILLEDINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(UNBILLEDINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smUnbilled
    pbcDirect.CurrentX = tmDTCtrls(HICREDITINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(HICREDITINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smHiCredit
    pbcDirect.CurrentX = tmDTCtrls(TOTALGROSSINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(TOTALGROSSINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smTotalGross
    pbcDirect.CurrentX = tmDTCtrls(DATEENTRDINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(DATEENTRDINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smDateEntrd
    pbcDirect.CurrentX = tmDTCtrls(NSFCHKSINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(NSFCHKSINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smNSFChks
    pbcDirect.CurrentX = tmDTCtrls(DATELSTINVINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(DATELSTINVINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smDateLstInv
    pbcDirect.CurrentX = tmDTCtrls(DATELSTPAYMINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(DATELSTPAYMINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smDateLstPaym
    pbcDirect.CurrentX = tmDTCtrls(AVGTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(AVGTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smAvgToPay
    pbcDirect.CurrentX = tmDTCtrls(LSTTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcDirect.CurrentY = tmDTCtrls(LSTTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcDirect.Print smLstToPay

End Sub
Private Sub pbcDm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        If (X >= tmDmCtrls(ilBox).fBoxX) And (X <= tmDmCtrls(ilBox).fBoxX + tmDmCtrls(ilBox).fBoxW) Then
            If (Y >= tmDmCtrls(ilBox).fBoxY) And (Y <= tmDmCtrls(ilBox).fBoxY + tmDmCtrls(ilBox).fBoxH) Then
                If (imPriceType <= 0) And (ilBox > DMPRICETYPEINDEX) Then
                    Beep
                    Exit Sub
                End If
                If ilBox > DMVALUEINDEX + 2 * (tgSpf.iATargets - 1) Then
                    Beep
                    Exit Sub
                End If
                mDmSetShow imDMBoxNo
                imDMBoxNo = ilBox
                mDmEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcDm_Paint()
    Dim ilBox As Integer
'    Dim llColor As Long

    For ilBox = imLBDmCtrls To UBound(tmDmCtrls) Step 1
        pbcDm.CurrentX = tmDmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDm.CurrentY = tmDmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDm.Print tmDmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcDmPriceType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcDmPriceType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imPriceType <> 0 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
        End If
        imPriceType = 0
        pbcDmPriceType_Paint
    ElseIf KeyAscii = Asc("M") Or (KeyAscii = Asc("m")) Then
        If imPriceType <> 1 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
        End If
        imPriceType = 1
        pbcDmPriceType_Paint
    ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
        If imPriceType <> 2 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
        End If
        imPriceType = 2
        pbcDmPriceType_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imPriceType = 0 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
            imPriceType = 1
            pbcDmPriceType_Paint
        ElseIf imPriceType = 1 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
            imPriceType = 2
            pbcDmPriceType_Paint
        ElseIf imPriceType = 2 Then
            tmDmCtrls(imDMBoxNo).iChg = True
            tmCtrls(DEMOINDEX).iChg = True
            imPriceType = 0
            pbcDmPriceType_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcDmPriceType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imPriceType = 0 Then
        tmDmCtrls(imDMBoxNo).iChg = True
        tmCtrls(DEMOINDEX).iChg = True
        imPriceType = 1
    ElseIf imPriceType = 1 Then
        tmDmCtrls(imDMBoxNo).iChg = True
        tmCtrls(DEMOINDEX).iChg = True
        imPriceType = 2
    ElseIf imPriceType = 2 Then
        tmDmCtrls(imDMBoxNo).iChg = True
        tmCtrls(DEMOINDEX).iChg = True
        imPriceType = 0
    End If
    pbcDmPriceType_Paint
    mSetCommands
End Sub
Private Sub pbcDmPriceType_Paint()
    pbcDmPriceType.Cls
    pbcDmPriceType.CurrentX = fgBoxInsetX
    pbcDmPriceType.CurrentY = 0 'fgBoxInsetY
    If imPriceType = 0 Then
        pbcDmPriceType.Print "N/A"
    ElseIf imPriceType = 1 Then
        pbcDmPriceType.Print "CPM"
    ElseIf imPriceType = 2 Then
        pbcDmPriceType.Print "CPP"
    Else
        pbcDmPriceType.Print "   "
    End If
End Sub
Private Sub pbcDmSTab_GotFocus()
    Dim ilBox As Integer
    If (imDMBoxNo >= imLBDmCtrls) And (imDMBoxNo <= UBound(tmDmCtrls)) Then
        If mDmTestFields(imDMBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mDmEnableBox imDMBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imDMBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            ilBox = 1
        Case DMPRICETYPEINDEX ' (first control within header)
            mDmSetShow imDMBoxNo
            imDMBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imDMBoxNo - 1
    End Select
    mDmSetShow imDMBoxNo
    imDMBoxNo = ilBox
    mDmEnableBox ilBox
End Sub
Private Sub pbcDmTab_GotFocus()
    Dim ilBox As Integer
    If (imDMBoxNo >= imLBDmCtrls) And (imDMBoxNo <= UBound(tmDmCtrls)) Then
        If mDmTestFields(imDMBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mDmEnableBox imDMBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imDMBoxNo
        Case -1 'Shift tab from button
            imTabDirection = -1  'Set-Right to left
            If imPriceType = 0 Then
                ilBox = DMPRICETYPEINDEX
            Else
                ilBox = DMVALUEINDEX + 2 * (tgSpf.iATargets - 1)
            End If
        Case DMPRICETYPEINDEX
            If imPriceType = 0 Then
                mDmSetShow imDMBoxNo
                imDMBoxNo = -1
                pbcTab.SetFocus
                Exit Sub
            End If
            ilBox = imDMBoxNo + 1
        Case DMDEMOINDEX
            If lbcDemo(0).ListIndex = 0 Then
                smTarget(0) = ""
                If tgSpf.iATargets <= 1 Then
                    mDmSetShow imDMBoxNo
                    imDMBoxNo = -1
                    pbcTab.SetFocus
                    Exit Sub
                Else
                    ilBox = DMDEMOINDEX + 2
                End If
            Else
                ilBox = imDMBoxNo + 1
            End If
        Case DMDEMOINDEX + 2
            If lbcDemo(1).ListIndex = 0 Then
                smTarget(1) = ""
                If tgSpf.iATargets <= 2 Then
                    mDmSetShow imDMBoxNo
                    imDMBoxNo = -1
                    pbcTab.SetFocus
                    Exit Sub
                Else
                    ilBox = DMDEMOINDEX + 4
                End If
            Else
                ilBox = imDMBoxNo + 1
            End If
        Case DMDEMOINDEX + 4
            If lbcDemo(2).ListIndex = 0 Then
                smTarget(2) = ""
                If tgSpf.iATargets <= 3 Then
                    mDmSetShow imDMBoxNo
                    imDMBoxNo = -1
                    pbcTab.SetFocus
                    Exit Sub
                Else
                    ilBox = DMDEMOINDEX + 6
                End If
            Else
                ilBox = imDMBoxNo + 1
            End If
        Case DMDEMOINDEX + 6
            If lbcDemo(3).ListIndex = 0 Then
                smTarget(3) = ""
                mDmSetShow imDMBoxNo
                imDMBoxNo = -1
                pbcTab.SetFocus
                Exit Sub
            Else
                ilBox = imDMBoxNo + 1
            End If
        Case DMVALUEINDEX + 2 * (tgSpf.iATargets - 1)
            mDmSetShow imDMBoxNo
            imDMBoxNo = -1
            pbcTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imDMBoxNo + 1
    End Select
    mDmSetShow imDMBoxNo
    imDMBoxNo = ilBox
    mDmEnableBox ilBox
End Sub
Private Sub pbcISCI_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcISCI_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imISCI <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imISCI = 0
        pbcISCI_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imISCI <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imISCI = 1
        pbcISCI_Paint
    ElseIf KeyAscii = Asc("W") Or (KeyAscii = Asc("w")) Then        '3-10-15 addl flag to show isci but remove hard-coded WW_
        If imISCI <> 2 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imISCI = 2
        pbcISCI_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imISCI = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imISCI = 1
            pbcISCI_Paint
        ElseIf imISCI = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imISCI = 2
            pbcISCI_Paint
        ElseIf imISCI = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imISCI = 0
            pbcISCI_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcISCI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imISCI = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imISCI = 1
    ElseIf imISCI = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imISCI = 2
    ElseIf imISCI = 2 Then
        tmCtrls(imBoxNo).iChg = True
        imISCI = 0
    End If
    pbcISCI_Paint
    mSetCommands
End Sub
Private Sub pbcISCI_Paint()
    pbcISCI.Cls
    pbcISCI.CurrentX = fgBoxInsetX
    pbcISCI.CurrentY = 0 'fgBoxInsetY
    If imISCI = 0 Then
        pbcISCI.Print "Yes"
    ElseIf imISCI = 1 Then
        pbcISCI.Print "No"
    ElseIf imISCI = 2 Then
        pbcISCI.Print "W/O Leader"
    Else
        pbcISCI.Print "   "
    End If
End Sub
Private Sub pbcNotDirect_Paint()
'    Dim llColor As Long

'    llColor = pbcDirect.ForeColor
'    pbcDirect.ForeColor = BLUE
    mPaintNotDirectTitle
    pbcNotDirect.CurrentX = tmNDTCtrls(PCT90INDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(PCT90INDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smPct90
    pbcNotDirect.CurrentX = tmNDTCtrls(CURRARINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(CURRARINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smCurrAR
    pbcNotDirect.CurrentX = tmNDTCtrls(UNBILLEDINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(UNBILLEDINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smUnbilled
    pbcNotDirect.CurrentX = tmNDTCtrls(HICREDITINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(HICREDITINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smHiCredit
    pbcNotDirect.CurrentX = tmNDTCtrls(TOTALGROSSINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(TOTALGROSSINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smTotalGross
    pbcNotDirect.CurrentX = tmNDTCtrls(DATEENTRDINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(DATEENTRDINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smDateEntrd
    pbcNotDirect.CurrentX = tmNDTCtrls(NSFCHKSINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(NSFCHKSINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smNSFChks
    pbcNotDirect.CurrentX = tmNDTCtrls(DATELSTINVINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(DATELSTINVINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smDateLstInv
    pbcNotDirect.CurrentX = tmNDTCtrls(DATELSTPAYMINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(DATELSTPAYMINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smDateLstPaym
    pbcNotDirect.CurrentX = tmNDTCtrls(AVGTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(AVGTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smAvgToPay
    pbcNotDirect.CurrentX = tmNDTCtrls(LSTTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcNotDirect.CurrentY = tmNDTCtrls(LSTTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcNotDirect.Print smLstToPay
'    pbcDirect.ForeColor = llColor
End Sub
Private Sub pbcPackage_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcPackage_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imPackage <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imPackage = 0
        pbcPackage_Paint
    ElseIf KeyAscii = Asc("T") Or (KeyAscii = Asc("t")) Then
        If imPackage <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imPackage = 1
        pbcPackage_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imPackage = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imPackage = 1
            pbcPackage_Paint
        ElseIf imPackage = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imPackage = 0
            pbcPackage_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcPackage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imPackage = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imPackage = 1
    ElseIf imPackage = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imPackage = 0
    End If
    pbcPackage_Paint
    mSetCommands
End Sub
Private Sub pbcPackage_Paint()
    pbcPackage.Cls
    pbcPackage.CurrentX = fgBoxInsetX
    pbcPackage.CurrentY = 0 'fgBoxInsetY
    If imPackage = 0 Then
        pbcPackage.Print "Daypart"
    ElseIf imPackage = 1 Then
        pbcPackage.Print "Time"
    Else
        pbcPackage.Print "   "
    End If
End Sub

Private Sub pbcPolitical_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcPolitical_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imPolitical <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imPolitical = 0
        pbcPolitical_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imPolitical <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imPolitical = 1
        pbcPolitical_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imPolitical = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imPolitical = 1
            pbcPolitical_Paint
        ElseIf imPolitical = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imPolitical = 0
            pbcPolitical_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcPolitical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imPolitical = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imPolitical = 1
    ElseIf imPolitical = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imPolitical = 0
    End If
    pbcPolitical_Paint
    mSetCommands
End Sub

Private Sub pbcPolitical_Paint()
    pbcPolitical.Cls
    pbcPolitical.CurrentX = fgBoxInsetX
    pbcPolitical.CurrentY = 0 'fgBoxInsetY
    If imPolitical = 0 Then
        pbcPolitical.Print "Yes"
    ElseIf imPolitical = 1 Then
        pbcPolitical.Print "No"
    Else
        pbcPolitical.Print "   "
    End If
End Sub

Private Sub pbcRateOnInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imRateOnInv <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRateOnInv = 0
        pbcRateOnInv_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imRateOnInv <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRateOnInv = 1
        pbcRateOnInv_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imRateOnInv = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imRateOnInv = 1
            pbcRateOnInv_Paint
        ElseIf imRateOnInv = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imRateOnInv = 0
            pbcRateOnInv_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcRateOnInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imRateOnInv = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imRateOnInv = 1
    ElseIf imRateOnInv = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imRateOnInv = 0
    End If
    pbcRateOnInv_Paint
    mSetCommands
End Sub
Private Sub pbcRateOnInv_Paint()
    pbcRateOnInv.Cls
    pbcRateOnInv.CurrentX = fgBoxInsetX
    pbcRateOnInv.CurrentY = 0 'fgBoxInsetY
    If imRateOnInv = 0 Then
        pbcRateOnInv.Print "Yes"
    ElseIf imRateOnInv = 1 Then
        pbcRateOnInv.Print "No"
    Else
        pbcRateOnInv.Print "   "
    End If
End Sub

Private Sub pbcRepInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("I") Or (KeyAscii = Asc("i")) Then
        If imRepInv <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRepInv = 0
        pbcRepInv_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        If imRepInv <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRepInv = 1
        pbcRepInv_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imRepInv = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imRepInv = 1
            pbcRepInv_Paint
        ElseIf imRepInv = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imRepInv = 0
            pbcRepInv_Paint
        End If
    End If
    mSetCommands

End Sub

Private Sub pbcRepInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imRepInv = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imRepInv = 1
    ElseIf imRepInv = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imRepInv = 0
    End If
    pbcRepInv_Paint
    mSetCommands
End Sub

Private Sub pbcRepInv_Paint()
    pbcRepInv.Cls
    pbcRepInv.CurrentX = fgBoxInsetX
    pbcRepInv.CurrentY = 0 'fgBoxInsetY
    If imRepInv = 0 Then
        pbcRepInv.Print "Internal"
    ElseIf imRepInv = 1 Then
        pbcRepInv.Print "External"
    Else
        pbcRepInv.Print "   "
    End If
End Sub

Private Sub pbcRepMG_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imRepMG <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRepMG = 0
        pbcRepMG_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imRepMG <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imRepMG = 1
        pbcRepMG_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imRepMG = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imRepMG = 1
            pbcRepMG_Paint
        ElseIf imRepMG = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imRepMG = 0
            pbcRepMG_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcRepMG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imRepMG = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imRepMG = 1
    ElseIf imRepMG = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imRepMG = 0
    End If
    pbcRepMG_Paint
    mSetCommands
End Sub

Private Sub pbcRepMG_Paint()
    pbcRepMG.Cls
    pbcRepMG.CurrentX = fgBoxInsetX
    pbcRepMG.CurrentY = 0 'fgBoxInsetY
    If imRepMG = 0 Then
        pbcRepMG.Print "Yes"
    ElseIf imRepMG = 1 Then
        pbcRepMG.Print "No"
    Else
        pbcRepMG.Print "   "
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    imBillTabDir = 0
    
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = AGENCYINDEX Then
        If mAgencyBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = BUYERINDEX Then
        If mPersonnelBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PAYABLEINDEX Then
        If mPersonnelBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = LKBOXINDEX Then
        If mLkBoxBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = EDICINDEX Then
        If mEDIBranch(0) Then
            Exit Sub
        End If
    End If
    If imBoxNo = EDIIINDEX Then
        If mEDIBranch(1) Then
            Exit Sub
        End If
    End If
    If imBoxNo = INVSORTINDEX Then
        If mInvSortBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = TERMSINDEX Then
        If mTermsBranch() Then
            Exit Sub
        End If
    End If
    
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxNoCtrls) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                    ilBox = 1
                    mSetCommands
                Else
                    mSetChg 1
                    ilBox = 2
                End If
            Case NAMEINDEX 'Name (first control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                If (cmcUpdate.Enabled) And (igAdvtCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case STATEINDEX
                If imShortForm Then
                    If imState < 0 Then
                        tmCtrls(STATEINDEX).iChg = True
                        imState = 0
                    End If
                    ilFound = False
                End If
                ilBox = POLITICALINDEX
            Case PRODINDEX
                If imShortForm Then
                    If Trim$(smProduct) = "" Then
                        smProduct = "[None]"
                    End If
                    ilFound = False
                End If
                ilBox = STATEINDEX
            Case AGYCODEINDEX
                If (tgSpf.sARepCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = REPCODEINDEX
            Case STNCODEINDEX
                If (tgSpf.sAAgyCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = AGYCODEINDEX
            Case COMPINDEX
                If (tgSpf.sAStnCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = STNCODEINDEX
            Case CREDITAPPROVALINDEX
                If tgSpf.iATargets = 0 Then
                    ilBox = COMPINDEX
                    If (tgSpf.iATargets = 0) Or (imPriceType < 0) Then
                        imPriceType = 0
                    End If
                Else
                    ilBox = DEMOINDEX
                End If
                If imShortForm Then
                    ilFound = False
                End If
            Case CREDITRESTRINDEX
                If tgUrf(0).sChgCrRt <> "I" Then
                    If tgSpf.iATargets = 0 Then
                        ilBox = COMPINDEX
                        If (tgSpf.iATargets = 0) Or (imPriceType < 0) Then
                            imPriceType = 0
                        End If
                    Else
                        ilBox = DEMOINDEX
                    End If
                Else
                    ilBox = CREDITAPPROVALINDEX
                End If
                If imShortForm Then
                    ilFound = False
                    If lbcCreditApproval.ListIndex < 0 Then
                        imChgMode = True
                        If (tgUrf(0).sChgCrRt <> "I") Then
                            lbcCreditApproval.ListIndex = 0   'Requires Checking
                        Else
                            lbcCreditApproval.ListIndex = 1   'Approved
                        End If
                        imChgMode = False
                        tmCtrls(CREDITAPPROVALINDEX).iChg = True
                    End If
                End If
            Case CREDITRESTRINDEX + 1
                If (tgUrf(0).sCredit <> "I") Or imShortForm Then
                    ilFound = False
                    If lbcCreditRestr.ListIndex < 0 Then
                        imChgMode = True
                        lbcCreditRestr.ListIndex = 0
                        imChgMode = False
                        tmCtrls(CREDITRESTRINDEX).iChg = True
                    End If
                End If
                ilBox = CREDITRESTRINDEX
            Case PAYMRATINGINDEX
                If (tgUrf(0).sCredit <> "I") Or imShortForm Then
                    ilFound = False
                Else
                    If lbcCreditRestr.ListIndex <> 1 Then
                        ilFound = False
                    End If
                End If
                ilBox = CREDITRESTRINDEX + 1
            Case CREDITAPPROVALINDEX
                If (tgSpf.iATargets = 0) Or imShortForm Then
                    ilFound = False
                    If (tgSpf.iATargets = 0) Or (imPriceType < 0) Then
                        imPriceType = 0
                    End If
                End If
                ilBox = DEMOINDEX
            Case CREDITRATINGINDEX
                If (tgUrf(0).sPayRate <> "I") Or imShortForm Then
                    ilFound = False
                    If lbcPaymRating.ListIndex < 0 Then
                        imChgMode = True
                        lbcPaymRating.ListIndex = 1
                        imChgMode = False
                        tmCtrls(PAYMRATINGINDEX).iChg = True
                    End If
                End If
                ilBox = PAYMRATINGINDEX
            Case REPINVINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = CREDITRATINGINDEX
            Case INVSORTINDEX
                If imShortForm Then
                    If imRepInv < 0 Then
                        tmCtrls(REPINVINDEX).iChg = True
                        'imRepInv = 0
                        If (igInternalAdfCount <> 0) And (igInternalAdfCount < UBound(tgCommAdf) / 2) Then
                            imRepInv = 1    'External
                        Else
                            imRepInv = 0    'internal
                        End If
                    End If
                    ilFound = False
                End If
                ilBox = REPINVINDEX 'CREDITRATINGINDEX
            Case RATEONINVINDEX
                If (lbcInvSort.ListCount = 2) Or imShortForm Then
                    If lbcInvSort.ListIndex < 0 Then
                        imChgMode = True
                        lbcInvSort.ListIndex = 1
                        imChgMode = False
                    End If
                    ilFound = False
                End If
                ilBox = INVSORTINDEX
            Case ISCIINDEX
                If imShortForm Then
                    If imRateOnInv < 0 Then
                        tmCtrls(RATEONINVINDEX).iChg = True
                        imRateOnInv = 0
                    End If
                    ilFound = False
                End If
                ilBox = RATEONINVINDEX
            'Case INVSORTINDEX
            '    If tgSpf.sAISCI = "A" Then
            '        If imISCI < 0 Then
            '            imISCI = 0
            '            tmCtrls(ISCIINDEX).iChg = True
            '        End If
            '        ilFound = False
            '    ElseIf tgSpf.sAISCI = "X" Then
            '        If imISCI < 0 Then
            '            imISCI = 1
            '            tmCtrls(ISCIINDEX).iChg = True
            '        End If
            '        ilFound = False
            '    End If
            '    If imShortForm Then
            '        If imISCI < 0 Then
            '            tmCtrls(ISCIINDEX).iChg = True
            '            If tgSpf.sAISCI = "Y" Then
            '                imISCI = 0
            '            Else
            '                imISCI = 1
            '            End If
            '        End If
            '        ilFound = False
            '    End If
            '    ilBox = ISCIINDEX
            Case PACKAGEINDEX
                If tgSpf.sAISCI = "A" Then
                    If imISCI < 0 Then
                        imISCI = 0
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                ElseIf tgSpf.sAISCI = "X" Then
                    If imISCI < 0 Then
                        imISCI = 1
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                End If
                If imShortForm Then
                    If imISCI < 0 Then
                        tmCtrls(ISCIINDEX).iChg = True
                        If tgSpf.sAISCI = "Y" Then
                            imISCI = 0
                        Else
                            imISCI = 1
                        End If
                    End If
                    ilFound = False
                End If
                ilBox = ISCIINDEX
            Case REPMGINDEX
                If (Not imShortForm) And ((tgSpf.sCPkOrdered = "Y") Or (tgSpf.sCPkAired = "Y")) Then
                    ilBox = PACKAGEINDEX
                Else
                    If tgSpf.sAISCI = "A" Then
                        If imISCI < 0 Then
                            imISCI = 0
                            tmCtrls(ISCIINDEX).iChg = True
                        End If
                        ilFound = False
                    ElseIf tgSpf.sAISCI = "X" Then
                        If imISCI < 0 Then
                            imISCI = 1
                            tmCtrls(ISCIINDEX).iChg = True
                        End If
                        ilFound = False
                    End If
                    If imShortForm Then
                        If imISCI < 0 Then
                            tmCtrls(ISCIINDEX).iChg = True
                            If tgSpf.sAISCI = "Y" Then
                                imISCI = 0
                            Else
                                imISCI = 1
                            End If
                        End If
                        ilFound = False
                    End If
                    ilBox = ISCIINDEX
                End If
            Case BONUSONINVINDEX
                If imShortForm Then
                    If imRepMG < 0 Then
                        tmCtrls(REPMGINDEX).iChg = True
                        imRepMG = 0
                    End If
                    ilFound = False
                End If
                ilBox = REPMGINDEX
            Case REFIDINDEX 'L.Bianchi 06/16/2021
                If (Not imShortForm) And ((tgSpf.sCPkOrdered = "Y") Or (tgSpf.sCPkAired = "Y")) Then
                        ilBox = BONUSONINVINDEX 'L.Bianchi 05/26/2021
                    Else
                        'If (lbcInvSort.ListCount = 2) Or imShortForm Then
                        '    imChgMode = True
                        '    lbcInvSort.ListIndex = 1
                        '    imChgMode = False
                        '    ilFound = False
                        'End If
                        'ilBox = INVSORTINDEX
                        If imShortForm Then
                            If imBonusOnInv < 0 Then
                                tmCtrls(BONUSONINVINDEX).iChg = True
                                If tgSpf.sDefFillInv <> "N" Then
                                    imBonusOnInv = 0
                                Else
                                    imBonusOnInv = 1
                                End If
                            End If
                            ilFound = False
                        End If
                        ilBox = DIRECTREFIDINDEX
                End If
                ilBox = BONUSONINVINDEX
                
            Case DIRECTINDEX 'L.Bianchi 06/16/2021
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = MEGAPHONEADVID
                
            Case MEGAPHONEADVID
                ilBox = CRMIDINDEX
                
            Case CRMIDINDEX
                ilBox = DIRECTREFIDINDEX
                
            Case DIRECTREFIDINDEX 'JW - 8/2/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
                If (Not imShortForm) And ((tgSpf.sCPkOrdered = "Y") Or (tgSpf.sCPkAired = "Y")) Then
                    ilBox = REFIDINDEX
                Else
                    If imShortForm Then
                        ilFound = False
                    End If
                    ilBox = REFIDINDEX
                End If
                
            Case CADDRINDEX
                mSetShow imBoxNo
                imBoxNo = DIRECTINDEX
                If rbcBill(1).Value Then
                    rbcBill(1).SetFocus
                Else
                    rbcBill(0).SetFocus
                End If
                Exit Sub
            Case BADDRINDEX
                If edcCAddr(2).Text <> "" Then
                    ilBox = CADDRINDEX + 2
                ElseIf edcCAddr(1).Text <> "" Then
                    ilBox = CADDRINDEX + 1
                Else
                    ilBox = CADDRINDEX
                End If
            Case ADDRIDINDEX
                If edcBAddr(2).Text <> "" Then
                    ilBox = BADDRINDEX + 2
                ElseIf edcBAddr(1).Text <> "" Then
                    ilBox = BADDRINDEX + 1
                Else
                    ilBox = BADDRINDEX
                End If
            Case EDICINDEX
                If lbcLkBox.ListCount <= 2 Then
                    ilBox = PAYABLEINDEX
                Else
                    ilBox = LKBOXINDEX
                End If
                If imShortForm Then
                    ilFound = False
                End If
            Case EDIIINDEX
                If (tgSpf.sAEDIC = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDICINDEX
'            Case PRTSTYLEINDEX
            Case TERMSINDEX
                If (tgSpf.sAEDII = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDIIINDEX
            Case TAXINDEX
'                If tgSpf.sAPrtStyle = "W" Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 0
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                End If
'                If tgSpf.sAPrtStyle = "N" Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 1
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                End If
'                If imShortForm Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 1
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                End If
'                ilBox = PRTSTYLEINDEX
                ilBox = TERMSINDEX
            Case SUPPRESSNETINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = TAXINDEX
            Case Else
                If imShortForm Then
                    If (ilBox >= STATEINDEX) And (ilBox <= BONUSONINVINDEX) Then
                        ilFound = False
                    End If
                    If (ilBox >= LKBOXINDEX) And (ilBox <= TAXINDEX) Then
                        ilFound = False
                    End If
                End If
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imState <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 0
        pbcState_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imState <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 1
        pbcState_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imState = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 1
            pbcState_Paint
        ElseIf imState = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 0
            pbcState_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 1
    ElseIf imState = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 0
    End If
    pbcState_Paint
    mSetCommands
End Sub

Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    If imState = 0 Then
        pbcState.Print "Active"
    ElseIf imState = 1 Then
        pbcState.Print "Dormant"
    Else
        pbcState.Print "   "
    End If
End Sub


Private Sub pbcSuppressNet_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcSuppressNet_KeyPress(KeyAscii As Integer)
' TTP 10622 - 2023-03-08 JJB
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imSuppressNet <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSuppressNet = 1
        pbcSuppressNet_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imSuppressNet <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSuppressNet = 0
        pbcSuppressNet_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSuppressNet = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSuppressNet = 1
            pbcSuppressNet_Paint
        ElseIf imSuppressNet = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSuppressNet = 0
            pbcSuppressNet_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcSuppressNet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' TTP 10622 - 2023-03-08 JJB
    If imSuppressNet = 0 Or imSuppressNet = -1 Then
        tmCtrls(imBoxNo).iChg = True
        imSuppressNet = 1
    ElseIf imSuppressNet = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imSuppressNet = 0
    End If
    pbcSuppressNet_Paint
    mSetCommands

End Sub

Private Sub pbcSuppressNet_Paint()
' TTP 10622 - 2023-03-08 JJB
    pbcSuppressNet.Cls
    pbcSuppressNet.CurrentX = fgBoxInsetX
    pbcSuppressNet.CurrentY = 0 'fgBoxInsetY
    If imSuppressNet = 0 Then
        pbcSuppressNet.Print "No"
    ElseIf imSuppressNet = 1 Then
        pbcSuppressNet.Print "Yes"
    Else
        pbcSuppressNet.Print "   "
    End If
End Sub


Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    mCESetShow imCEBoxNo
    imCEBoxNo = -1
    mDmSetShow imDMBoxNo
    imDMBoxNo = -1
    imBillTabDir = 0
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = AGENCYINDEX Then
        If mAgencyBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = BUYERINDEX Then
        If mPersonnelBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PAYABLEINDEX Then
        If mPersonnelBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = LKBOXINDEX Then
        If mLkBoxBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = EDICINDEX Then
        If mEDIBranch(0) Then
            Exit Sub
        End If
    End If
    If imBoxNo = EDIIINDEX Then
        If mEDIBranch(1) Then
            Exit Sub
        End If
    End If
    If imBoxNo = INVSORTINDEX Then
        If mInvSortBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = TERMSINDEX Then
        If mTermsBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxNoCtrls) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Shift tab from button
                imTabDirection = -1  'Set-Right to left
                If rbcBill(1).Value Then
                    If imShortForm Then
                        ilBox = ADDRIDINDEX 'BUYERINDEX
                    Else
                        If Not imTaxDefined Then
                            imBoxNo = TAXINDEX
                            pbcSTab.SetFocus
                            Exit Sub
                        End If
                        ilBox = TAXINDEX
                    End If
                Else
                    imBoxNo = DIRECTINDEX
                    rbcBill(0).SetFocus
                    Exit Sub
                End If
            Case STATEINDEX
                If imShortForm Then
                    If imState < 0 Then
                        tmCtrls(STATEINDEX).iChg = True
                        imState = 0
                    End If
                    ilFound = False
                End If
                ilBox = PRODINDEX
            Case PRODINDEX
                If imShortForm Then
                    If Trim$(smProduct) = "" Then
                        smProduct = "[None]"
                    End If
                    ilFound = False
                End If
                ilBox = SPERSONINDEX
            Case AGENCYINDEX
                If (tgSpf.sARepCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = REPCODEINDEX
            Case REPCODEINDEX
                If (tgSpf.sAAgyCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = AGYCODEINDEX
            Case AGYCODEINDEX
                If (tgSpf.sAStnCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = STNCODEINDEX
            Case COMPINDEX
                If (tgSpf.iATargets = 0) Or imShortForm Then
                    ilFound = False
                    If (tgSpf.iATargets = 0) Or (imPriceType < 0) Then
                        imPriceType = 0
                    End If
                End If
                ilBox = DEMOINDEX
            Case DEMOINDEX
                If (tgUrf(0).sChgCrRt <> "I") Or imShortForm Then
                    ilFound = False
                    If imShortForm Then
                        If lbcCreditApproval.ListIndex < 0 Then
                            imChgMode = True
                            If (tgUrf(0).sChgCrRt <> "I") Then
                                lbcCreditApproval.ListIndex = 0   'Requires Checking
                            Else
                                lbcCreditApproval.ListIndex = 1   'Approved
                            End If
                            imChgMode = False
                            tmCtrls(CREDITAPPROVALINDEX).iChg = True
                        End If
                    End If
                End If
                ilBox = CREDITAPPROVALINDEX
            Case CREDITAPPROVALINDEX
                If (tgUrf(0).sCredit <> "I") Or imShortForm Then
                    ilFound = False
                    If lbcCreditRestr.ListIndex < 0 Then
                        imChgMode = True
                        lbcCreditRestr.ListIndex = 0
                        imChgMode = False
                        tmCtrls(CREDITRESTRINDEX).iChg = True
                    End If
                End If
                ilBox = CREDITRESTRINDEX
            Case CREDITRESTRINDEX
                If (tgUrf(0).sCredit <> "I") Or imShortForm Then
                    ilFound = False
                Else
                    If lbcCreditRestr.ListIndex <> 1 Then
                        ilFound = False
                    End If
                End If
                ilBox = CREDITRESTRINDEX + 1
            Case CREDITRESTRINDEX + 1
                If (tgUrf(0).sPayRate <> "I") Or imShortForm Then
                    ilFound = False
                    If lbcPaymRating.ListIndex < 0 Then
                        imChgMode = True
                        lbcPaymRating.ListIndex = 1
                        imChgMode = False
                        tmCtrls(PAYMRATINGINDEX).iChg = True
                    End If
                End If
                ilBox = PAYMRATINGINDEX
            Case PAYMRATINGINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = CREDITRATINGINDEX
            Case CREDITRATINGINDEX
                'If imShortForm Then
                '    If imRateOnInv < 0 Then
                '        tmCtrls(RATEONINVINDEX).iChg = True
                '        imRateOnInv = 0
                '    End If
                '    ilFound = False
                'End If
                If imShortForm Then
                    If imRepInv < 0 Then
                        tmCtrls(REPINVINDEX).iChg = True
                        'imRepInv = 0
                        If (igInternalAdfCount <> 0) And (igInternalAdfCount < UBound(tgCommAdf) / 2) Then
                            imRepInv = 1    'External
                        Else
                            imRepInv = 0    'internal
                        End If
                    End If
                    ilFound = False
                End If
                ilBox = REPINVINDEX 'INVSORTINDEX    'RATEONINVINDEX
            Case REPINVINDEX
                If (lbcInvSort.ListCount = 2) Or imShortForm Then
                    If lbcInvSort.ListIndex < 0 Then
                        imChgMode = True
                        lbcInvSort.ListIndex = 1
                        imChgMode = False
                    End If
                    ilFound = False
                End If
                ilBox = INVSORTINDEX    'RATEONINVINDEX
            Case INVSORTINDEX
                'If (tgUrf(0).iSlfCode <= 0) And (tgUrf(0).iRemoteUserID <= 0) And ((tgSpf.sCPkOrdered = "Y") Or (tgSpf.sCPkAired = "Y")) Then
                '    ilBox = PACKAGEINDEX
                'Else
                '    mSetShow imBoxNo
                '    imBoxNo = DIRECTINDEX
                '    If rbcBill(1).Value Then
                '        rbcBill(1).SetFocus
                '    Else
                '        rbcBill(0).SetFocus
                '    End If
                '    Exit Sub
                'End If
                If imShortForm Then
                    If imRateOnInv < 0 Then
                        tmCtrls(RATEONINVINDEX).iChg = True
                        imRateOnInv = 0
                    End If
                    ilFound = False
                End If
                ilBox = RATEONINVINDEX
            Case RATEONINVINDEX
                If tgSpf.sAISCI = "A" Then
                    If imISCI < 0 Then
                        imISCI = 0
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                ElseIf tgSpf.sAISCI = "X" Then
                    If imISCI < 0 Then
                        imISCI = 1
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                End If
                If imShortForm Then
                    If imISCI < 0 Then
                        tmCtrls(ISCIINDEX).iChg = True
                        If tgSpf.sAISCI = "Y" Then
                            imISCI = 0
                        Else
                            imISCI = 1
                        End If
                    End If
                    ilFound = False
                End If
                ilBox = ISCIINDEX
            Case ISCIINDEX
                'If (lbcInvSort.ListCount = 2) Or imShortForm Then
                '    imChgMode = True
                '    lbcInvSort.ListIndex = 1
                '    imChgMode = False
                '    ilFound = False
                'End If
                If (Not imShortForm) And ((tgSpf.sCPkOrdered = "Y") Or (tgSpf.sCPkAired = "Y")) Then
                    ilBox = PACKAGEINDEX
                Else
                    'mSetShow imBoxNo
                    'imBoxNo = DIRECTINDEX
                    'If rbcBill(1).Value Then
                    '    rbcBill(1).SetFocus
                    'Else
                    '    rbcBill(0).SetFocus
                    'End If
                    'Exit Sub
                    If imShortForm Then
                        If imRepMG < 0 Then
                            tmCtrls(REPMGINDEX).iChg = True
                            imRepMG = 0
                        End If
                        ilFound = False
                    End If
                    ilBox = REPMGINDEX
                End If
            Case PACKAGEINDEX
                If imShortForm Then
                    If imRepMG < 0 Then
                        tmCtrls(REPMGINDEX).iChg = True
                        imRepMG = 0
                    End If
                    ilFound = False
                End If
                ilBox = REPMGINDEX
            Case REPMGINDEX
                If imShortForm Then
                    If imBonusOnInv < 0 Then
                        tmCtrls(BONUSONINVINDEX).iChg = True
                        If tgSpf.sDefFillInv <> "N" Then
                            imBonusOnInv = 0
                        Else
                            imBonusOnInv = 1
                        End If
                    End If
                    ilFound = False
                End If
                ilBox = BONUSONINVINDEX
            Case BONUSONINVINDEX
                If imShortForm Then
                    ilFound = False
                End If
                mSetShow imBoxNo
                ilBox = REFIDINDEX
            Case REFIDINDEX
                If imShortForm Then
                    ilFound = False
                End If
                'mSetShow imBoxNo
                ilBox = DIRECTREFIDINDEX
            Case DIRECTREFIDINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ' mSetShow imBoxNo
                ilBox = MEGAPHONEADVID
                ' mSetShow imBoxNo
                ' imBoxNo = DIRECTINDEX
'                If rbcBill(1).Value Then
'                    rbcBill(1).SetFocus
'                Else
'                    rbcBill(0).SetFocus
'                End If
                
            Case DIRECTINDEX
                If rbcBill(0).Value Then
                    imBoxNo = -1
                    If (cmcUpdate.Enabled) And (igAdvtCallSource = CALLNONE) Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                Else
                    imBoxNo = CADDRINDEX
                    mEnableBox imBoxNo
                    Exit Sub
                End If
             
            Case MEGAPHONEADVID
                If imShortForm Then
                    ilFound = False
                End If
                'mSetShow imBoxNo
                ilBox = CRMIDINDEX
'            Case CRMIDINDEX
'                imBoxNo = DIRECTINDEX
'                mEnableBox imBoxNo
'                Exit Sub
                
            Case BADDRINDEX
                If edcBAddr(0).Text = "" Then
                    If edcBAddr(1).Text = "" Then
                        If edcBAddr(2).Text = "" Then
                            mSetShow imBoxNo
                            ilBox = ADDRIDINDEX 'BUYERINDEX
                            imBoxNo = ilBox
                            mEnableBox ilBox
                            Exit Sub
                        End If
                    End If
                End If
                ilBox = BADDRINDEX + 1
            Case BADDRINDEX + 1
                If edcBAddr(1).Text = "" Then
                    If edcBAddr(2).Text = "" Then
                        mSetShow imBoxNo
                        ilBox = ADDRIDINDEX 'BUYERINDEX
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
                ilBox = BADDRINDEX + 2
            Case BUYERINDEX
                If imShortForm Then
                    ilFound = False
                Else
                    ilFound = True
                End If
                ilBox = PAYABLEINDEX
            Case PAYABLEINDEX
                If imShortForm Then
                    ilFound = False
                    ilBox = LKBOXINDEX
                Else
                    If lbcLkBox.ListCount <= 2 Then
                        If tgSpf.sAEDIC = "N" Then
                            ilFound = False
                        End If
                        ilBox = EDICINDEX
                    Else
                        ilBox = LKBOXINDEX
                    End If
                End If
            Case LKBOXINDEX
                If (tgSpf.sAEDIC = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDICINDEX
            Case EDICINDEX
                If (tgSpf.sAEDII = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDIIINDEX
            Case EDIIINDEX
'                If tgSpf.sAPrtStyle = "W" Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 0
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                End If
'                If tgSpf.sAPrtStyle = "N" Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 1
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                    tmCtrls(PRTSTYLEINDEX).iChg = True
'                End If
'                If imShortForm Then
'                    If imPrtStyle < 0 Then
'                        imPrtStyle = 1
'                        tmCtrls(PRTSTYLEINDEX).iChg = True
'                    End If
'                    ilFound = False
'                End If
'                ilBox = PRTSTYLEINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = TERMSINDEX
'            Case PRTSTYLEINDEX
            Case TERMSINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = SUPPRESSNETINDEX
            Case TAXINDEX 'last control
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igAdvtCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            
            Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
                If imShortForm Then
                    ilFound = False
                End If
                
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igAdvtCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
                
            Case Else
                If imShortForm Then
                    If (ilBox >= POLITICALINDEX) And (ilBox <= BONUSONINVINDEX) Then
                        ilFound = False
                    End If
                    If (ilBox >= BUYERINDEX) And (ilBox <= TAXINDEX) Then
                        ilFound = False
                    End If
                End If
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcAdvt_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcDirect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcNotDirect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcBill_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcBill(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            plcDirect.Visible = False
            pbcDirect.Visible = False
            If (Not imShortForm) Then
                plcNotDirect.Visible = True
                pbcNotDirect.Visible = True
            End If
            tmCtrls(CADDRINDEX).iReq = False
            'tmCtrls(PRTSTYLEINDEX).iReq = False
            tmCtrls(TERMSINDEX).iReq = False
            tmCtrls(TAXINDEX).iReq = False
            'imMaxNoCtrls = PACKAGEINDEX
            'imMaxNoCtrls = REFIDINDEX 'L.Bianchi 05/26/2021
            imMaxNoCtrls = CRMIDINDEX
            
            If tmAdf.sBillAgyDir <> "A" Then
                tmCtrls(DIRECTINDEX).iChg = True
            Else
                tmCtrls(DIRECTINDEX).iChg = False
            End If
        Else
            plcDirect.Visible = True
            pbcDirect.Visible = True
            plcNotDirect.Visible = False
            pbcNotDirect.Visible = False
            'imMaxNoCtrls = TAXINDEX
            imMaxNoCtrls = SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
            pbcDirect_Paint
            If tmAdf.sBillAgyDir <> "D" Then
                tmCtrls(DIRECTINDEX).iChg = True
            Else
                tmCtrls(DIRECTINDEX).iChg = False
            End If
        End If
        mSetCommands
    End If
End Sub
Private Sub rbcBill_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = DIRECTINDEX
    imBillTabDir = 1
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case PRODINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        Case SPERSONINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSPerson, edcDropDown, imChgMode, imLbcArrowSetting
        Case AGENCYINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAgency, edcDropDown, imChgMode, imLbcArrowSetting
        Case COMPINDEX
            imLbcArrowSetting = False
            Select Case imCEBoxNo
                Case CECOMPINDEX
                    gProcessLbcClick lbcComp(0), edcCEDropDown, imChgMode, imLbcArrowSetting
                Case CECOMPINDEX + 1
                    gProcessLbcClick lbcComp(1), edcCEDropDown, imChgMode, imLbcArrowSetting
                Case CEEXCLINDEX
                    gProcessLbcClick lbcExcl(0), edcCEDropDown, imChgMode, imLbcArrowSetting
                Case CEEXCLINDEX + 1
                    gProcessLbcClick lbcExcl(1), edcCEDropDown, imChgMode, imLbcArrowSetting
            End Select
        Case DEMOINDEX
        Case INVSORTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcInvSort, edcDropDown, imChgMode, imLbcArrowSetting
        Case BUYERINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcBuyer, edcDropDown, imChgMode, imLbcArrowSetting
        Case PAYABLEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcPayable, edcDropDown, imChgMode, imLbcArrowSetting
        Case LKBOXINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcLkBox, edcDropDown, imChgMode, imLbcArrowSetting
        Case EDICINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcEDI(0), edcDropDown, imChgMode, imLbcArrowSetting
        Case EDIIINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcEDI(1), edcDropDown, imChgMode, imLbcArrowSetting
        Case TERMSINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTerms, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub plcCE_Paint()
    plcCE.CurrentX = 0
    plcCE.CurrentY = 0
    plcCE.Print "Product Protection/Program Exclusions"
End Sub
Private Sub plcDemo_Paint()
    plcDemo.CurrentX = 0
    plcDemo.CurrentY = 0
    plcDemo.Print "Demos"
End Sub

Private Sub mPopLbcName()
    Dim tlAdfExt As ADFEXT    'Advertiser extract record
    Dim ilAdf As Integer
    Dim slName As String

    lbcName.Clear
    'For ilAdf = 1 To UBound(tgCommAdf) - 1 Step 1
    For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
        tlAdfExt = tgCommAdf(ilAdf)
        If (igPopExternalAdvt = True) Or (tlAdfExt.sRepInvGen <> "E") Then
            If (tlAdfExt.sState <> "D") Or (igPopAdfAgfDormant) Then
                slName = Trim$(tlAdfExt.sName)
                gFindMatch slName, 0, lbcName
                If gLastFound(lbcName) = -1 Then
                    lbcName.AddItem slName
                    lbcName.ItemData(lbcName.NewIndex) = tlAdfExt.iCode
                End If
            End If
        End If
    Next ilAdf

End Sub

'4/3/15: Clear address fields if changed to bill agency and no receivable records exist
Private Sub mClearAddressIfRequired()
    Dim ilAdf As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    '7/9/15: added test of rbcBill
    If (tmAdf.iCode > 0) And (rbcBill(0).Value) Then
        ilAdf = gBinarySearchAdf(tmAdf.iCode)
        If ilAdf <> -1 Then
            If (tgCommAdf(ilAdf).sBillAgyDir = "D") Then
                ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Rvf.Btr", "RvfAdfCode")
                If Not ilRet Then
                    ilRet = gIICodeRefExist(Advt, tmAdf.iCode, "Phf.Btr", "PhfAdfCode")
                    If Not ilRet Then
                        For ilLoop = 0 To 2 Step 1
                            tmCtrls(CADDRINDEX + ilLoop).iChg = True
                            edcCAddr(ilLoop).Text = ""
                        Next ilLoop
                        For ilLoop = 0 To 2 Step 1
                            tmCtrls(BADDRINDEX + ilLoop).iChg = True
                            edcBAddr(ilLoop).Text = ""
                        Next ilLoop
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub mEnableSplitCue()
    Dim ilVpf As Integer
    
    cmcSplitCue.Enabled = False
    If (Asc(tgSpf.sAutoType2) And AUDIOVAULTRPS) = AUDIOVAULTRPS Then
        For ilVpf = LBound(tgVpf) To UBound(tgVpf) Step 1
            If (Asc(tgVpf(ilVpf).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                cmcSplitCue.Enabled = True
            End If
        Next ilVpf
    End If
End Sub

Private Sub mReadAxf(ilAdfCode As Integer)
    Dim ilRet As Integer
    
    tmAxfSrchKey1.iCode = tmAdf.iCode
    ilRet = btrGetEqual(hmAxf, tmAxf, imAxfRecLen, tmAxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        smAVIndicatorID = Trim$(tmAxf.sAudioVaultID)
        smXDSCue = Trim$(tmAxf.sXDSCue)
    Else
        smAVIndicatorID = ""
        smXDSCue = ""
        tmAxf.lCode = 0
    End If

End Sub

Private Function mAddOrUpdateAxf() As Integer
    Dim ilRet As Integer
    
    If tmAxf.lCode > 0 Then
        If Trim$(smAVIndicatorID) <> "" Then
            tmAxf.sAudioVaultID = smAVIndicatorID
            tmAxf.sXDSCue = smXDSCue
            ilRet = btrUpdate(hmAxf, tmAxf, imAxfRecLen)
        Else
            ilRet = btrDelete(hmAxf)
        End If
    Else
        If Trim$(smAVIndicatorID) <> "" Then
            tmAxf.lCode = 0
            tmAxf.iAdfCode = tmAdf.iCode
            tmAxf.sAudioVaultID = smAVIndicatorID
            tmAxf.sXDSCue = smXDSCue
            tmAxf.sUnused = ""
            ilRet = btrInsert(hmAxf, tmAxf, imAxfRecLen, INDEXKEY0)
        Else
            ilRet = BTRV_ERR_NONE
        End If
    End If
    mAddOrUpdateAxf = ilRet
End Function
Private Sub mPopPDFEMail()
    Dim ilRow As Integer
    Dim ilRet As Integer
    
    For ilRow = LBound(tmPdf) To UBound(tmPdf) Step 1
        tmPdf(ilRow).lCode = 0
        tmPdf(ilRow).sName = ""
        tmPdf(ilRow).sPhone = ""
        tmPdf(ilRow).sEMailAddress = ""
    Next ilRow
    If imSelectedIndex <= 0 Then
        Exit Sub
    End If
    ilRow = 0
    tmPdfSrchKey1.iCode = tmAdf.iCode
    ilRet = btrGetEqual(hmPDF, tmPdf(ilRow), imPdfRecLen, tmPdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE)
        If tmAdf.iCode <> tmPdf(ilRow).iAdfCode Then
            Exit Do
        End If
        ilRow = ilRow + 1
        ReDim Preserve tmPdf(LBound(tmPdf) To UBound(tmPdf) + 1) As PDF
        If ilRow >= 4 Then
            Exit Do
        End If
        ilRet = btrGetNext(hmPDF, tmPdf(ilRow), imPdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If (ilRow > 0) And (ilRow <> 4) Then
        tmPdf(ilRow).lCode = 0
        tmPdf(ilRow).sName = ""
        tmPdf(ilRow).sPhone = ""
        tmPdf(ilRow).sEMailAddress = ""
    End If
End Sub

Private Function mRemovePDFEMail() As Integer
    Dim ilRet As Integer
    Dim tlPdf As PDF
    
    tmPdfSrchKey1.iCode = tmAdf.iCode
    ilRet = btrGetEqual(hmPDF, tlPdf, imPdfRecLen, tmPdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE)
        ilRet = btrDelete(hmPDF)
        tmPdfSrchKey1.iCode = tmAdf.iCode
        ilRet = btrGetEqual(hmPDF, tlPdf, imPdfRecLen, tmPdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
End Function

Private Function mAddPDFEMail() As Integer
    Dim ilRet As Integer
    Dim ilPdf As Integer
    For ilPdf = 0 To UBound(tmPdf) - 1 Step 1
        tmPdf(ilPdf).lCode = 0
        tmPdf(ilPdf).iAdfCode = tmAdf.iCode
        tmPdf(ilPdf).iAgfCode = 0
        tmPdf(ilPdf).sUnused = ""
        ilRet = btrInsert(hmPDF, tmPdf(ilPdf), imPdfRecLen, INDEXKEY0)
    Next ilPdf
End Function
Private Sub mPaintAdvtTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcAdvt.ForeColor
    slFontName = pbcAdvt.FontName
    flFontSize = pbcAdvt.FontSize
    pbcAdvt.ForeColor = BLUE
    pbcAdvt.FontBold = False
    pbcAdvt.FontSize = 7
    pbcAdvt.FontName = "Arial"
    pbcAdvt.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    ''For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    For ilLoop = NAMEINDEX To CRMIDINDEX Step 1
        If ilLoop <> CREDITRESTRINDEX + 1 Then
            If ilLoop = CREDITRESTRINDEX Then
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + tmCtrls(ilLoop + 1).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            Else
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            End If
            
            'L.Bianchi 06/01/2021
            If ilLoop = REFIDINDEX Then
                'pbcAdvt.Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW - 20, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            End If
            
            If ilLoop = CRMIDINDEX Then
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW - 20, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            End If
            
            pbcAdvt.CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
            pbcAdvt.CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            Select Case ilLoop
                Case NAMEINDEX
                    pbcAdvt.Print "Name"
                Case ABBRINDEX
                    pbcAdvt.Print "Abbreviation"
                Case POLITICALINDEX
                    pbcAdvt.Print "Political"
                Case STATEINDEX
                    pbcAdvt.Print "Active/Dormant"
                Case PRODINDEX
                    pbcAdvt.Print "default Product"
                Case SPERSONINDEX
                    pbcAdvt.Print "default Salesperson"
                Case AGENCYINDEX
                    pbcAdvt.Print "default Agency"
                Case REPCODEINDEX
                    pbcAdvt.Print "G/L Adv Code"
                Case AGYCODEINDEX
                    pbcAdvt.Print "Agency Adv Code"
                Case STNCODEINDEX
                    pbcAdvt.Print "Station Adv Code"
                Case COMPINDEX
                    pbcAdvt.Print "Product Protection/Program Exclusions"
                Case DEMOINDEX
                    pbcAdvt.Print "Demos"
                Case CREDITAPPROVALINDEX
                    pbcAdvt.Print "Credit Approval"
                Case CREDITRESTRINDEX
                    pbcAdvt.Print "Credit Restriction"
                Case PAYMRATINGINDEX
                    pbcAdvt.Print "Payment Rating"
                Case CREDITRATINGINDEX
                    pbcAdvt.Print "Credit Rating"
                Case REPINVINDEX
                    pbcAdvt.Print "Rep Inv"
                Case INVSORTINDEX
                    pbcAdvt.Print "Invoice Sort"
                Case RATEONINVINDEX
                    pbcAdvt.Print "Rates"
                Case ISCIINDEX
                    pbcAdvt.Print "ISCI"
                Case REPMGINDEX
                    pbcAdvt.Print "Rep MG"
                Case BONUSONINVINDEX
                    pbcAdvt.Print "Fill"
                Case PACKAGEINDEX
                    pbcAdvt.Print "Package as"
                Case REFIDINDEX 'L.Bianchi 05/26/2021
                    pbcAdvt.Print "Ref Id"
                Case DIRECTREFIDINDEX 'JW
                    pbcAdvt.Print "Direct Ref Id"
                Case MEGAPHONEADVID 'JJB
                    pbcAdvt.Print "Megaphone Advertiser ID" 'Megaphone SWO Phase 1 2024/03/15
                Case CRMIDINDEX ' JD 09/19/2022
                    pbcAdvt.Print "CRM ID"
                    Exit For
            End Select
            If ilLoop = INVSORTINDEX Then
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 30, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                pbcAdvt.CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
                pbcAdvt.CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcAdvt.Print "Allow on"
                pbcAdvt.CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
                pbcAdvt.CurrentY = tmCtrls(ilLoop).fBoxY + tmCtrls(ilLoop).fBoxH / 2 - 30
                pbcAdvt.Print "Invoices:"
            End If
            If ilLoop = BKOUTPOOLINDEX Then
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop).fBoxW - 15, tmCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                pbcAdvt.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                pbcAdvt.CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
                pbcAdvt.CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcAdvt.Print "Pool"
            End If
        End If
    Next ilLoop
    pbcAdvt.FontSize = flFontSize
    pbcAdvt.FontName = slFontName
    pbcAdvt.FontSize = flFontSize
    pbcAdvt.ForeColor = llColor
    pbcAdvt.FontBold = True


End Sub

Private Sub mPaintDirectTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcDirect.ForeColor
    slFontName = pbcDirect.FontName
    flFontSize = pbcDirect.FontSize
    pbcDirect.ForeColor = BLUE
    pbcDirect.FontBold = False
    pbcDirect.FontSize = 7
    pbcDirect.FontName = "Arial"
    pbcDirect.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcDirect.Line (tmCtrls(CADDRINDEX).fBoxX - 15, tmCtrls(CADDRINDEX).fBoxY - 15)-Step(tmCtrls(CADDRINDEX).fBoxW + 15, tmCtrls(ADDRIDINDEX).fBoxY - 30), BLUE, B
    pbcDirect.CurrentX = tmCtrls(CADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDirect.CurrentY = tmCtrls(CADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcDirect.Print "Contract Address"
    pbcDirect.Line (tmCtrls(BADDRINDEX).fBoxX - 15, tmCtrls(BADDRINDEX).fBoxY - 15)-Step(tmCtrls(BADDRINDEX).fBoxW + 15, tmCtrls(ADDRIDINDEX).fBoxY - 30), BLUE, B
    pbcDirect.CurrentX = tmCtrls(BADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDirect.CurrentY = tmCtrls(BADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcDirect.Print "Billing Address"
    For ilLoop = ADDRIDINDEX To TAXINDEX Step 1
        pbcDirect.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        pbcDirect.CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcDirect.CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case CADDRINDEX
                pbcDirect.Print "Contract Address"
            Case BADDRINDEX
                pbcDirect.Print "Billing Address"
            Case ADDRIDINDEX
                pbcDirect.Print "Address ID"
            Case BUYERINDEX
                pbcDirect.Print "default Buyer/Phone # and Extension/Fax #"
            Case PAYABLEINDEX
                pbcDirect.Print "Payables Contact/Phone # and Extension/Fax #"
            Case LKBOXINDEX
                pbcDirect.Print "Lock Box Name"
            Case EDICINDEX
                pbcDirect.Print "EDI for Contracts"
            Case EDIIINDEX
                pbcDirect.Print "EDI for Invoices"
            Case TERMSINDEX
                pbcDirect.Print "Terms"
            Case TAXINDEX
                pbcDirect.Print "Commercial Tax"
        End Select
    Next ilLoop

    For ilLoop = SUPPRESSNETINDEX To UNUSEDINDEX Step 1 ' TTP 10622 - 2023-03-08 JJB
        Select Case ilLoop
            Case SUPPRESSNETINDEX
                pbcDirect.Line (tmCtrls(SUPPRESSNETINDEX).fBoxX - 15, tmCtrls(SUPPRESSNETINDEX).fBoxY - 15)-Step(tmCtrls(SUPPRESSNETINDEX).fBoxW / 2, tmCtrls(SUPPRESSNETINDEX).fBoxH + 15), BLUE, B
                pbcDirect.CurrentX = tmCtrls(SUPPRESSNETINDEX).fBoxX + 15
                pbcDirect.CurrentY = tmCtrls(SUPPRESSNETINDEX).fBoxY - 15
                pbcDirect.Print "Suppress Net Amount for Trade Invoices"
            Case UNUSEDINDEX
                pbcDirect.Line (2760, tmCtrls(SUPPRESSNETINDEX).fBoxY - 15)-Step(9325, tmCtrls(SUPPRESSNETINDEX).fBoxH + 15), BLUE, B
                pbcDirect.CurrentX = 2760 'tmCtrls(UNUSEDINDEX).fBoxX + 15
                pbcDirect.CurrentX = 3700 'tmCtrls(UNUSEDINDEX).fBoxX + 15
                pbcDirect.CurrentY = 1750 'tmCtrls(UNUSEDINDEX).fBoxY - 15
                pbcDirect.Print ""
        End Select
    Next ilLoop

    For ilLoop = PCT90INDEX To LSTTOPAYINDEX Step 1
        pbcDirect.Line (tmDTCtrls(ilLoop).fBoxX, tmDTCtrls(ilLoop).fBoxY)-Step(tmDTCtrls(ilLoop).fBoxW - 15, tmDTCtrls(ilLoop).fBoxH - 10), LIGHTERYELLOW, BF
        If ilLoop = TOTALGROSSINDEX Then
            pbcDirect.Line (tmDTCtrls(ilLoop).fBoxX - 15, tmDTCtrls(ilLoop).fBoxY - 15)-Step(tmDTCtrls(ilLoop).fBoxW + tmDTCtrls(ilLoop + 1).fBoxW + 30, tmDTCtrls(ilLoop).fBoxH + 15), BLUE, B
        ElseIf ilLoop = DATEENTRDINDEX Then
            'Box part of TOTALGROSSINDEX
        Else
            pbcDirect.Line (tmDTCtrls(ilLoop).fBoxX - 15, tmDTCtrls(ilLoop).fBoxY - 15)-Step(tmDTCtrls(ilLoop).fBoxW + 15, tmDTCtrls(ilLoop).fBoxH + 15), BLUE, B
        End If
        pbcDirect.CurrentX = tmDTCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcDirect.CurrentY = tmDTCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case PCT90INDEX
                pbcDirect.Print "% Over 90"
            Case CURRARINDEX
                pbcDirect.Print "Current A/R Total"
            Case UNBILLEDINDEX
                pbcDirect.Print "Unbilled + Projected"
            Case HICREDITINDEX
                pbcDirect.Print "Highest A/R Total"
            Case TOTALGROSSINDEX
                pbcDirect.Print "Total Gross"
            Case DATEENTRDINDEX
                pbcDirect.Print "since"
            Case NSFCHKSINDEX
                pbcDirect.Print "# NSF Checks"
            Case DATELSTINVINDEX
                pbcDirect.Print "Date Last Billed"
            Case DATELSTPAYMINDEX
                pbcDirect.Print "Date Last Payment"
            Case AVGTOPAYINDEX
                pbcDirect.Print "Avg # Days to Pay"
            Case LSTTOPAYINDEX
                pbcDirect.Print "# Days to Pay Last Payment"
        End Select
    Next ilLoop
    
    pbcDirect.FontSize = flFontSize
    pbcDirect.FontName = slFontName
    pbcDirect.FontSize = flFontSize
    pbcDirect.ForeColor = llColor
    pbcDirect.FontBold = True
End Sub


Private Sub mPaintNotDirectTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcNotDirect.ForeColor
    slFontName = pbcNotDirect.FontName
    flFontSize = pbcNotDirect.FontSize
    pbcNotDirect.ForeColor = BLUE
    pbcNotDirect.FontBold = False
    pbcNotDirect.FontSize = 7
    pbcNotDirect.FontName = "Arial"
    pbcNotDirect.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilLoop = PCT90INDEX To LSTTOPAYINDEX Step 1
        pbcNotDirect.Line (tmNDTCtrls(ilLoop).fBoxX, tmNDTCtrls(ilLoop).fBoxY)-Step(tmNDTCtrls(ilLoop).fBoxW - 15, tmNDTCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
        If ilLoop = TOTALGROSSINDEX Then
            pbcNotDirect.Line (tmNDTCtrls(ilLoop).fBoxX - 15, tmNDTCtrls(ilLoop).fBoxY - 15)-Step(tmNDTCtrls(ilLoop).fBoxW + tmNDTCtrls(ilLoop + 1).fBoxW + 30, tmNDTCtrls(ilLoop).fBoxH + 15), BLUE, B
        ElseIf ilLoop = DATEENTRDINDEX Then
        Else
            pbcNotDirect.Line (tmNDTCtrls(ilLoop).fBoxX - 15, tmNDTCtrls(ilLoop).fBoxY - 15)-Step(tmNDTCtrls(ilLoop).fBoxW + 15, tmNDTCtrls(ilLoop).fBoxH + 15), BLUE, B
        End If
        pbcNotDirect.CurrentX = tmNDTCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcNotDirect.CurrentY = tmNDTCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case PCT90INDEX
                pbcNotDirect.Print "% Over 90"
            Case CURRARINDEX
                pbcNotDirect.Print "Current A/R Total"
            Case UNBILLEDINDEX
                pbcNotDirect.Print "Unbilled + Projected"
            Case HICREDITINDEX
                pbcNotDirect.Print "Highest A/R Total"
            Case TOTALGROSSINDEX
                pbcNotDirect.Print "Total Gross"
            Case DATEENTRDINDEX
                pbcNotDirect.Print "since"
            Case NSFCHKSINDEX
                pbcNotDirect.Print "# NSF Checks"
            Case DATELSTINVINDEX
                pbcNotDirect.Print "Date Last Billed"
            Case DATELSTPAYMINDEX
                pbcNotDirect.Print "Date Last Payment"
            Case AVGTOPAYINDEX
                pbcNotDirect.Print "Avg # Days to Pay"
            Case LSTTOPAYINDEX
                pbcNotDirect.Print "# Days to Pay Last Payment"
        End Select
    Next ilLoop
    
    pbcNotDirect.FontSize = flFontSize
    pbcNotDirect.FontName = slFontName
    pbcNotDirect.FontSize = flFontSize
    pbcNotDirect.ForeColor = llColor
    pbcNotDirect.FontBold = True
End Sub

Private Sub edcRefId_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRefId_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
