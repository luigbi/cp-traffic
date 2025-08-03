VERSION 5.00
Begin VB.Form Agency 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11160
   ClientLeft      =   17595
   ClientTop       =   2520
   ClientWidth     =   15780
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
   ScaleHeight     =   11160
   ScaleWidth      =   15780
   Begin VB.TextBox edcSuppressNet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   9405
      MaxLength       =   6
      TabIndex        =   58
      Top             =   4500
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox pbcSuppressNet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9360
      ScaleHeight     =   210
      ScaleWidth      =   1350
      TabIndex        =   37
      Top             =   4200
      Width           =   1350
   End
   Begin VB.TextBox edcCRMID 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   8625
      TabIndex        =   30
      Top             =   8655
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox pbcXMLDates 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3090
      ScaleHeight     =   210
      ScaleWidth      =   450
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8685
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox edcXMLCall 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   8535
      MaxLength       =   6
      TabIndex        =   40
      Top             =   8085
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox pbcDigitRating 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12810
      ScaleHeight     =   210
      ScaleWidth      =   990
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1215
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox edcXMLBand 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   8445
      MaxLength       =   6
      TabIndex        =   41
      Top             =   7695
      Visible         =   0   'False
      Width           =   585
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
      Left            =   12510
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcExport 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1125
      ScaleHeight     =   210
      ScaleWidth      =   1350
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7710
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   30
      Width           =   3675
   End
   Begin VB.PictureBox pbcPackage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   585
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6780
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcRating 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   13605
      MaxLength       =   7
      TabIndex        =   21
      Top             =   1785
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5685
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   9315
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   5730
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
      Left            =   9300
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   5730
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
      Left            =   16365
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11925
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6015
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   30
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4530
      Width           =   30
   End
   Begin VB.ListBox lbcCity 
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
      Left            =   13950
      Sorted          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   9195
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   795
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
      Index           =   1
      Left            =   12120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3540
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcISCI 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   840
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7275
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   16035
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   645
      Top             =   4185
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
      Left            =   15270
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   1845
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
      Left            =   9285
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3540
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
      Left            =   15300
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3810
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
      Left            =   12105
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcCreditLimit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   12525
      MaxLength       =   10
      TabIndex        =   19
      Top             =   8370
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   9300
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1800
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
      Left            =   9225
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   10440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8535
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcRefId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5820
      MaxLength       =   36
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5715
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
      Left            =   9615
      Picture         =   "Agency.frx":0000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   9315
      MaxLength       =   40
      TabIndex        =   28
      Top             =   2655
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
      Left            =   9315
      MaxLength       =   40
      TabIndex        =   27
      Top             =   2400
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
      Left            =   14355
      MaxLength       =   40
      TabIndex        =   32
      Top             =   2655
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
      Left            =   14355
      MaxLength       =   40
      TabIndex        =   31
      Top             =   2400
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
      Left            =   14355
      MaxLength       =   40
      TabIndex        =   29
      Top             =   2145
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcStnCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   15975
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1215
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox edcRepCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   14085
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   9300
      MaxLength       =   40
      TabIndex        =   26
      Top             =   2145
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcAbbr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   12270
      MaxLength       =   5
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox plcScreen 
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
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1455
   End
   Begin VB.TextBox edcComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   12510
      MaxLength       =   6
      TabIndex        =   10
      Top             =   6540
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
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
      Left            =   7395
      TabIndex        =   52
      Top             =   4800
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
      Left            =   6270
      TabIndex        =   51
      Top             =   4800
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
      Left            =   5130
      TabIndex        =   50
      Top             =   4800
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
      Left            =   4005
      TabIndex        =   49
      Top             =   4800
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
      Left            =   2880
      TabIndex        =   48
      Top             =   4800
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
      Left            =   1755
      TabIndex        =   47
      Top             =   4800
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
      Left            =   630
      TabIndex        =   46
      Top             =   4800
      Width           =   1050
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   5310
      Width           =   240
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   165
      ScaleHeight     =   270
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   4890
      Width           =   435
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3465
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5265
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4050
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5310
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcAgy 
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
      Height          =   1440
      Index           =   1
      Left            =   840
      Picture         =   "Agency.frx":00FA
      ScaleHeight     =   1440
      ScaleWidth      =   8490
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9015
      Width           =   8490
   End
   Begin VB.PictureBox pbcAgy 
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
      Height          =   3885
      Index           =   0
      Left            =   345
      Picture         =   "Agency.frx":D454
      ScaleHeight     =   3885
      ScaleWidth      =   8490
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   8490
   End
   Begin VB.PictureBox plcBkgd 
      ForeColor       =   &H00000000&
      Height          =   3660
      Left            =   270
      ScaleHeight     =   3600
      ScaleWidth      =   8580
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   8640
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
      Left            =   10065
      TabIndex        =   57
      Top             =   360
      Width           =   795
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   4260
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Agency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Agency.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Agency.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Agency input screen code
Option Explicit
Option Compare Text

'Agency Field Areas
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 38)  As FIELDAREA
Dim tmARCtrls(0 To 11)  As FIELDAREA

Dim imLBCtrls As Integer
Dim imExport As Integer             '0=CSI Form; 1=Agency (OMD) Form; 2=XML
Dim imXMLDates As Integer           '0=M-Su; 1= Aired
Dim imPrtStyle As Integer           '0=Wide; 1=Narrow
Dim imDigitRating As Integer        '1=one digit rating; 2=two digit rating
Dim imISCI As Integer               '0=Yes; 1= No
Dim imPackage As Integer            '0=Daypart; 1= Time
Dim imState As Integer              '0=Active; 1= Dormant
Dim imBoxNo As Integer              'Current Agency Box
Dim tmEDICode() As SORTCODE
Dim smEDICodeTag As String
Dim tmInvSortCode() As SORTCODE
Dim smInvSortCodeTag As String
Dim tmLkBoxCode() As SORTCODE
Dim smLkBoxCodeTag As String
Dim tmTermsCode() As SORTCODE
Dim smTermsCodeTag As String
Dim imSuppressNet As Integer        '0=No (default), 1=Yes  Suppress Net Amount for Trade Invoices  ' TTP 10622 - 2023-03-08 JJB

Dim tmPayableCode() As SORTCODE
Dim smPayableCodeTag As String

Dim tmBuyerCode() As SORTCODE
Dim smBuyerCodeTag As String

Dim tmTaxSortCode() As SORTCODE
Dim smTaxSortCodeTag As String

Dim hmAgf As Integer                'Agency file handle
Dim tmAgf As AGF                    'AGF record image
Dim tmAgfSrchKey As INTKEY0         'AGF key record image
Dim imAgfRecLen As Integer          'AGF record length
Dim hmPnf As Integer                'Product file handle
Dim tmBPnf As PNF
Dim tmPPnf As PNF
Dim imPnfRecLen As Integer
Dim tmPnfSrchKey As INTKEY0
Dim tmPnfSrchKey1 As PNFKEY1
Dim imNewPnfCode() As Integer
Dim hmPDF As Integer                'Media code file handle
Dim imPdfRecLen As Integer          'McF record length
Dim tmPdfSrchKey1 As INTKEY0
Dim tmPdfSrchKey2 As INTKEY0
Dim tmPdf() As PDF
Dim bmPDFEMailChgd As Boolean
Dim hmSaf As Integer
Dim hmAdf As Integer
Dim tmAgfx As AGFX                  'AGFX record Image 'L.Bianchi 05/25/2021
'Dim hmDsf As Integer               'Delete Stamp file handle
'Dim tmDsf As DSF                   'DSF record image
'Dim tmDsfSrchKey As LONGKEY0       'DSF key record image
'Dim imDsfRecLen As Integer         'DSF record length
'Dim tmRec As LPOPREC
Dim imPbcIndex As Integer
Dim imChgMode As Integer            'Change mode status (so change not entered when in change)
Dim imEdcChgMode As Integer         'Change mode status (so change not entered when in change)
Dim imBSMode As Integer             'Backspace flag
Dim imSelectedIndex As Integer      'Index of selected record (0 if new)
Dim imTerminate As Integer          'True = terminating task, False= OK
Dim smName As String                'Agency name
Dim smCity As String                'Agency City ID
Dim smBuyer As String               'Buyer name, saved to determine if changed
Dim smPayable As String             'Payable name, saved to determine if changed
Dim smOrigBuyer As String           'Buyer name, saved to determine if changed
Dim smOrigPayable As String         'Payable name, saved to determine if changed
Dim smSPerson As String             'Salesperson name, saved to determine if changed
Dim smLkBox As String               'Lock box, saved to determine if changed
Dim smCRMID As String               'saved to determine if changed
Dim smInvSort As String
Dim smTerms As String
Dim smTax As String
Dim smEDIC As String                'EDI service for contracts, saved to determine if changed
Dim smEDII As String                'EDI service for invoices, saved to determine if changed
Dim imUpdateAllowed As Integer      'User can update records
Dim imComboBoxIndex As Integer
Dim imCreditRestrFirst As Integer   'First time at field-set default if required
Dim imPaymRatingFirst As Integer    'First time at field-set default if required
'Dim imPrtStyleFirst As Integer     'First time at field-set default if required
Dim imScriptFirst As Integer        'First time at field-set default if required
Dim imCombo As Integer              'True=Combo salesperson; False=Standard allow salesperson
Dim imTaxFirst As Integer           'First time at field-set default if required
Dim imCommFirst As Integer          'First time at field-set default if required
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
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer       '0=left to right (Tab); -1=right to left (Shift tab)
Dim imTaxDefined As Integer
Dim imPopReqd As Integer            'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imShortForm As Integer          'True=Only process min number of fields
Dim imChgSaveFlag As Integer        'Indicates if any changed saved

Dim fmAdjFactorW As Single          'Width adjustment factor
Dim fmAdjFactorH As Single          'Width adjustment factor

Const NAMEINDEX = 1                 'Name control/field
Const ABBRINDEX = 2                 'Last name control/field
Const CITYINDEX = 3
Const STATEINDEX = 4                'Active/Dormant control/index
Const COMMINDEX = 5                 'Commission control/index
Const SPERSONINDEX = 6              'Salesperson control/field
Const DIGITRATINGINDEX = 7          '1 or 2 Digit Rating
Const REPCODEINDEX = 8              'Rep Agency Code control/field
Const STNCODEINDEX = 9              'Station Agency Code control/field
Const CREDITAPPROVALINDEX = 10
Const CREDITRESTRINDEX = 11         'Credit restriction control/field
Const PAYMRATINGINDEX = 13          'Payment rating control/field
Const CREDITRATINGINDEX = 14
Const ISCIINDEX = 15                'ISCI control/field
Const INVSORTINDEX = 16             'Invoice sort control/field
Const PACKAGEINDEX = 17
Const CADDRINDEX = 18               'Contract address control/field (3 fields)
Const BADDRINDEX = 21               'Billing address control/field (3 fields)
Const BUYERINDEX = 24               'Buyer control/field
Const PAYABLEINDEX = 25             'Buyer control/field
Const CRMIDINDEX = 26               'CRM ID
Const LKBOXINDEX = 27               'Lock box control/field
Const REFIDINDEX = 28               'RefId Contral Field
Const EDICINDEX = 29                'EDI for Contracts control/field
Const EDIIINDEX = 30                'EDI for Invoices control/field
Const TERMSINDEX = 31
Const TAXINDEX = 32                 'Tax 1 and 2 control/fieldConst EXPORTFORMINDEX = 29    'Contract Export control/field
Const EXPORTFORMINDEX = 33
Const XMLCALLINDEX = 34
Const XMLBANDINDEX = 35
Const XMLDATESINDEX = 36
Const SUPPRESSNETINDEX = 37         'Suppress Net Amount for Trade Invoices ' TTP 10622 - 2023-03-08 JJB
Const UNUSEDINDEX = 38              'Filler (Unused)                        ' TTP 10622 - 2023-03-08 JJB

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
    
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
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
            lacCode.Caption = str$(tmAgf.iCode)
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
    'Remove personnel incase user press undo, then changed agencies
    For ilLoop = LBound(imNewPnfCode) To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    
    ReDim imNewPnfCode(0 To 0) As Integer
    pbcAgy(imPbcIndex).Cls
    mPaintAgyTitle imPbcIndex
    
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            smName = slStr
            mSetChg NAMEINDEX   'altered flag set so field is saved
        End If
    End If
    
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    
    pbcAgy_Paint imPbcIndex
    Screen.MousePointer = vbDefault  'Default
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
    
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igAgyCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgAgyName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgAgyName    'New name
            End If
            cbcSelect_Change
            If sgAgyName <> "" Then
                mSetCommands
                gFindMatch sgAgyName, 1, cbcSelect
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
    
    For ilLoop = LBound(imNewPnfCode) To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    If igAgyCallSource <> CALLNONE Then
        igAgyCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
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
    If igAgyCallSource <> CALLNONE Then
        sgAgyName = smName & ", " & smCity 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgAgyName = "[New]"
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
    
    'If not saved- clear any personnel added
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = LBound(imNewPnfCode) To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmPnf)
        End If
    Next ilLoop
    
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    
    If imChgSaveFlag Then
        sgAgencyTag = ""
        mPopulate
    End If
    
    If igAgyCallSource <> CALLNONE Then
        If sgAgyName = "[New]" Then
            igAgyCallSource = CALLCANCELLED
        Else
            igAgyCallSource = CALLDONE
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
    
    mSetShow imBoxNo
    imBoxNo = -1
    
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case NAMEINDEX
            lbcName.Visible = Not lbcName.Visible
        Case CITYINDEX
            lbcCity.Visible = Not lbcCity.Visible
        Case SPERSONINDEX
            lbcSPerson.Visible = Not lbcSPerson.Visible
        Case CREDITAPPROVALINDEX
            lbcCreditApproval.Visible = Not lbcCreditApproval.Visible
        Case CREDITRESTRINDEX
            lbcCreditRestr.Visible = Not lbcCreditRestr.Visible
        Case PAYMRATINGINDEX
            lbcPaymRating.Visible = Not lbcPaymRating.Visible
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
    
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
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
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Adf.Btr", "ADFAGFCODE") 'adfagfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Advertiser references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Cdf.Btr", "CDFAGFCODE")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Comment references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Chf.Btr", "CHFAGFCODE") 'chfagfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Lst.Mkd", "LstAgfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate Log Spot references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Phf.Btr", "PhfAgfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables History references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Agency, tmAgf.iCode, "Rvf.Btr", "RvfAgfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID), vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE, False) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        gGetSyncDateTime slSyncDate, slSyncTime
        slStamp = gFileDateTime(sgDBPath & "Agf.Btr")
        tmPnfSrchKey1.iAgfCode = tmAgf.iCode
        ilRet = btrGetGreaterOrEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Do While (ilRet = BTRV_ERR_NONE) And (tlPnf.iAdfCode = tmAgf.iCode)
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
            tmPnfSrchKey1.iAgfCode = tmAgf.iCode
            ilRet = btrGetGreaterOrEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Loop
        
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
        
        ReDim imNewPnfCode(0 To 0) As Integer
        ilRet = btrDelete(hmAgf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Agency
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "AGF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmAgf.iRemoteID
'            tmDsf.lAutoCode = tmAgf.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'        End If
'        If ilRet <> BTRV_ERR_NONE Then
'            Screen.MousePointer = vbDefault
'            ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
'            Exit Sub
'        End If

        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        
        'If Traffic!lbcAgency.Tag <> "" Then
        '    If slStamp = Traffic!lbcAgency.Tag Then
        '        Traffic!lbcAgency.Tag = FileDateTime(sgDBPath & "Agf.Btr")
        '    End If
        'End If
        
        If sgAgencyTag <> "" Then
            If slStamp = sgAgencyTag Then
                sgAgencyTag = gFileDateTime(sgDBPath & "Agf.Btr")
            End If
        End If
        
        'Traffic!lbcAgency.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgAgency()
        cbcSelect.RemoveItem imSelectedIndex
        'Remove from tgCommAdf
        
        For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
            If tgCommAgf(ilLoop).iCode = tmAgf.iCode Then
                For ilIndex = ilLoop To UBound(tgCommAgf) - 1 Step 1
                    tgCommAgf(ilIndex) = tgCommAgf(ilIndex + 1)
                Next ilIndex
                ReDim Preserve tgCommAgf(LBound(tgCommAgf) To UBound(tgCommAgf) - 1) As AGFEXT
                Exit For
            End If
        Next ilLoop
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
        For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
            tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
            ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmPnf)
            End If
        Next ilLoop
        ReDim imNewPnfCode(0 To 0) As Integer
        Screen.MousePointer = vbDefault
    End If
    
     'L.Bianchi '05/31/2021' start
    tmAgfx.iCode = tmAgf.iCode
    SQLQuery = "select * from AGFX_Agencies WHERE agfxCode =" & tmAgfx.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        SQLQuery = "DELETE FROM AGFX_Agencies WHERE agfxCode = " & tmAgfx.iCode
         If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "Agency-cmcErase"
         End If
    End If
    
    tmAgfx.iCode = 0
    tmAgfx.sRefId = ""
    'L.Bianchi '05/31/2021' End
    
    'Remove focus from control and make invisible
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcAgy(imPbcIndex).Cls
    mPaintAgyTitle imPbcIndex
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
    gCtrlGotFocus cmcErase
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
    
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbYesNo + vbQuestion, "Merge Agency")
    If ilRet = vbNo Then
        Exit Sub
    End If
    
    ilRet = MsgBox("Are all other users off the traffic system", vbYesNo + vbQuestion, "Merge Agency")
    If ilRet = vbNo Then
        Exit Sub
    End If
    
    igMergeCallSource = AGENCIESLIST
    Merge.Show vbModal
    Screen.MousePointer = vbHourglass
    pbcAgy(imPbcIndex).Cls
    mPaintAgyTitle imPbcIndex
    cbcSelect.Clear
    mPopulate
    cbcSelect.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = AGENCIESLIST
    igRptType = 0
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Agency^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    ''Screen.MousePointer = vbDefault    'Default
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub

Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
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
        pbcAgy(imPbcIndex).Cls
        pbcAgy_Paint imPbcIndex
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    edcCRMID.Text = ""
    pbcAgy(imPbcIndex).Cls
    mPaintAgyTitle imPbcIndex
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
    gCtrlGotFocus cmcUndo
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
    
    slName = Trim$(smName) & ", " & smCity   'Save name
    imSvSelectedIndex = imSelectedIndex
    
    If mSaveRecChg(False) = False Then
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
    ilCode = tmAgf.iCode
    cbcSelect.Clear
    sgAgencyTag = ""
    mPopulate
    
    For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tgAgency(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
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
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
End Sub

Private Sub edcCRMID_Change()
    ' tmCtrls(CRMIDINDEX).iChg = True
    mSetChg CRMIDINDEX
End Sub

Private Sub edcCRMID_LostFocus()
    If Len(edcCRMID.Text) > 0 And IsNumeric(edcCRMID.Text) Then
        If Val(edcCRMID.Text) < 2147483647 Then
            tmAgfx.lCrmId = CLng(edcCRMID.Text)
        End If
    Else
        edcCRMID.Text = ""
    End If
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
Private Sub edcAbbr_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcAbbr_GotFocus()
    If edcAbbr.Text = "" Then
        edcAbbr.Text = Left$(smName, 5)
        mSetChg imBoxNo   'Change event not generated
        mSetCommands
    End If
    gCtrlGotFocus edcAbbr
End Sub
Private Sub edcAbbr_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcBAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcBAddr_GotFocus(Index As Integer)
    gCtrlGotFocus edcBAddr(Index)
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
    gCtrlGotFocus edcCAddr(Index)
End Sub
Private Sub edcComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcComm_GotFocus()
    If (edcComm.Text = "") And imCommFirst Then
        edcComm.Text = 15
        mSetChg imBoxNo   'Change event not generated
        mSetCommands
    End If
    imCommFirst = False
    gCtrlGotFocus edcComm
End Sub
Private Sub edcComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcComm.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcComm.Text
    slStr = Left$(slStr, edcComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcComm.SelStart - edcComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
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
    gCtrlGotFocus edcCreditLimit
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
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case NAMEINDEX
            If Not imEdcChgMode Then
                imEdcChgMode = True
                smName = edcDropDown.Text
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcDropDown, lbcName, imBSMode, slStr)
                If ilRet = 1 Then   'input was ""
                    lbcName.ListIndex = -1
                    smName = ""
                ElseIf ilRet = 0 Then
                    smName = edcDropDown.Text   'lbcName.List(lbcName.ListIndex)
                End If
                imEdcChgMode = False
            End If
        Case CITYINDEX
            If Not imEdcChgMode Then
                imEdcChgMode = True
                smCity = edcDropDown.Text
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcDropDown, lbcCity, imBSMode, slStr)
                If ilRet = 1 Then   'input was ""
                    lbcCity.ListIndex = -1
                ElseIf ilRet = 0 Then
                    smCity = edcDropDown.Text
                End If
                imEdcChgMode = False
            End If
        Case SPERSONINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSPerson, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSPerson.ListIndex = 1
            End If
        Case CREDITAPPROVALINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCreditApproval, imBSMode, imComboBoxIndex
        Case CREDITRESTRINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCreditRestr, imBSMode, imComboBoxIndex
        Case PAYMRATINGINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcPaymRating, imBSMode, imComboBoxIndex
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
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case NAMEINDEX
        Case CITYINDEX
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
        Case CREDITAPPROVALINDEX
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
        Case INVSORTINDEX
'            If lbcInvSort.ListCount = 1 Then
'                lbcInvSort.ListIndex = 0
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
            Case NAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcName, imLbcArrowSetting
            Case CITYINDEX
                gProcessArrowKey Shift, KeyCode, lbcCity, imLbcArrowSetting
            Case SPERSONINDEX
                gProcessArrowKey Shift, KeyCode, lbcSPerson, imLbcArrowSetting
            Case CREDITAPPROVALINDEX
                gProcessArrowKey Shift, KeyCode, lbcCreditApproval, imLbcArrowSetting
            Case CREDITRESTRINDEX
                gProcessArrowKey Shift, KeyCode, lbcCreditRestr, imLbcArrowSetting
            Case PAYMRATINGINDEX
                gProcessArrowKey Shift, KeyCode, lbcPaymRating, imLbcArrowSetting
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
    If (imBoxNo <> INVSORTINDEX) And (imBoxNo <> EDICINDEX) And (imBoxNo <> EDIIINDEX) And (imBoxNo <> TERMSINDEX) And (imBoxNo <> EDICINDEX) And (imBoxNo <> EDIIINDEX) And (imBoxNo <> LKBOXINDEX) And (imBoxNo <> BUYERINDEX) And (imBoxNo <> SPERSONINDEX) And (imBoxNo <> PAYABLEINDEX) Then
        edcDropDown.Text = gRemoveIllegalPastedChar(edcDropDown.Text)
    End If
    Select Case imBoxNo
        Case NAMEINDEX
            'smName = edcDropDown.Text
        Case CITYINDEX
            'smCity = edcDropDown.Text
        Case BUYERINDEX
            smBuyer = edcDropDown.Text
        Case PAYABLEINDEX
            smPayable = edcDropDown.Text
    End Select
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case SPERSONINDEX, BUYERINDEX, PAYABLEINDEX, LKBOXINDEX, INVSORTINDEX, EDICINDEX, EDIIINDEX, TERMSINDEX
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
    gCtrlGotFocus edcRepCode
End Sub
Private Sub edcStnCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcStnCode_GotFocus()
    gCtrlGotFocus edcStnCode
End Sub



Private Sub edcXMLBand_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcXMLBand_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcXMLCall_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcXMLCall_GotFocus()
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


Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(AGENCIESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcAgy(imPbcIndex).Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcAgy(imPbcIndex).Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
    Me.KeyPreview = True
    Agency.Refresh
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
        If imBoxNo > 0 Then
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
           ' Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
            Me.Width = 12650
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
    Erase tmPayableCode
    Erase tmTaxSortCode
    Erase tmBuyerCode
    Erase tmPdf

    ilRet = btrClose(hmPDF)
    btrDestroy hmPDF
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    
    Set Agency = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
Private Sub lbcCity_Click()
    gProcessLbcClick lbcCity, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCity_GotFocus()
    gCtrlGotFocus ActiveControl
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
    gProcessLbcClick lbcName, edcDropDown, imChgMode, imLbcArrowSetting
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
'*      Procedure Name:mBuyerPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Buyer Personnel       *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mBuyerPop(ilAgyCode As Integer, slRetainName As String, ilReturnCode As Integer)
'
'   mBuyerPop
'   Where:
'       ilAgyCode (I)- Agency code value
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
            'ilRet = gPopPersonnelBox(Agency, 1, ilAgyCode, "B", True, 2, lbcBuyer, lbcBuyerCode)
            ilRet = gPopPersonnelBox(Agency, 1, ilAgyCode, "B", True, 2, lbcBuyer, tmBuyerCode(), smBuyerCodeTag)
        'End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mBuyerPopErr
            gCPErrorMsg ilRet, "mBuyerPop (gPopPersonnelBox)", Agency
            On Error GoTo 0
            'Filter out any contact not associated with this agency
            If imSelectedIndex = 0 Then
                For ilLoop = UBound(tmBuyerCode) - 1 To 0 Step -1 'lbcBuyerCode.ListCount - 1 To 0 Step -1
                    ilFound = False
                    slNameCode = tmBuyerCode(ilLoop).sKey  'lbcBuyerCode.List(ilLoop)
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
                        lbcBuyer.RemoveItem ilLoop
                        'lbcBuyerCode.RemoveItem ilLoop
                        gRemoveItemFromSortCode ilLoop, tmBuyerCode()
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
'
    Dim ilLoop As Integer
    tmAgf.iCode = 0         'Required by pop buyer and contact
    lbcName.ListIndex = -1
    lbcCity.ListIndex = -1
    smName = ""
    smCity = ""
    edcAbbr.Text = ""
    imState = -1
    edcComm.Text = ""
    lbcSPerson.ListIndex = -1
    edcRepCode.Text = ""
    edcStnCode.Text = ""
    lbcCreditRestr.ListIndex = -1
    edcCreditLimit = ""
    lbcPaymRating.ListIndex = -1
    lbcCreditApproval.ListIndex = -1
    edcRating.Text = ""
    edcXMLBand.Text = ""
    imDigitRating = -1
    imISCI = -1
    imPackage = -1
    lbcInvSort.ListIndex = -1
    For ilLoop = 0 To 2 Step 1
        edcCAddr(ilLoop).Text = ""
        edcBAddr(ilLoop).Text = ""
    Next ilLoop
    
    edcRefId.Text = "" 'L.Bianchi 04/15/2021
    lbcTerms.ListIndex = -1 'L.Bianchi 04/15/2021
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    lbcBuyer.Clear
    'lbcBuyerCode.Clear
    'lbcBuyerCode.Tag = ""
    ReDim tmBuyerCode(0 To 0) As SORTCODE
    smBuyerCodeTag = ""
    lbcBuyer.ListIndex = -1
    lbcPayable.Clear
    'lbcPayableCode.Clear
    'lbcPayableCode.Tag = ""
    ReDim tmPayableCode(0 To 0) As SORTCODE
    smPayableCodeTag = ""
    lbcPayable.ListIndex = -1
    lbcLkBox.ListIndex = -1
    lbcEDI(0).ListIndex = -1
    lbcEDI(1).ListIndex = -1
    imExport = -1
    imXMLDates = -1
    imPrtStyle = -1
    lbcTax.ListIndex = -1
    smSPerson = ""
    smCRMID = ""
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
    imCreditRestrFirst = True
    imPaymRatingFirst = True
'    imPrtStyleFirst = True
    imScriptFirst = True
    imTaxFirst = True
    imCommFirst = True
    bmPDFEMailChgd = False
    ReDim tmPdf(0 To 0) As PDF
    tmCtrls(CRMIDINDEX).sShow = ""
    edcCRMID.Text = ""
    tmAgfx.lCrmId = 0
    tmAgfx.iInvFeatures = 0
    edcSuppressNet.Text = "  " ' TTP 10622 - 2023-03-08 JJB
    imSuppressNet = -1         ' TTP 10622 - 2023-03-08 JJB
End Sub
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
    'Screen.MousePointer = vbHourGlass
    sgArfCallType = "A"
    igArfCallSource = CALLSOURCEAGENCY
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
            slStr = "Agency^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    If sgArfName = "PDF EMail" Then
        ReDim tgPdf(0 To UBound(tmPdf)) As PDF
        For ilPdf = 0 To UBound(tmPdf) - 1 Step 1
            tgPdf(ilPdf) = tmPdf(ilPdf)
        Next ilPdf
        sgCommandStr = smName
        PDFEMailPersonnel.Show vbModal
    Else
        sgCommandStr = slStr
        NmAddr.Show vbModal
    End If
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
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
'*            Comments: Populate Agency DP service list*
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
    ilOffSet(0) = gFieldOffset("Arf", "ArfType")
    ilEDIC = lbcEDI(0).ListIndex
    ilEDII = lbcEDI(1).ListIndex
    If ilEDIC >= 1 Then
        slEDIC = lbcEDI(0).List(ilEDIC)
    End If
    If ilEDII >= 1 Then
        slEDII = lbcEDI(1).List(ilEDII)
    End If
    If lbcEDI(0).ListCount <> lbcEDI(1).ListCount Then  'Required by Branch logic
        lbcEDI(0).Clear
    End If
    'ilRet = gIMoveListBox(Agency, lbcEDI(0), lbcEDICode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Agency, lbcEDI(0), tmEDICode(), smEDICodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEDIPopErr
        gCPErrorMsg ilRet, "mEDIPop (gIMoveListBox)", Agency
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
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            'mNameCityPop True, True
            'If imTerminate Then
            '    Exit Sub
            'End If
            lbcName.height = gListBoxHeight(lbcName.ListCount, 12)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcName.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            gFindMatch smName, 0, lbcName
            If gLastFound(lbcName) >= 0 Then
                imChgMode = True
                lbcName.ListIndex = gLastFound(lbcName)
                edcDropDown.Text = lbcName.List(lbcName.ListIndex)
                imChgMode = False
            Else
                If smName <> "" Then
                    imChgMode = True
                    lbcName.ListIndex = -1
                    edcDropDown.Text = smName
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcName.ListIndex = -1
                    edcDropDown.Text = ""
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ABBRINDEX 'Abbreviation
            edcAbbr.Width = tmCtrls(ilBoxNo).fBoxW
            edcAbbr.MaxLength = 7
            gMoveFormCtrl pbcAgy(imPbcIndex), edcAbbr, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAbbr.Visible = True  'Set visibility
            edcAbbr.SetFocus
        Case CITYINDEX 'City
            'mNameCityPop True, True
            'If imTerminate Then
            '    Exit Sub
            'End If
            lbcCity.height = gListBoxHeight(lbcCity.ListCount, 12)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 5
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCity.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            gFindMatch smCity, 0, lbcCity
            If gLastFound(lbcCity) >= 0 Then
                imChgMode = True
                lbcCity.ListIndex = gLastFound(lbcCity)
                edcDropDown.Text = lbcCity.List(lbcCity.ListIndex)
                imChgMode = False
            Else
                If smCity <> "" Then
                    imChgMode = True
                    lbcCity.ListIndex = -1
                    edcDropDown.Text = smCity
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcCity.ListIndex = -1
                    edcDropDown.Text = ""
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STATEINDEX   'Active/Dormant
            If imState < 0 Then
                imState = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
        Case COMMINDEX 'Commission
            edcComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcComm.MaxLength = 6
            gMoveFormCtrl pbcAgy(imPbcIndex), edcComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcComm.Visible = True  'Set visibility
            edcComm.SetFocus
        Case SPERSONINDEX   'Salesperson
            mSPersonPop
            If imTerminate Then
                Exit Sub
            End If
            lbcSPerson.height = gListBoxHeight(lbcSPerson.ListCount, 11)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
        Case DIGITRATINGINDEX   'ISCI on Invoices
            If imDigitRating < 0 Then
                imDigitRating = 1  '1=one digit rating
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcDigitRating.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcDigitRating, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcDigitRating_Paint
            pbcDigitRating.Visible = True
            pbcDigitRating.SetFocus
        Case REPCODEINDEX 'Rep Agency Code
            If Trim$(edcRepCode.Text) = "" Then
                edcRepCode.Text = gGetNextGPNo(hmAdf, hmAgf, hmSaf)
            End If
            edcRepCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcRepCode.MaxLength = 10
            gMoveFormCtrl pbcAgy(imPbcIndex), edcRepCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRepCode.Visible = True  'Set visibility
            edcRepCode.SetFocus
        Case STNCODEINDEX 'Station Agency Code
            edcStnCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcStnCode.MaxLength = 10
            gMoveFormCtrl pbcAgy(imPbcIndex), edcStnCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcStnCode.Visible = True  'Set visibility
            edcStnCode.SetFocus
        Case CREDITAPPROVALINDEX 'Credit Approval
            lbcCreditApproval.height = gListBoxHeight(lbcCreditApproval.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 8
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCreditApproval.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcCreditApproval.Width, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcCreditApproval.ListIndex < 0 Then
                If (tgUrf(0).sChgCrRt <> "I") Then
                    lbcCreditApproval.ListIndex = 0   'Requires Checking
                Else
                    lbcCreditApproval.ListIndex = 1   'Approved
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
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2 - cmcDropDown.Width
            edcDropDown.MaxLength = 24
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
            gMoveFormCtrl pbcAgy(imPbcIndex), edcCreditLimit, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCreditLimit.Visible = True  'Set visibility
            edcCreditLimit.SetFocus
        Case PAYMRATINGINDEX 'Payment rating
            lbcPaymRating.height = gListBoxHeight(lbcPaymRating.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 13
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcPaymRating.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If lbcPaymRating.ListIndex < 0 Then
                lbcPaymRating.ListIndex = 1   'Normal
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
            gMoveFormCtrl pbcAgy(imPbcIndex), edcRating, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRating.Visible = True  'Set visibility
            edcRating.SetFocus
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
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcISCI, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcISCI_Paint
            pbcISCI.Visible = True
            pbcISCI.SetFocus
        Case INVSORTINDEX   'Invoice sorting
            mInvSortPop
            If imTerminate Then
                Exit Sub
            End If
            lbcInvSort.height = gListBoxHeight(lbcInvSort.ListCount, 6)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX - tmCtrls(ilBoxNo).fBoxW, tmCtrls(ilBoxNo).fBoxY
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
        Case PACKAGEINDEX   'ISCI on Invoices
            If imPackage < 0 Then
                imPackage = 1   'Time jim 6/20/97 was 0    'Daypart
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcPackage.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcPackage, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcPackage_Paint
            pbcPackage.Visible = True
            pbcPackage.SetFocus
        Case CADDRINDEX 'Contract Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Billing Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcAgy(imPbcIndex), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BUYERINDEX 'Product
            mBuyerPop tmAgf.iCode, "", -1
            If imTerminate Then
                Exit Sub
            End If
            lbcBuyer.height = gListBoxHeight(lbcBuyer.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 64
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
            mPayablePop tmAgf.iCode, "", -1
            If imTerminate Then
                Exit Sub
            End If
            lbcPayable.height = gListBoxHeight(lbcPayable.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 64
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
            
        Case CRMIDINDEX
            'edcCRMID.Text = CStr(tmAgfx.lCrmId)
            edcCRMID.Width = tmCtrls(ilBoxNo).fBoxW
            edcCRMID.MaxLength = 10
            gMoveFormCtrl pbcAgy(imPbcIndex), edcCRMID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCRMID.Visible = True  'Set visibility
            edcCRMID.SetFocus
            
        Case LKBOXINDEX 'Lock box
            mLkBoxPop
            If imTerminate Then
                Exit Sub
            End If
            lbcLkBox.height = gListBoxHeight(lbcLkBox.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLkBox.Move edcDropDown.Left, edcDropDown.Top - lbcLkBox.height
            imChgMode = True
            If lbcLkBox.ListIndex < 0 Then
                If lbcLkBox.ListCount > 2 Then
                    lbcLkBox.ListIndex = 2  'Pick first name
                Else
                    lbcLkBox.ListIndex = 1   '[New]
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
        Case REFIDINDEX 'L.Bianchi 04/15/2021
            edcRefId.Width = tmCtrls(ilBoxNo).fBoxW
            edcRefId.MaxLength = 36
            gMoveFormCtrl pbcAgy(imPbcIndex), edcRefId, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRefId.Visible = True  'Set visibility
            edcRefId.SetFocus
        Case EDICINDEX   'EDI service for Contracts
            mEDIPop
            If imTerminate Then
                Exit Sub
            End If
            lbcEDI(0).height = gListBoxHeight(lbcEDI(0).ListCount, 7)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
        Case TERMSINDEX   'Invoice sorting
            mTermsPop
            If imTerminate Then
                Exit Sub
            End If
            lbcTerms.height = gListBoxHeight(lbcTerms.ListCount, 6)
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2 + cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
            gMoveFormCtrl pbcAgy(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
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
        Case EXPORTFORMINDEX   'Contract Export form
            If imExport < 0 Then
                imExport = 0    'CSI Form
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcExport.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcExport, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcExport_Paint
            pbcExport.Visible = True
            pbcExport.SetFocus
        Case SUPPRESSNETINDEX   'Suppress Net Amount for Trade Invoices  ' TTP 10622 - 2023-03-08 JJB
            If imSuppressNet < 0 Then
                imSuppressNet = -1
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcSuppressNet.Width = 2695 'tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcSuppressNet, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcSuppressNet_Paint
            pbcSuppressNet.Visible = True
            pbcSuppressNet.SetFocus
'        Case PRTSTYLEINDEX   'Print style
'            If imPrtStyle < 0 Then
'                imPrtStyle = 1    'Narrow
'                tmCtrls(ilBoxNo).iChg = True
'                mSetCommands
'            End If
'            pbcPrtStyle.Width = tmCtrls(ilBoxNo).fBoxW
'            gMoveFormCtrl pbcAgy(imPbcIndex), pbcPrtStyle, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
'            pbcPrtStyle_Paint
'            pbcPrtStyle.Visible = True
'            pbcPrtStyle.SetFocus
        Case XMLCALLINDEX 'XML Proposal Call Letters
            edcXMLCall.Width = tmCtrls(ilBoxNo).fBoxW
            edcXMLCall.MaxLength = 4
            gMoveFormCtrl pbcAgy(imPbcIndex), edcXMLCall, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcXMLCall.Visible = True  'Set visibility
            edcXMLCall.SetFocus
        Case XMLBANDINDEX 'XML Proposal Band
            edcXMLBand.Width = tmCtrls(ilBoxNo).fBoxW
            edcXMLBand.MaxLength = 2
            gMoveFormCtrl pbcAgy(imPbcIndex), edcXMLBand, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcXMLBand.Visible = True  'Set visibility
            edcXMLBand.SetFocus
        Case XMLDATESINDEX   'XML Dates
            If imXMLDates < 0 Then
                imXMLDates = 0    'M-Su Form
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcXMLDates.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAgy(imPbcIndex), pbcXMLDates, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcXMLDates_Paint
            pbcXMLDates.Visible = True
            pbcXMLDates.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
'   mInitParameters
'   Where:
'
    Dim ilRet As Integer    'Return Status
    imLBCtrls = 1
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    Agency.height = cmcReport.Top + 5 * cmcReport.height / 3
    gCenterStdAlone Agency
    'Agency.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imAgfRecLen = Len(tmAgf)  'Get and save AGF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imEdcChgMode = False
    imBSMode = False
    imBypassSetting = False
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) And ((tgSpf.iBTax(1) = 0)) Then
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
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
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

    hmPnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPnf, "", sgDBPath & "PNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PNF.Btr)", Agency
    On Error GoTo 0
    imPnfRecLen = Len(tmBPnf)
    hmPDF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPDF, "", sgDBPath & "PDF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PDF.Btr)", Agency
    On Error GoTo 0
    ReDim tmPdf(0 To 0) As PDF
    imPdfRecLen = Len(tmPdf(0))
    hmSaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "SAF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SAF.Btr)", Agency
    On Error GoTo 0
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", Agency
    On Error GoTo 0
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "DSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: DSF.Btr)", Agency
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    hmAgf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", Agency
    On Error GoTo 0
    lbcSPerson.Clear    'Force list box to be populated
    mSPersonPop
    If imTerminate Then
        Exit Sub
    End If
    lbcLkBox.Clear 'Force list box to be populated
    mLkBoxPop
    If imTerminate Then
        Exit Sub
    End If
    lbcEDI(0).Clear 'Force list box to be populated
    lbcEDI(1).Clear
    mEDIPop
    If imTerminate Then
        Exit Sub
    End If
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
    lbcName.Clear 'Force list box to be populated
    lbcCity.Clear
    'mNameCityPop True, True
    'If imTerminate Then
    '    Exit Sub
    'End If
'    gCenterModalForm Agency
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    imCreditRestrFirst = True
    imPaymRatingFirst = True
'    imPrtStyleFirst = True
    imScriptFirst = True
    imTaxFirst = True
    imCommFirst = True
    plcScreen_Paint
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
    Dim llYAdj As Long
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim ilGap As Integer
    Dim llMax As Long
    Dim llShortMax As Long
    
    flTextHeight = pbcAgy(imPbcIndex).TextHeight("1") - 35
    'Position panel and picture areas with panel
    pbcAgy(imPbcIndex).Visible = True
    If imPbcIndex = 1 Then
        pbcAgy(0).Visible = False
        llYAdj = pbcAgy(imPbcIndex).height + 100
    Else
        pbcAgy(1).Visible = False
        llYAdj = 0
    End If
    plcBkgd.Move 210, 540, pbcAgy(imPbcIndex).Width + fgPanelAdj, pbcAgy(imPbcIndex).height + fgPanelAdj
    pbcAgy(imPbcIndex).Move plcBkgd.Left + fgBevelX, plcBkgd.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Abbreviation
    gSetCtrl tmCtrls(ABBRINDEX), 2850, tmCtrls(NAMEINDEX).fBoxY, 1395, fgBoxStH
    'City ID
    gSetCtrl tmCtrls(CITYINDEX), 4260, tmCtrls(NAMEINDEX).fBoxY, 1395, fgBoxStH
    'State
    gSetCtrl tmCtrls(STATEINDEX), 5670, tmCtrls(NAMEINDEX).fBoxY, 1395, fgBoxStH
    '% commission
    gSetCtrl tmCtrls(COMMINDEX), 7080, tmCtrls(NAMEINDEX).fBoxY + llYAdj, 1400, fgBoxStH
    'Salesperson
    gSetCtrl tmCtrls(SPERSONINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY + llYAdj, 2805, fgBoxStH
    tmCtrls(SPERSONINDEX).iReq = False
    '1 or 2 Digit Rating
    gSetCtrl tmCtrls(DIGITRATINGINDEX), 2850, tmCtrls(SPERSONINDEX).fBoxY, 1395, fgBoxStH
    tmCtrls(DIGITRATINGINDEX).iReq = False
    'Rep Agency Code
    gSetCtrl tmCtrls(REPCODEINDEX), 4260, tmCtrls(SPERSONINDEX).fBoxY, 1395, fgBoxStH
    tmCtrls(REPCODEINDEX).iReq = False
    'Station Agency Code
    gSetCtrl tmCtrls(STNCODEINDEX), 5670, tmCtrls(SPERSONINDEX).fBoxY, 1395, fgBoxStH
    tmCtrls(STNCODEINDEX).iReq = False
    'Credit Approval
    gSetCtrl tmCtrls(CREDITAPPROVALINDEX), 7080, tmCtrls(SPERSONINDEX).fBoxY, 1400, fgBoxStH
    If tgUrf(0).sChgCrRt <> "I" Then
        tmCtrls(CREDITAPPROVALINDEX).iReq = False
    End If
    'Credit restriction
    gSetCtrl tmCtrls(CREDITRESTRINDEX), 30, tmCtrls(SPERSONINDEX).fBoxY + fgStDeltaY + llYAdj, 1755, fgBoxStH
    If tgUrf(0).sCredit <> "I" Then
        tmCtrls(CREDITRESTRINDEX).iReq = False
    End If
    'gSetCtrl tmCtrls(CREDITRESTRINDEX + 1), 1770, tmCtrls(CREDITRESTRINDEX).fBoxY, 1065, fgBoxStH
    gSetCtrl tmCtrls(CREDITRESTRINDEX + 1), 1770, tmCtrls(CREDITRESTRINDEX).fBoxY, 1050, fgBoxStH
    tmCtrls(CREDITRESTRINDEX + 1).iReq = False
    'Payment Rating
    gSetCtrl tmCtrls(PAYMRATINGINDEX), 2860, tmCtrls(CREDITRESTRINDEX).fBoxY, 1385, fgBoxStH
    If tgUrf(0).sPayRate <> "I" Then
        tmCtrls(PAYMRATINGINDEX).iReq = False
    End If
    'Credit Rating
    gSetCtrl tmCtrls(CREDITRATINGINDEX), 4260, tmCtrls(CREDITRESTRINDEX).fBoxY, 885, fgBoxStH
    tmCtrls(CREDITRATINGINDEX).iReq = False
    'ISCI on invoices
    gSetCtrl tmCtrls(ISCIINDEX), 5160, tmCtrls(CREDITRESTRINDEX).fBoxY, 1050, fgBoxStH
    If (tgSpf.sAISCI = "A") Or (tgSpf.sAISCI = "X") Then
        tmCtrls(ISCIINDEX).iReq = False
    End If
    'Invoice sorting
    gSetCtrl tmCtrls(INVSORTINDEX), 6225, tmCtrls(CREDITRESTRINDEX).fBoxY, 840, fgBoxStH
    tmCtrls(INVSORTINDEX).iReq = False
    'Package Invoice Show
    gSetCtrl tmCtrls(PACKAGEINDEX), 7080, tmCtrls(CREDITRESTRINDEX).fBoxY, 1400, fgBoxStH
    tmCtrls(PACKAGEINDEX).iReq = False
    'Contract Address
    If imShortForm Then
        gSetCtrl tmCtrls(CADDRINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 4215, fgBoxStH
    Else
        gSetCtrl tmCtrls(CADDRINDEX), 30, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 4220, fgBoxStH
    End If
    gSetCtrl tmCtrls(CADDRINDEX + 1), 30, tmCtrls(CADDRINDEX).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(CADDRINDEX + 2), 30, tmCtrls(CADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 2).iReq = False
    'Billing Address
    If imShortForm Then
        'gSetCtrl tmCtrls(BADDRINDEX), 4260, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 4215, fgBoxStH
        gSetCtrl tmCtrls(BADDRINDEX), 4260, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 4230, fgBoxStH
    Else
        'gSetCtrl tmCtrls(BADDRINDEX), 4260, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 4215, fgBoxStH
        gSetCtrl tmCtrls(BADDRINDEX), 4260, tmCtrls(CREDITRESTRINDEX).fBoxY + fgStDeltaY, 4230, fgBoxStH
    End If
    tmCtrls(BADDRINDEX).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 1), 4260, tmCtrls(BADDRINDEX).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 2), 4260, tmCtrls(BADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 2).iReq = False
        
        'Buyer Name
    gSetCtrl tmCtrls(BUYERINDEX), 30, tmCtrls(CADDRINDEX).fBoxY + fgAddDeltaY, 5735, fgBoxStH
    tmCtrls(BUYERINDEX).iReq = False
    'Payable Contact Name
    gSetCtrl tmCtrls(PAYABLEINDEX), 30, tmCtrls(BUYERINDEX).fBoxY + fgStDeltaY + llYAdj, 4560, fgBoxStH
    tmCtrls(PAYABLEINDEX).iReq = False
    
    gSetCtrl tmCtrls(CRMIDINDEX), 4599, tmCtrls(BUYERINDEX).fBoxY + fgStDeltaY + llYAdj, 1160, fgBoxStH
    tmCtrls(CRMIDINDEX).iReq = False
    
    'LkBox
    If imShortForm Then
        gSetCtrl tmCtrls(LKBOXINDEX), 5775, tmCtrls(PAYABLEINDEX).fBoxY, 2715, fgBoxStH
    Else
        gSetCtrl tmCtrls(LKBOXINDEX), 5775, tmCtrls(BUYERINDEX).fBoxY, 2715, fgBoxStH
    End If
    tmCtrls(LKBOXINDEX).iReq = False
    'RefBox L.Bianchi
    gSetCtrl tmCtrls(REFIDINDEX), 5775, tmCtrls(PAYABLEINDEX).fBoxY, 2705, fgBoxStH
    tmCtrls(REFIDINDEX).iReq = False
    
    'EDI Service Contract
    gSetCtrl tmCtrls(EDICINDEX), 30, tmCtrls(PAYABLEINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
    tmCtrls(EDICINDEX).iReq = False
    'EDI Service Contract
    gSetCtrl tmCtrls(EDIIINDEX), 1935, tmCtrls(EDICINDEX).fBoxY, 1890, fgBoxStH
    tmCtrls(EDIIINDEX).iReq = False
    'Contract Print Style
'    gSetCtrl tmCtrls(PRTSTYLEINDEX), 5175, tmCtrls(EDICINDEX).fBoxY, 1320, fgBoxStH
'    If (tgSpf.sAPrtStyle = "W") Or (tgSpf.sAPrtStyle = "N") Then
'        tmCtrls(PRTSTYLEINDEX).iReq = False
'    End If
    'Terms
    gSetCtrl tmCtrls(TERMSINDEX), 3840, tmCtrls(EDICINDEX).fBoxY, 1200, fgBoxStH
    tmCtrls(TERMSINDEX).iReq = False
    'Sales Tax
    gSetCtrl tmCtrls(TAXINDEX), 5055, tmCtrls(EDICINDEX).fBoxY, 1050, fgBoxStH
    If Not imTaxDefined Then
        tmCtrls(TAXINDEX).iReq = False
    End If
    'Contract Export form
    gSetCtrl tmCtrls(EXPORTFORMINDEX), 6735, tmCtrls(EDICINDEX).fBoxY, 375, fgBoxStH
    tmCtrls(EXPORTFORMINDEX).iReq = False
    'Call Letter override for XML Proposal
    gSetCtrl tmCtrls(XMLCALLINDEX), 7125, tmCtrls(EDICINDEX).fBoxY, 540, fgBoxStH
    tmCtrls(XMLCALLINDEX).iReq = False
    'Band override for XML Proposal
    gSetCtrl tmCtrls(XMLBANDINDEX), 7680, tmCtrls(EDICINDEX).fBoxY, 375, fgBoxStH
    tmCtrls(XMLBANDINDEX).iReq = False
    'XML Proposal date form
    gSetCtrl tmCtrls(XMLDATESINDEX), 8070, tmCtrls(EDICINDEX).fBoxY, 405, fgBoxStH
    tmCtrls(XMLDATESINDEX).iReq = False
    
    'Suppress Net Amount For Trade Invoices ' TTP 10622 - 2023-03-08 JJB
    gSetCtrl tmCtrls(SUPPRESSNETINDEX), 30, tmCtrls(EDICINDEX).fBoxY + 350, 3825, fgBoxStH
    tmCtrls(SUPPRESSNETINDEX).iReq = False
    'Unused
    gSetCtrl tmCtrls(UNUSEDINDEX), 3825, tmCtrls(EDICINDEX).fBoxY + 350, 1400, fgBoxStH
    tmCtrls(UNUSEDINDEX).iReq = False
    
    If imPbcIndex = 1 Then
        tmCtrls(COMMINDEX).iReq = False
        tmCtrls(CREDITAPPROVALINDEX).iReq = False
        tmCtrls(CREDITRESTRINDEX).iReq = False
        tmCtrls(PAYMRATINGINDEX).iReq = False
        tmCtrls(ISCIINDEX).iReq = False
        'tmCtrls(PRTSTYLEINDEX).iReq = False
        tmCtrls(XMLBANDINDEX).iReq = False
        tmCtrls(XMLCALLINDEX).iReq = False
        tmCtrls(XMLDATESINDEX).iReq = False
        tmCtrls(TAXINDEX).iReq = False
    End If
    '10/25/14: One pixel removed from top and left side when using macromedia fireworks
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilBox).fBoxX = tmCtrls(ilBox).fBoxX - 15
        tmCtrls(ilBox).fBoxY = tmCtrls(ilBox).fBoxY - 15
    Next ilBox
    
    
    'Pct 90
    gSetCtrl tmARCtrls(PCT90INDEX), 30, 3180, 1515, fgBoxStH
    'Currect A/R
    gSetCtrl tmARCtrls(CURRARINDEX), 1560, tmARCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Unbilled
    gSetCtrl tmARCtrls(UNBILLEDINDEX), 3090, tmARCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Hi Credit
    gSetCtrl tmARCtrls(HICREDITINDEX), 4620, tmARCtrls(PCT90INDEX).fBoxY, 1515, fgBoxStH
    'Total Gross
    gSetCtrl tmARCtrls(TOTALGROSSINDEX), 6150, tmARCtrls(PCT90INDEX).fBoxY, 1260, fgBoxStH
    'Date
    gSetCtrl tmARCtrls(DATEENTRDINDEX), 7440, tmARCtrls(PCT90INDEX).fBoxY, 1035, fgBoxStH
    'NSF Checkes
    gSetCtrl tmARCtrls(NSFCHKSINDEX), 30, 3520, 1515, fgBoxStH
    'Date Last Billed
    gSetCtrl tmARCtrls(DATELSTINVINDEX), 1560, tmARCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Date Last Payment
    gSetCtrl tmARCtrls(DATELSTPAYMINDEX), 3090, tmARCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Avg # days to pay
    gSetCtrl tmARCtrls(AVGTOPAYINDEX), 4620, tmARCtrls(NSFCHKSINDEX).fBoxY, 1515, fgBoxStH
    'Avg # days to pay
    gSetCtrl tmARCtrls(LSTTOPAYINDEX), 6150, tmARCtrls(NSFCHKSINDEX).fBoxY, 2320, fgBoxStH
    
    For ilBox = PCT90INDEX To LSTTOPAYINDEX Step 1
        tmARCtrls(ilBox).fBoxX = tmARCtrls(ilBox).fBoxX - 15
        tmARCtrls(ilBox).fBoxY = tmARCtrls(ilBox).fBoxY - 15
    Next ilBox
    
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
        'If ilLoop <= POLITICALINDEX Then
        '    If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llShortMax Then
        '        llShortMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        '    End If
        'End If

    Next ilLoop
    
    
    llMax = 0
    For ilLoop = PCT90INDEX To LSTTOPAYINDEX Step 1
        If tmARCtrls(ilLoop).fBoxX >= 0 Then
            tmARCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmARCtrls(ilLoop).fBoxW)
            Do While (tmARCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmARCtrls(ilLoop).fBoxW = tmARCtrls(ilLoop).fBoxW + 1
            Loop
            tmARCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmARCtrls(ilLoop).fBoxX)
            Do While (tmARCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmARCtrls(ilLoop).fBoxX = tmARCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmARCtrls(ilLoop).fBoxX > 90 Then
                    If tmARCtrls(ilLoop - 1).fBoxX + tmARCtrls(ilLoop - 1).fBoxW + 15 < tmARCtrls(ilLoop).fBoxX Then
                        tmARCtrls(ilLoop - 1).fBoxW = tmARCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmARCtrls(ilLoop - 1).fBoxX + tmARCtrls(ilLoop - 1).fBoxW + 15 > tmARCtrls(ilLoop).fBoxX Then
                        tmARCtrls(ilLoop - 1).fBoxW = tmARCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmARCtrls(ilLoop).fBoxX + tmARCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmARCtrls(ilLoop).fBoxX + tmARCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    tmARCtrls(LSTTOPAYINDEX).fBoxW = tmARCtrls(LSTTOPAYINDEX).fBoxW + 15
    
    pbcAgy(0).Picture = LoadPicture("")
    pbcAgy(1).Picture = LoadPicture("")
    If imShortForm Then
        llShortMax = llMax
        pbcAgy(1).Width = llShortMax + 15
        'pbcAgy(1).Width = llShortMax + 2 * fgBevelX + 30
    Else
        pbcAgy(0).Width = llMax
        'pbcAgy(0).Width = llMax + 2 * fgBevelX + 15
    End If
    plcBkgd.Width = llMax + 2 * fgBevelX + 15
    
    cbcSelect.Left = plcBkgd.Left + plcBkgd.Width - cbcSelect.Width
    lacCode.Left = plcBkgd.Left + plcBkgd.Width - lacCode.Width
    
    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    
    cmcErase.Left = Agency.Width / 2 - cmcErase.Width / 2
    cmcUpdate.Left = cmcErase.Left - cmcUpdate.Width - ilGap
    cmcCancel.Left = cmcUpdate.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    
    cmcUndo.Left = cmcErase.Left + cmcErase.Width + ilGap
    cmcMerge.Left = cmcUndo.Left + cmcUndo.Width + ilGap
    cmcReport.Left = cmcMerge.Left + cmcMerge.Width + ilGap
    pbcAgy(0).BackColor = WHITE
    pbcAgy(1).BackColor = WHITE
    
    
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
    igMNmCallSource = CALLSOURCEAGENCY
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
            slStr = "Agency^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
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
    'ilRet = gIMoveListBox(Agency, lbcInvSort, lbcInvSortCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Agency, lbcInvSort, tmInvSortCode(), smInvSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvSortPopErr
        gCPErrorMsg ilRet, "mInvSortPop (gIMoveListBox)", Agency
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
        imDoubleClickName = False
        mLkBoxBranch = False
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
    'Screen.MousePointer = vbHourGlass
    sgArfCallType = "L"
    igArfCallSource = CALLSOURCEAGENCY
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
            slStr = "Agency^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
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
    'ilRet = gIMoveListBox(Agency, lbcLkBox, lbcLkBoxCode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Agency, lbcLkBox, tmLkBoxCode(), smLkBoxCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLkBoxPopErr
        gCPErrorMsg ilRet, "mLkBoxPop (gIMoveListBox)", Agency
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
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmAgf.sName = smName
    End If
    If Not ilTestChg Or tmCtrls(ABBRINDEX).iChg Then
        tmAgf.sAbbr = edcAbbr.Text
    End If
    If Not ilTestChg Or tmCtrls(CITYINDEX).iChg Then
        tmAgf.sCityID = smCity
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmAgf.sState = "A"
            Case 1  'Dormant
                tmAgf.sState = "D"
            Case Else
                tmAgf.sState = "A"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(COMMINDEX).iChg Then
        slStr = edcComm.Text
        If (slStr = "") And (imShortForm) Then
            slStr = "15"
        End If
        'gStrToPDN slStr, 2, 3, tmAgf.sComm
        tmAgf.iComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(SPERSONINDEX).iChg Then
        If lbcSPerson.ListIndex >= 2 Then
            'If lbcSPerson.ListIndex <= Traffic!lbcSalesperson.ListCount + 1 Then    'Note: +2-1
                slNameCode = tgSalesperson(lbcSPerson.ListIndex - 2).sKey  'Traffic!lbcSalesperson.List(lbcSPerson.ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmAgf.iSlfCode = CInt(slCode)
            'Else
            '    slNameCode = Traffic!lbcSPersonCombo.List(lbcSPerson.ListIndex - 2 - Traffic!lbcSalesperson.ListCount)
            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    On Error GoTo mMoveCtrlToRecErr
            '    gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            '    On Error GoTo 0
            '    slCode = Trim$(slCode)
            '    tmAgf.iSlfCode = -CInt(slCode)
            'End If
        Else
            tmAgf.iSlfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(DIGITRATINGINDEX).iChg Then
        Select Case imDigitRating
            Case 1  'Yes
                tmAgf.s1or2DigitRating = "1"
            Case 2  'No
                tmAgf.s1or2DigitRating = "2"
            Case Else
                tmAgf.s1or2DigitRating = "1"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(REPCODEINDEX).iChg Then
        If tgSpf.sARepCodes = "N" Then
            tmAgf.sCodeRep = ""
        Else
            tmAgf.sCodeRep = edcRepCode.Text
        End If
    End If
    If Not ilTestChg Or tmCtrls(STNCODEINDEX).iChg Then
        If tgSpf.sAStnCodes = "N" Then
            tmAgf.sCodeStn = ""
        Else
            tmAgf.sCodeStn = edcStnCode.Text
        End If
    End If
    If Not ilTestChg Or tmCtrls(CREDITAPPROVALINDEX).iChg Then
        Select Case lbcCreditApproval.ListIndex
            Case 0  'Requires Checking
                tmAgf.sCrdApp = "R"
            Case 1  'Approved
                tmAgf.sCrdApp = "A"
            Case 2  'Denied
                tmAgf.sCrdApp = "D"
            Case Else
'                tmAgf.sCrdApp = "R"
                If (tgUrf(0).sChgCrRt <> "I") Then
                    tmAgf.sCrdApp = "R"   'Requires Checking
                Else
                    tmAgf.sCrdApp = "A"   'Approved
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CREDITRESTRINDEX).iChg Then
        Select Case lbcCreditRestr.ListIndex
            Case 0  'No restrictions
                tmAgf.sCreditRestr = "N"
            Case 1  'Credit Limit
                tmAgf.sCreditRestr = "L"
            Case 2  'cash in advance weekly
                tmAgf.sCreditRestr = "W"
            Case 3  'Cash in advance monthly
                tmAgf.sCreditRestr = "M"
            Case 4  'Cash in advance quarterly
                tmAgf.sCreditRestr = "T"
            Case 5  'Prohibit
                tmAgf.sCreditRestr = "P"
            Case Else
                If (tgUrf(0).sCredit <> "I") Or imShortForm Then
                    tmAgf.sCreditRestr = "N"
                Else
                    tmAgf.sCreditRestr = ""
                End If
        End Select
    End If
    If lbcCreditRestr.ListIndex <> 1 Then
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmAgf.sCreditLimit
        tmAgf.lCreditLimit = 0
    Else
        If Not ilTestChg Or tmCtrls(CREDITRESTRINDEX + 1).iChg Then
            slStr = edcCreditLimit.Text
            'gStrToPDN slStr, 2, 5, tmAgf.sCreditLimit
            tmAgf.lCreditLimit = gStrDecToLong(slStr, 2)
        End If
    End If
    If Not ilTestChg Or tmCtrls(PAYMRATINGINDEX).iChg Then
        Select Case lbcPaymRating.ListIndex
            Case 0  'Quick
                tmAgf.sPaymRating = "0"
            Case 1  'Normal
                tmAgf.sPaymRating = "1"
            Case 2  'Slow
                tmAgf.sPaymRating = "2"
            Case 3  'Difficult
                tmAgf.sPaymRating = "3"
            Case 4  'in Collection
                tmAgf.sPaymRating = "4"
            Case Else
                If (tgUrf(0).sPayRate <> "I") Or imShortForm Then
                    tmAgf.sPaymRating = "1"
                Else
                    tmAgf.sPaymRating = ""
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CREDITRATINGINDEX).iChg Then
        tmAgf.sCrdRtg = edcRating.Text
    End If
    If Not ilTestChg Or tmCtrls(ISCIINDEX).iChg Then
        Select Case imISCI
            Case 0  'Yes
                tmAgf.sShowISCI = "Y"
            Case 1  'No
                tmAgf.sShowISCI = "N"
            Case 2  'yes and W/O Leader
                tmAgf.sShowISCI = "W"
            Case Else
                If tgSpf.sAISCI = "A" Then
                    tmAgf.sShowISCI = "Y"
                ElseIf tgSpf.sAISCI = "X" Then
                    tmAgf.sShowISCI = "N"
                Else
                    tmAgf.sShowISCI = ""
                End If
        End Select
    End If
    If Not ilTestChg Or tmCtrls(INVSORTINDEX).iChg Then
        If lbcInvSort.ListIndex >= 2 Then
            slNameCode = tmInvSortCode(lbcInvSort.ListIndex - 2).sKey 'lbcInvSortCode.List(lbcInvSort.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iMnfSort = CInt(slCode)
        Else
            tmAgf.iMnfSort = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(PACKAGEINDEX).iChg Then
        Select Case imPackage
            Case 0  'Daypart
                tmAgf.sPkInvShow = "D"
            Case 1  'Time
                tmAgf.sPkInvShow = "T"
            Case Else
                '6/27/12:  To match change when tabbing into field
                tmAgf.sPkInvShow = "T"  '"D"
        End Select
    End If
    For ilLoop = 0 To 2 Step 1
        If Not ilTestChg Or tmCtrls(CADDRINDEX + ilLoop).iChg Then
            tmAgf.sCntrAddr(ilLoop) = edcCAddr(ilLoop).Text
        End If
    Next ilLoop
    For ilLoop = 0 To 2 Step 1
        If Not ilTestChg Or tmCtrls(BADDRINDEX + ilLoop).iChg Then
            tmAgf.sBillAddr(ilLoop) = edcBAddr(ilLoop).Text
        End If
    Next ilLoop
    If Not ilTestChg Or tmCtrls(BUYERINDEX).iChg Then
        If lbcBuyer.ListIndex >= 2 Then
            slNameCode = tmBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iPnfBuyer = CInt(slCode)
        Else
            tmAgf.iPnfBuyer = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(PAYABLEINDEX).iChg Then
        If lbcPayable.ListIndex >= 2 Then
            slNameCode = tmPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iPnfPay = CInt(slCode)
        Else
            tmAgf.iPnfPay = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(LKBOXINDEX).iChg Then
        If lbcLkBox.ListIndex >= 2 Then
            slNameCode = tmLkBoxCode(lbcLkBox.ListIndex - 2).sKey  'lbcLkBoxCode.List(lbcLkBox.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iArfLkCode = CInt(slCode)
        Else
            tmAgf.iArfLkCode = 0
        End If
    End If
    
    If Not ilTestChg Or tmCtrls(REFIDINDEX).iChg Then 'L.Bianchi 04/15/2021
        tmAgfx.sRefId = edcRefId.Text
    End If
    
    slStr = edcCRMID.Text
    tmAgfx.lCrmId = gStrDecToLong(slStr, 0)
'    If Not ilTestChg Or tmCtrls(CRMIDINDEX).iChg Then
'        If Not IsNull(edcCRMID.Text) And Len(edcCRMID.Text) > 0 Then
'            tmAgfx.lCrmId = CInt(edcCRMID.Text)
'        End If
'    End If
    
    If Not ilTestChg Or tmCtrls(EDICINDEX).iChg Then
        If lbcEDI(0).ListIndex >= 1 Then
            slNameCode = tmEDICode(lbcEDI(0).ListIndex - 1).sKey   'lbcEDICode.List(lbcEDI(0).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iArfCntrCode = CInt(slCode)
        Else
            tmAgf.iArfCntrCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(EDIIINDEX).iChg Then
        If lbcEDI(1).ListIndex >= 1 Then
            slNameCode = tmEDICode(lbcEDI(1).ListIndex - 1).sKey   'lbcEDICode.List(lbcEDI(1).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iArfInvCode = CInt(slCode)
        Else
            tmAgf.iArfInvCode = 0
        End If
    End If

'    If Not ilTestChg Or tmCtrls(PRTSTYLEINDEX).iChg Then
'        Select Case imPrtStyle
'            Case 0  'Wide
'                tmAgf.sCntrPrtSz = "W"
'            Case 1  'Narrow
'                tmAgf.sCntrPrtSz = "N"
'            Case Else
                If tgSpf.sAPrtStyle = "W" Then
                    tmAgf.sCntrPrtSz = "W"
                ElseIf tgSpf.sAPrtStyle = "N" Then
                    tmAgf.sCntrPrtSz = "N"
                Else
                    tmAgf.sCntrPrtSz = ""
                End If
'        End Select
'    End If

    If Not ilTestChg Or tmCtrls(TERMSINDEX).iChg Then
        If lbcTerms.ListIndex >= 2 Then
            slNameCode = tmTermsCode(lbcTerms.ListIndex - 2).sKey 'lbcInvSortCode.List(lbcInvSort.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Agency
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmAgf.iMnfInvTerms = CInt(slCode)
        Else
            tmAgf.iMnfInvTerms = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(TAXINDEX).iChg Then
        '12/17/06-Change to tax by agency or vehicle
        'Select Case lbcTax.ListIndex
        '    Case 0  'No; No
        '        tmAgf.sSlsTax(0) = "N"
        '        tmAgf.sSlsTax(1) = "N"
        '    Case 1  'Yes; No
        '        tmAgf.sSlsTax(0) = "Y"
        '        tmAgf.sSlsTax(1) = "N"
        '    Case 2  'No; Yes
        '        tmAgf.sSlsTax(0) = "N"
        '        tmAgf.sSlsTax(1) = "Y"
        '    Case 3  'Yes; Yes
        '        tmAgf.sSlsTax(0) = "Y"
        '        tmAgf.sSlsTax(1) = "Y"
        '    Case Else
        '        If Not imTaxDefined Then
        '            tmAgf.sSlsTax(0) = "N"
        '            tmAgf.sSlsTax(1) = "N"
        '        Else
        '            tmAgf.sSlsTax(0) = ""
        '            tmAgf.sSlsTax(1) = ""
        '        End If
        'End Select
        If lbcTax.ListIndex >= 1 Then
            tmAgf.iTrfCode = lbcTax.ItemData(lbcTax.ListIndex)
        Else
            tmAgf.iTrfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(EXPORTFORMINDEX).iChg Then
        Select Case imExport
            Case 0  'CSI Form
                tmAgf.sCntrExptForm = "C"
            Case 1  'Agency (OMD) Form
                tmAgf.sCntrExptForm = "O"
            Case 2  'Agency (XML) Form
                tmAgf.sCntrExptForm = "X"
            Case Else
                tmAgf.sCntrExptForm = "C"
        End Select
    End If
    
    If Not ilTestChg Or tmCtrls(SUPPRESSNETINDEX).iChg Then ' TTP 10622 - 2023-03-08 JJB
        Select Case imSuppressNet
            Case 0  'No
                tmAgfx.iInvFeatures = 0
            Case 1  'Yes
                tmAgfx.iInvFeatures = 1
        End Select
    End If
    
    If Not ilTestChg Or tmCtrls(XMLCALLINDEX).iChg Then
        tmAgf.sXMLCallLetters = edcXMLCall.Text
    End If
    If Not ilTestChg Or tmCtrls(XMLBANDINDEX).iChg Then
        tmAgf.sXMPProposalBand = edcXMLBand.Text
    End If
    If Not ilTestChg Or tmCtrls(XMLDATESINDEX).iChg Then
        Select Case imXMLDates
            Case 0  'M-Su
                tmAgf.sXMLDates = "M"
            Case 1  'Aired
                tmAgf.sXMLDates = "A"
            Case Else
                tmAgf.sXMLDates = "M"
        End Select
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
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
    smName = Trim$(tmAgf.sName)
    edcAbbr.Text = Trim$(tmAgf.sAbbr)
    edcRefId.Text = Trim$(tmAgfx.sRefId)
    
    edcCRMID.Text = ""
    If tmAgfx.lCrmId <> 0 Then
        edcCRMID.Text = gLongToStrDec(tmAgfx.lCrmId, 0)
    End If
    'slStr = gLongToStrDec(tmAgfx.lCrmId, 0)
    'gFormatStr slStr, FMTLEAVEBLANK, 0, edcCRMID.Text
    'edcCRMID.Text = slStr
    
    smCity = Trim$(tmAgf.sCityID)
    If tmAgf.sState = "D" Then
        imState = 1 'Dormant
    Else
        imState = 0 'Active
    End If
    'gPDNToStr tmAgf.sComm, 2, slStr
    slStr = gIntToStrDec(tmAgf.iComm, 2)
    edcComm.Text = slStr
    'look up salesperson name from code number
    lbcSPerson.ListIndex = 1
    smSPerson = ""
    'If tmAgf.iSlfCode >= 0 Then
        slRecCode = Trim$(str$(tmAgf.iSlfCode))
        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
            On Error GoTo 0
            If slRecCode = slCode Then
                lbcSPerson.ListIndex = ilLoop + 2
                smSPerson = lbcSPerson.List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    'Else
    '    slRecCode = Trim$(Str$(-tmAgf.iSlfCode))
    '    For ilLoop = 0 To Traffic!lbcSPersonCombo.ListCount - 1 Step 1
    '        slNameCode = Traffic!lbcSPersonCombo.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '        On Error GoTo mMoveRecToCtrlErr
    '        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
    '        On Error GoTo 0
    '        If slRecCode = slCode Then
    '            lbcSPerson.ListIndex = ilLoop + 2 + Traffic!lbcSalesperson.ListCount
    '            smSPerson = lbcSPerson.List(ilLoop + 2 + Traffic!lbcSalesperson.ListCount)
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If
    Select Case tmAgf.s1or2DigitRating
        Case "1"
            imDigitRating = 1
        Case "2"
            imDigitRating = 2
        Case Else
            imDigitRating = -1
    End Select
    edcRepCode.Text = Trim$(tmAgf.sCodeRep)
    edcStnCode.Text = Trim$(tmAgf.sCodeStn)
    Select Case tmAgf.sCreditRestr
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
    If tmAgf.sCreditRestr = "L" Then
        'gPDNToStr tmAgf.sCreditLimit, 2, slStr
        slStr = gLongToStrDec(tmAgf.lCreditLimit, 2)
        edcCreditLimit.Text = slStr
    Else
        edcCreditLimit.Text = ""
    End If
    Select Case tmAgf.sPaymRating
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
    Select Case tmAgf.sCrdApp
        Case "R"
            lbcCreditApproval.ListIndex = 0
        Case "A"
            lbcCreditApproval.ListIndex = 1
        Case "D"
            lbcCreditApproval.ListIndex = 2
        Case Else
            lbcCreditApproval.ListIndex = -1
    End Select
    edcRating.Text = Trim$(tmAgf.sCrdRtg)
    Select Case tmAgf.sShowISCI
        Case "Y"
            imISCI = 0
        Case "N"
            imISCI = 1
        Case "W"
            imISCI = 2
        Case Else
            imISCI = -1
    End Select
    lbcInvSort.ListIndex = 1
    smInvSort = ""
    slRecCode = Trim$(str$(tmAgf.iMnfSort))
    For ilLoop = 0 To UBound(tmInvSortCode) - 1 Step 1 'lbcInvSortCode.ListCount - 1 Step 1
        slNameCode = tmInvSortCode(ilLoop).sKey   'lbcInvSortCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcInvSort.ListIndex = ilLoop + 2
            smInvSort = lbcInvSort.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    Select Case tmAgf.sPkInvShow
        Case "D"
            imPackage = 0
        Case "T"
            imPackage = 1
        Case Else
            imPackage = -1
    End Select
    For ilLoop = 0 To 2 Step 1
        edcCAddr(ilLoop).Text = Trim$(tmAgf.sCntrAddr(ilLoop))
    Next ilLoop
    For ilLoop = 0 To 2 Step 1
        edcBAddr(ilLoop).Text = Trim$(tmAgf.sBillAddr(ilLoop))
    Next ilLoop
    'ReDim imNewPnfCode(1 To 1) As Integer
    ReDim imNewPnfCode(0 To 0) As Integer
    mBuyerPop tmAgf.iCode, "", -1
    lbcBuyer.ListIndex = 1
    smBuyer = ""
    smOrigBuyer = ""
    slRecCode = Trim$(str$(tmAgf.iPnfBuyer))
    For ilLoop = 0 To UBound(tmBuyerCode) - 1 Step 1 'lbcBuyerCode.ListCount - 1 Step 1
        slNameCode = tmBuyerCode(ilLoop).sKey  'lbcBuyerCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcBuyer.ListIndex = ilLoop + 2
            smBuyer = lbcBuyer.List(ilLoop + 2)
            smOrigBuyer = smBuyer
            Exit For
        End If
    Next ilLoop
    mPayablePop tmAgf.iCode, "", -1
    lbcPayable.ListIndex = 1
    smPayable = ""
    smOrigPayable = ""
    slRecCode = Trim$(str$(tmAgf.iPnfPay))
    For ilLoop = 0 To UBound(tmPayableCode) - 1 Step 1  'lbcPayableCode.ListCount - 1 Step 1
        slNameCode = tmPayableCode(ilLoop).sKey    'lbcPayableCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
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
    slRecCode = Trim$(str$(tmAgf.iArfLkCode))
    For ilLoop = 0 To UBound(tmLkBoxCode) - 1 Step 1 'lbcLkBoxCode.ListCount - 1 Step 1
        slNameCode = tmLkBoxCode(ilLoop).sKey  'lbcLkBoxCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcLkBox.ListIndex = ilLoop + 2
            smLkBox = lbcLkBox.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    lbcEDI(0).ListIndex = 0
    smEDIC = ""
    slRecCode = Trim$(str$(tmAgf.iArfCntrCode))
    For ilLoop = 0 To UBound(tmEDICode) - 1 Step 1  'lbcEDICode.ListCount - 1 Step 1
        slNameCode = tmEDICode(ilLoop).sKey   'lbcEDICode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcEDI(0).ListIndex = ilLoop + 1
            smEDIC = lbcEDI(0).List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    lbcEDI(1).ListIndex = 0
    smEDII = ""
    slRecCode = Trim$(str$(tmAgf.iArfInvCode))
    For ilLoop = 0 To UBound(tmEDICode) - 1 Step 1  'lbcEDICode.ListCount - 1 Step 1
        slNameCode = tmEDICode(ilLoop).sKey    'lbcEDICode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
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
    Select Case tmAgf.sCntrExptForm
        Case "C"
            imExport = 0
        Case "O"
            imExport = 1
        Case "X"
            imExport = 2
        Case Else
            imExport = -1
    End Select
'    Select Case tmAgf.sCntrPrtSz
'        Case "W"
'            imPrtStyle = 0
'        Case "N"
'            imPrtStyle = 1
'        Case Else
'            imPrtStyle = -1
'    End Select

    If tmAgfx.iInvFeatures = 1 Then ' TTP 10622 - 2023-03-08 JJB
        imSuppressNet = 1 'Suppress
    Else
        imSuppressNet = 0 'No Suppression
    End If
    
    edcXMLBand.Text = Trim$(tmAgf.sXMPProposalBand)
    edcXMLCall.Text = Trim$(tmAgf.sXMLCallLetters)
    Select Case tmAgf.sXMLDates
        Case "M"
            imXMLDates = 0
        Case "A"
            imXMLDates = 1
        Case Else
            imXMLDates = -1
    End Select
    lbcTerms.ListIndex = 1
    smTerms = lbcTerms.List(1)
    slRecCode = Trim$(str$(tmAgf.iMnfInvTerms))
    For ilLoop = 0 To UBound(tmTermsCode) - 1 Step 1 'lbcInvSortCode.ListCount - 1 Step 1
        slNameCode = tmTermsCode(ilLoop).sKey   'lbcInvSortCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Agency
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcTerms.ListIndex = ilLoop + 2
            smTerms = lbcTerms.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    '12/17/06-Change to tax by agency or vehicle
    'If (tmAgf.sSlsTax(0) = "N") And (tmAgf.sSlsTax(1) = "N") Then
    '    lbcTax.ListIndex = 0
    'ElseIf (tmAgf.sSlsTax(0) = "Y") And (tmAgf.sSlsTax(1) = "N") Then
    '    lbcTax.ListIndex = 1
    'ElseIf (tmAgf.sSlsTax(0) = "N") And (tmAgf.sSlsTax(1) = "Y") Then
    '    lbcTax.ListIndex = 2
    'ElseIf (tmAgf.sSlsTax(0) = "Y") And (tmAgf.sSlsTax(1) = "Y") Then
    '    lbcTax.ListIndex = 3
    'Else
    '    lbcTax.ListIndex = -1
    'End If
    'look up Tax from code number
    lbcTax.ListIndex = 0
    smTax = lbcTax.List(lbcTax.ListIndex)
    slRecCode = Trim$(str$(tmAgf.iTrfCode))
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

    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    'gPDNToStr tmAgf.sPct90, 0, slStr
    slStr = gIntToStrDec(tmAgf.iPct90, 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 0, smPct90
    gPDNToStr tmAgf.sCurrAR, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smCurrAR
    gPDNToStr tmAgf.sUnbilled, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smUnbilled
    gPDNToStr tmAgf.sHiCredit, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smHiCredit
    gPDNToStr tmAgf.sTotalGross, 2, slStr
    slStr = gRoundStr(slStr, "1.00", 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, smTotalGross
    gUnpackDate tmAgf.iDateEntrd(0), tmAgf.iDateEntrd(1), smDateEntrd
    smNSFChks = Trim$(str$(tmAgf.iNSFChks))
    gUnpackDate tmAgf.iDateLstInv(0), tmAgf.iDateLstInv(1), smDateLstInv
    gUnpackDate tmAgf.iDateLstPaym(0), tmAgf.iDateLstPaym(1), smDateLstPaym
    smAvgToPay = Trim$(str$(tmAgf.iAvgToPay))
    smLstToPay = Trim$(str$(tmAgf.iLstToPay))
    smNoInvPd = Trim$(str$(tmAgf.iNoInvPd))
    ' edcCRMID.Text = CStr(tmAgfx.lCrmId)
    
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
    If (Trim$(smName) <> "") And (Trim$(smCity) <> "") Then    'Test name
        slStr = Trim$(smName) & ", " & Trim$(smCity)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                slStr = Trim$(smName) & ", " & Trim$(smCity)
                If slStr = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Agency already defined, enter a different name or city ID", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    smName = Trim$(tmAgf.sName) 'Reset text
                    smCity = Trim$(tmAgf.sCityID)
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
    Dim ilSlfCode As Integer
    Dim ilSlf As Integer

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
    'gInitStdAlone Agency, slStr, ilTestSystem
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
    If (ilSlfCode > 0) Or (igAgyCallSource = CALLSOURCECONTRACT) Or (tgUrf(0).iRemoteUserID > 0) Then
        imShortForm = True
    Else
        imShortForm = False
    End If
    If imShortForm Then
        imPbcIndex = 1
    Else
        imPbcIndex = 0
    End If
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igAgyCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igAgyCallSource = CALLNONE
    'End If
    If igAgyCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgAgyName = slStr
        Else
            sgAgyName = ""
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
Private Sub mPayablePop(ilAgyCode As Integer, slRetainName As String, ilReturnCode As Integer)
'
'   mPayablePop
'   Where:
'       ilAgyCode (I)- Agency code value
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilPnf As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ilIndex = lbcPayable.ListIndex
    If ilIndex > 0 Then
        slName = lbcPayable.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'If imSelectedIndex > 0 Then 'Change mode
        'If imSelectedIndex = 0 Then
        '    ilRet = gPopPersonnelBox(Advt, 0, ilAdvtCode, "P", False, lbcPayable, lbcPayableCode)
        'Else
            'ilRet = gPopPersonnelBox(Agency, 1, ilAgyCode, "P", True, 2, lbcPayable, lbcPayableCode)
            ilRet = gPopPersonnelBox(Agency, 1, ilAgyCode, "P", True, 2, lbcPayable, tmPayableCode(), smPayableCodeTag)
        'End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPayablePopErr
            gCPErrorMsg ilRet, "mPayablePop (gPopPersonnelBox)", Agency
            On Error GoTo 0
            'Filter out any contact not associated with this agency
            If imSelectedIndex = 0 Then
                For ilLoop = UBound(tmPayableCode) - 1 To 0 Step -1 'lbcPayableCode.ListCount - 1 To 0 Step -1
                    ilFound = False
                    slNameCode = tmPayableCode(ilLoop).sKey    'lbcPayableCode.List(ilLoop)
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
                        gRemoveItemFromSortCode ilLoop, tmPayableCode()
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
    igPersonnelCallSource = CALLSOURCEAGENCY
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
                slNameCode = tmBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
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
                slNameCode = tmPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
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
        tmAgf.iCode = 0
    End If
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Agency^Test\" & sgUserName & "\" & Trim$(str$(igPersonnelCallSource)) & "\Agy" & "\" & Trim$(str$(tmAgf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & Trim$(str$(igPersonnelCallSource)) & "\Agy" & "\" & Trim$(str$(tmAgf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igPersonnelCallSource)) & "\Agy" & "\" & Trim$(Str$(tmAgf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igPersonnelCallSource)) & "\Agy" & "\" & Trim$(Str$(tmAgf.iCode)) & "\" & slBuyerOrPayable & "\" & sgPersonnelName
    '    End If
    'End If
    ilBoxNo = imBoxNo
    'lgShellRet = Shell(sgExePath & "Persnnel.Exe " & slStr, 1)
    'Agency.Enabled = False
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
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    imBoxNo = ilBoxNo
    'Screen.MousePointer = vbDefault    'Default
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
            smBuyerCodeTag = ""
            mBuyerPop tmAgf.iCode, sgPersonnelName, Val(slReturnCode)
            If imTerminate Then
                mPersonnelBranch = False
                Exit Function
            End If
            slName30 = sgPersonnelName  'Don't test phone Number
            'gFindPartialMatch slName30, 2, 30, lbcBuyer
            'sgPersonnelName = ""
            'If gLastFound(lbcBuyer) > 1 Then
            '    imChgMode = True
            '    lbcBuyer.ListIndex = gLastFound(lbcBuyer)
            '    edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
            '    imChgMode = False
            '    'slNameCode = tmBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            '    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    'imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '    'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '    If slNewFlag = "Y" Then
            '        slNameCode = tmBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '        ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '    End If
            '    mPersonnelBranch = False
            '    mSetChg BUYERINDEX
            'Else
            '    imChgMode = True
            '    lbcBuyer.ListIndex = 1
            '    edcDropDown.Text = lbcBuyer.List(1)
            '    imChgMode = False
            '    mSetChg BUYERINDEX
            '    'edcDropDown.SetFocus
            '    If edcDropDown.Visible Then
            '        edcDropDown.SetFocus
            '    Else
            '        pbcClickFocus.SetFocus
            '    End If
            '    Exit Function
            'End If
            ilFound = False
            For ilLoop = LBound(tmBuyerCode) To UBound(tmBuyerCode) - 1 Step 1
                slNameCode = tmBuyerCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = Val(slReturnCode) Then
                    ilFound = True
                    imChgMode = True
                    lbcBuyer.ListIndex = ilLoop + 2
                    edcDropDown.Text = lbcBuyer.List(lbcBuyer.ListIndex)
                    imChgMode = False
                    If slNewFlag = "Y" Then
                        slNameCode = tmBuyerCode(lbcBuyer.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
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
            smPayableCodeTag = ""
            mPayablePop tmAgf.iCode, sgPersonnelName, Val(slReturnCode)
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
            '    'slNameCode = tmPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            '    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    'imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '    'ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
            '    If slNewFlag = "Y" Then
            '        slNameCode = tmPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcPayableCode.List(lbcPayable.ListIndex - 2)
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        imNewPnfCode(UBound(imNewPnfCode)) = Val(slCode)
            '        ReDim Preserve imNewPnfCode(1 To UBound(imNewPnfCode) + 1) As Integer
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
            For ilLoop = LBound(tmPayableCode) To UBound(tmPayableCode) - 1 Step 1
                slNameCode = tmPayableCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = Val(slReturnCode) Then
                    ilFound = True
                    imChgMode = True
                    lbcPayable.ListIndex = ilLoop + 2
                    edcDropDown.Text = lbcPayable.List(lbcPayable.ListIndex)
                    imChgMode = False
                    If slNewFlag = "Y" Then
                        slNameCode = tmPayableCode(lbcPayable.ListIndex - 2).sKey  'lbcBuyerCode.List(lbcBuyer.ListIndex - 2)
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
    'ilRet = gPopAgyBoxNameCityBox(Agency, cbcSelect, Traffic!lbcAgency, True, lbcName, True, lbcCity)
    ilRet = gPopAgyBoxNameCityBox(Agency, cbcSelect, tgAgency(), sgAgencyTag, True, lbcName, True, lbcCity)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopAgyBox)", Agency
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
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

    slNameCode = tgAgency(ilSelectIndex - 1).sKey  'Traffic!lbcAgency.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", Agency
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmAgfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Agency
    On Error GoTo 0
    If tmAgf.iPnfBuyer > 0 Then
        tmPnfSrchKey.iCode = tmAgf.iPnfBuyer
        ilRet = btrGetEqual(hmPnf, tmBPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        'On Error GoTo mReadRecErr
        'gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Agency
        'On Error GoTo 0
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            If ilShowMissingMsg Then
                MsgBox "Buyer Name Missing", vbOKOnly + vbExclamation
            End If
            tmBPnf.iCode = 0
            tmAgf.iPnfBuyer = 0
        Else
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual- Buyer)", Agency
            On Error GoTo 0
        End If
    Else
        tmBPnf.iCode = 0
    End If
    If tmAgf.iPnfPay > 0 Then
        tmPnfSrchKey.iCode = tmAgf.iPnfPay
        ilRet = btrGetEqual(hmPnf, tmPPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        'On Error GoTo mReadRecErr
        'gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Agency
        'On Error GoTo 0
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            If ilShowMissingMsg Then
                MsgBox "Payables Contact Name Missing", vbOKOnly + vbExclamation
            End If
            tmPPnf.iCode = 0
            tmAgf.iPnfPay = 0
        Else
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual- Payables)", Agency
            On Error GoTo 0
        End If
    Else
        tmPPnf.iCode = 0
    End If
    
    'L.Bianchi '05/28/2021' start
    tmAgfx.iCode = tmAgf.iCode
    SQLQuery = "SELECT * FROM AGFX_Agencies WHERE agfxCode =" & tmAgfx.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        If Not IsNull(rs!agfxRefId) Then
            tmAgfx.sRefId = rs!agfxRefId
        End If
        
        If Not IsNull(rs!agfxCrmId) Then
            tmAgfx.lCrmId = rs!agfxCrmId
        End If
        
        If Not IsNull(rs!agfxInvFeatures) Then
            tmAgfx.iInvFeatures = rs!agfxInvFeatures
        Else
            tmAgfx.iInvFeatures = 0
        End If
    Else
        tmAgfx.sRefId = ""
        tmAgfx.lCrmId = 0
        tmAgfx.iInvFeatures = -1
    End If
    'L.Bianchi '05/28/2021' End
    
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
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
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim slStr As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlPnf As PNF
    Dim slEDII As String
    Dim rs As ADODB.Recordset
    Dim llCount As Long
    
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    
    If Val(edcCRMID.Text) > 2147483646 Then
        MsgBox "CRM ID cannot be larger than 2147483646", vbOKOnly + vbExclamation
        mSaveRec = False
        Exit Function
    End If
    If Len(edcCRMID.Text) < 1 Then
        tmAgfx.lCrmId = 0
    Else
        tmAgfx.lCrmId = CLng(edcCRMID.Text)
    End If

    tmAgfx.iInvFeatures = IIF(edcSuppressNet.Text = "Yes", 1, 0)  ' TTP 10622 - 2023-03-08 JJB
    
    'If mTestFields(REFIDINDEX, SHOWMSG) = NO Then
        'mSaveRec = False
        'Exit Function
    'End If
    
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    '9745
    If Not mXmlChoicesOk() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If (imSelectedIndex = 0) And (imPbcIndex = 1) Then 'Set defaults
        edcComm.Text = ""
        tmCtrls(COMMINDEX).iChg = True
        If (tgUrf(0).sChgCrRt <> "I") Then
            lbcCreditApproval.ListIndex = 0   'Requires Checking
        Else
            lbcCreditApproval.ListIndex = 1   'Approved
        End If
        tmCtrls(CREDITAPPROVALINDEX).iChg = True
        lbcCreditRestr.ListIndex = 0   'No Limit
        tmCtrls(CREDITRESTRINDEX).iChg = True
        lbcPaymRating.ListIndex = 1   'Normal
        tmCtrls(PAYMRATINGINDEX).iChg = True
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Agf.Btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE, False) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        If (imSelectedIndex = 0) And (imShortForm) Then 'New selected
            mMoveCtrlToRec False
        Else
            mMoveCtrlToRec True
        End If
        tmAgf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
        If imSelectedIndex = 0 Then 'New selected
            tmAgf.iCode = 0  'Autoincrement
            'slStr = ""
            'gStrToPDN slStr, 0, 2, tmAgf.sPct90
            tmAgf.iPct90 = 0
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAgf.sCurrAR
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAgf.sUnbilled
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAgf.sHiCredit
            slStr = ""
            gStrToPDN slStr, 2, 6, tmAgf.sTotalGross
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, tmAgf.iDateEntrd(0), tmAgf.iDateEntrd(1)
            tmAgf.iNSFChks = 0
            slStr = ""
            gPackDate slStr, tmAgf.iDateLstInv(0), tmAgf.iDateLstInv(1)
            slStr = ""
            gPackDate slStr, tmAgf.iDateLstPaym(0), tmAgf.iDateLstPaym(1)
            tmAgf.iAvgToPay = 0
            tmAgf.iLstToPay = 0
            tmAgf.iNoInvPd = 0
            tmAgf.iMerge = 0
            '10148 removed fields
           ' tmAgf.iRemoteID = tgUrf(0).iRemoteUserID
            'tmAgf.iAutoCode = tmAgf.iCode
            ilRet = btrInsert(hmAgf, tmAgf, imAgfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            '10148 removed fields
'            tmAgf.iSourceID = tgUrf(0).iRemoteUserID
'            gPackDate slSyncDate, tmAgf.iSyncDate(0), tmAgf.iSyncDate(1)
'            gPackTime slSyncTime, tmAgf.iSyncTime(0), tmAgf.iSyncTime(1)
            ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    
    'L.Bianchi 05/28/2021 start
    SQLQuery = "SELECT * FROM AGFX_Agencies WHERE agfxCode =" & tmAgf.iCode
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        SQLQuery = "UPDATE AGFX_Agencies SET "
        SQLQuery = SQLQuery & "     agfxRefId       = '" & Trim(tmAgfx.sRefId) & "', "
        SQLQuery = SQLQuery & "     agfxCrmId       = " & tmAgfx.lCrmId & ", "
        SQLQuery = SQLQuery & "     agfxInvFeatures = " & tmAgfx.iInvFeatures & " "
        SQLQuery = SQLQuery & " WHERE "
        SQLQuery = SQLQuery & "     agfxCode        = " & tmAgf.iCode
        
         If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", "Agency-mSaveRec"
        End If
    Else
        SQLQuery = "INSERT INTO AGFX_Agencies "
        SQLQuery = SQLQuery & "("
        SQLQuery = SQLQuery & " agfxCode,  "
        SQLQuery = SQLQuery & " agfxRefId, "
        SQLQuery = SQLQuery & " agfxCrmId, "
        SQLQuery = SQLQuery & " agfxInvFeatures  "
        SQLQuery = SQLQuery & ")"
        SQLQuery = SQLQuery & " VALUES "
        SQLQuery = SQLQuery & "("
        SQLQuery = SQLQuery & tmAgf.iCode & ","
        SQLQuery = SQLQuery & "'" & tmAgfx.sRefId & "',"
        SQLQuery = SQLQuery & tmAgfx.lCrmId & ","
        SQLQuery = SQLQuery & tmAgfx.iInvFeatures
        SQLQuery = SQLQuery & ")"
            
        If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", "Agency-mSaveRec"
        End If
    End If
    'L.Bianchi 05/28/2021 End
    
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Agency
    On Error GoTo 0
    If (imSelectedIndex = 0) Then   'And (tgSpf.sRemoteUsers = "Y") Then 'New selected
    '10148 removed DO
        'Do
            'tmAgfSrchKey.iCode = tmAgf.iCode
            'ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Agency)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, Agency
            'On Error GoTo 0
            'tmAgf.iRemoteID = tgUrf(0).iRemoteUserID
            'tmAgf.iAutoCode = tmAgf.iCode
            'tmAgf.iSourceID = tgUrf(0).iRemoteUserID
            'gPackDate slSyncDate, tmAgf.iSyncDate(0), tmAgf.iSyncDate(1)
            'gPackTime slSyncTime, tmAgf.iSyncTime(0), tmAgf.iSyncTime(1)
            'ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
            'slMsg = "mSaveRec (btrUpdate:Agency)"
        'Loop While ilRet = BTRV_ERR_CONFLICT
        'On Error GoTo mSaveRecErr
        'gBtrvErrorMsg ilRet, slMsg, Agency
        'On Error GoTo 0
        tgCommAgf(UBound(tgCommAgf)).iCode = tmAgf.iCode
        tgCommAgf(UBound(tgCommAgf)).sName = tmAgf.sName
        tgCommAgf(UBound(tgCommAgf)).sCityID = tmAgf.sCityID
        tgCommAgf(UBound(tgCommAgf)).sCreditRestr = tmAgf.sCreditRestr
        tgCommAgf(UBound(tgCommAgf)).iMnfSort = tmAgf.iMnfSort
        tgCommAgf(UBound(tgCommAgf)).sState = tmAgf.sState
        tgCommAgf(UBound(tgCommAgf)).iTrfCode = tmAgf.iTrfCode
        tgCommAgf(UBound(tgCommAgf)).s1or2DigitRating = tmAgf.s1or2DigitRating
        'ReDim Preserve tgCommAgf(1 To UBound(tgCommAgf) + 1) As AGFEXT
        ReDim Preserve tgCommAgf(0 To UBound(tgCommAgf) + 1) As AGFEXT
    Else
        ilRet = gBinarySearchAgf(tmAgf.iCode)
        If ilRet <> -1 Then
            tgCommAgf(ilRet).iCode = tmAgf.iCode
            tgCommAgf(ilRet).sName = tmAgf.sName
            tgCommAgf(ilRet).sCityID = tmAgf.sCityID
            tgCommAgf(ilRet).sCreditRestr = tmAgf.sCreditRestr
            tgCommAgf(ilRet).iMnfSort = tmAgf.iMnfSort
            tgCommAgf(ilRet).sState = tmAgf.sState
            tgCommAgf(ilRet).iTrfCode = tmAgf.iTrfCode
            tgCommAgf(ilRet).s1or2DigitRating = tmAgf.s1or2DigitRating
        End If
    End If
    'For ilLoop = 1 To UBound(imNewPnfCode) - 1 Step 1
    For ilLoop = 0 To UBound(imNewPnfCode) - 1 Step 1
        tmPnfSrchKey.iCode = imNewPnfCode(ilLoop)
        ilRet = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            Do
                tlPnf.iAgfCode = tmAgf.iCode
                tlPnf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tlPnf.iSyncDate(0), tlPnf.iSyncDate(1)
                gPackTime slSyncTime, tlPnf.iSyncTime(0), tlPnf.iSyncTime(1)
                ilRet = btrUpdate(hmPnf, tlPnf, imPnfRecLen)
                slMsg = "mSaveRec (btrUpdate:Personnel)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Agency
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
    ''If Traffic!lbcAgency.Tag <> "" Then
    ''    If slStamp = Traffic!lbcAgency.Tag Then
    ''        Traffic!lbcAgency.Tag = FileDateTime(sgDBPath & "Agf.Btr")
    ''    End If
    ''End If
    'If sgAgencyTag <> "" Then
    '    If slStamp = sgAgencyTag Then
    '        sgAgencyTag = gFileDateTime(sgDBPath & "Agf.Btr")
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    'Traffic!lbcAgency.RemoveItem imSelectedIndex - 1
    '    gRemoveItemFromSortCode imSelectedIndex - 1, tgAgency()
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityId)
    'cbcSelect.AddItem slName
    'Do While Len(slName) < Len(tmAgf.sName) + Len(tmAgf.sCityId) + 2
    '    slName = slName & " "
    'Loop
    'slName = slName + "\" + LTrim$(Str$(tmAgf.iCode))
    ''Traffic!lbcAgency.AddItem slName
    'gAddItemToSortCode slName, tgAgency(), True
    'slName = Trim$(tmAgf.sName)
    'gFindMatch slName, 0, lbcName
    'If gLastFound(lbcName) < 0 Then
    '    lbcName.AddItem slName
    'End If
    'slName = Trim$(tmAgf.sCityId)
    'gFindMatch slName, 0, lbcCity
    'If gLastFound(lbcCity) < 0 Then
    '    lbcCity.AddItem slName
    'End If
    'cbcSelect.AddItem "[New]", 0
    'mBuyerPop tmAgf.iCode, ""
    'mPayablePop tmAgf.iCode, ""
    imChgSaveFlag = True
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
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
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & Trim$(smName) & ", " & Trim$(smCity)
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcAgy_Paint imPbcIndex
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
    If ilBoxNo < imLBCtrls Or (ilBoxNo > UBound(tmCtrls)) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlagStr tmAgf.sName, smName, tmCtrls(ilBoxNo)
        Case ABBRINDEX 'Abbreviation
            gSetChgFlag tmAgf.sAbbr, edcAbbr, tmCtrls(ilBoxNo)
        Case CITYINDEX 'City
            gSetChgFlagStr tmAgf.sCityID, smCity, tmCtrls(ilBoxNo)
        Case STATEINDEX
        Case COMMINDEX  'Commission
            'gPDNToStr tmAgf.sComm, 2, slStr
            slStr = gIntToStrDec(tmAgf.iComm, 2)
            gSetChgFlag slStr, edcComm, tmCtrls(ilBoxNo)
        Case SPERSONINDEX   'Salesperson
            gSetChgFlag smSPerson, lbcSPerson, tmCtrls(ilBoxNo)
        Case DIGITRATINGINDEX
        Case REPCODEINDEX 'Rep Code
            gSetChgFlag tmAgf.sCodeRep, edcRepCode, tmCtrls(ilBoxNo)
        Case STNCODEINDEX 'Station Code
            gSetChgFlag tmAgf.sCodeStn, edcStnCode, tmCtrls(ilBoxNo)
        Case CREDITAPPROVALINDEX
            Select Case tmAgf.sCrdApp
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
        Case CREDITRESTRINDEX
            Select Case tmAgf.sCreditRestr
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
            'gPDNToStr tmAgf.sCreditLimit, 2, slStr
            slStr = gLongToStrDec(tmAgf.lCreditLimit, 2)
            gSetChgFlag slStr, edcCreditLimit, tmCtrls(ilBoxNo)
        Case PAYMRATINGINDEX
            Select Case tmAgf.sPaymRating
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
        Case CREDITRATINGINDEX
            gSetChgFlag tmAgf.sCrdRtg, edcRating, tmCtrls(ilBoxNo)
        Case ISCIINDEX
        Case INVSORTINDEX   'Invoice sorting
            gSetChgFlag smInvSort, lbcInvSort, tmCtrls(ilBoxNo)
        Case PACKAGEINDEX
        Case CADDRINDEX 'Contract address
            gSetChgFlag tmAgf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Contract address
            gSetChgFlag tmAgf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Contract address
            gSetChgFlag tmAgf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Billing address
            gSetChgFlag tmAgf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Billing address
            gSetChgFlag tmAgf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Billing address
            gSetChgFlag tmAgf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BUYERINDEX 'Buyer Name
            gSetChgFlag smOrigBuyer, lbcBuyer, tmCtrls(ilBoxNo)
        Case PAYABLEINDEX 'Buyer Name
            gSetChgFlag smOrigPayable, lbcPayable, tmCtrls(ilBoxNo)
        Case CRMIDINDEX
            If Len(edcCRMID.Text) > 0 And Not IsNumeric(edcCRMID.Text) Then
                edcCRMID.Text = ""  ' Don't allow anything but numbers
            End If
            If tmCtrls(ilBoxNo).iChg <> True Then
                slStr = ""
                If tmAgfx.lCrmId <> 0 Then
                    slStr = CStr(tmAgfx.lCrmId)
                End If
                gSetChgFlag slStr, edcCRMID, tmCtrls(ilBoxNo)
                If Len(edcCRMID.Text) > 0 And IsNumeric(edcCRMID.Text) Then
                    tmAgfx.lCrmId = CLng(edcCRMID.Text)
                Else
                    edcCRMID.Text = ""
                End If
            End If
        Case LKBOXINDEX   'Lock box
            gSetChgFlag smLkBox, lbcLkBox, tmCtrls(ilBoxNo)
        'L.Bianchi 04/15/2021
        Case REFIDINDEX   'EDI Service for contract
            gSetChgFlag tmAgfx.sRefId, edcRefId, tmCtrls(ilBoxNo)
        Case EDICINDEX   'EDI Service for contract
            gSetChgFlag smEDIC, lbcEDI(0), tmCtrls(ilBoxNo)
        Case EDIIINDEX   'EDI Service for Invoice
            gSetChgFlag smEDII, lbcEDI(1), tmCtrls(ilBoxNo)
            If bmPDFEMailChgd Then
                tmCtrls(ilBoxNo).iChg = True
            End If
'        Case PRTSTYLEINDEX
        Case TERMSINDEX   'Invoice sorting
            gSetChgFlag smTerms, lbcTerms, tmCtrls(ilBoxNo)
        Case TAXINDEX
            '12/17/06-Change to tax by agency or vehicle
            'If (tmAgf.sSlsTax(0) = "N") And (tmAgf.sSlsTax(1) = "N") Then
            '    slStr = lbcTax.List(0)
            'ElseIf (tmAgf.sSlsTax(0) = "Y") And (tmAgf.sSlsTax(1) = "N") Then
            '    slStr = lbcTax.List(1)
            'ElseIf (tmAgf.sSlsTax(0) = "N") And (tmAgf.sSlsTax(1) = "Y") Then
            '    slStr = lbcTax.List(2)
            'ElseIf (tmAgf.sSlsTax(0) = "Y") And (tmAgf.sSlsTax(1) = "Y") Then
            '    slStr = lbcTax.List(3)
            'Else
            '    slStr = ""
            'End If
            gSetChgFlag smTax, lbcTax, tmCtrls(ilBoxNo)
        Case XMLCALLINDEX
            gSetChgFlag tmAgf.sXMLCallLetters, edcXMLCall, tmCtrls(ilBoxNo)
        Case XMLBANDINDEX
            gSetChgFlag tmAgf.sXMPProposalBand, edcXMLBand, tmCtrls(ilBoxNo)
        Case SUPPRESSNETINDEX
            'gSetChgFlag tmAgfx.iInvFeatures, edcSuppressNet, tmCtrls(ilBoxNo)
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
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
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
    If Not ilAltered And (tgUrf(0).sMerge = "I") And (tgUrf(0).iRemoteID = 0) And (imUpdateAllowed) Then
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
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcDropDown.SetFocus
        Case ABBRINDEX 'Abbreviation
            edcAbbr.SetFocus
        Case CITYINDEX 'City
            edcDropDown.SetFocus
        Case STATEINDEX   'Active/Dormant
            pbcState.SetFocus
        Case COMMINDEX 'Commission
            edcComm.SetFocus
        Case SPERSONINDEX   'Salesperson
            edcDropDown.SetFocus
        Case DIGITRATINGINDEX   '1 or 2 digit rating
            pbcDigitRating.SetFocus
        Case REPCODEINDEX 'Rep Agency Code
            edcRepCode.SetFocus
        Case STNCODEINDEX 'Station Agency Code
            edcStnCode.SetFocus
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
        Case ISCIINDEX   'ISCI on Invoices
            pbcISCI.SetFocus
        Case INVSORTINDEX   'Invoice sorting
            edcDropDown.SetFocus
        Case PACKAGEINDEX   'Package Invoice Show
            pbcPackage.SetFocus
        Case CADDRINDEX 'Contract Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Billing Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BUYERINDEX 'Buyer
            edcDropDown.SetFocus
        Case PAYABLEINDEX 'Product
            edcDropDown.SetFocus
        Case CRMIDINDEX
            edcCRMID.SetFocus
        Case LKBOXINDEX 'Lock box
            edcDropDown.SetFocus
        Case EDICINDEX   'EDI service for Contracts
            edcDropDown.SetFocus
        Case EDIIINDEX   'EDI service for Contracts
            edcDropDown.SetFocus
        Case TERMSINDEX   'Terms
            edcDropDown.SetFocus
        Case TAXINDEX   'Tax
            edcDropDown.SetFocus
        Case EXPORTFORMINDEX   'Contract Export form
            pbcExport.SetFocus
'        Case PRTSTYLEINDEX   'Print style
'            pbcPrtStyle.SetFocus
        Case XMLCALLINDEX 'XML Proposal Band
            edcXMLCall.SetFocus
        Case XMLBANDINDEX 'XML Proposal Band
            edcXMLBand.SetFocus
        Case XMLDATESINDEX   'XML Dates
            pbcXMLDates.SetFocus
        Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
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
    Dim slStr As String
    Dim ilPos As Integer
    Dim slFirst As String
    Dim slLast As String
    Dim flWidth As Single

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    '2/4/16: Add filter to handle the case where the name has illegal characters and it was pasted into the field
    If (ilBoxNo = NAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcDropDown.Text)
        edcDropDown.Text = slStr
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            lbcName.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            gSetShow pbcAgy(imPbcIndex), smName, tmCtrls(ilBoxNo)
        Case ABBRINDEX 'Name
            edcAbbr.Visible = False  'Set visibility
            slStr = edcAbbr.Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CITYINDEX 'City
            lbcCity.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            gSetShow pbcAgy(imPbcIndex), smCity, tmCtrls(ilBoxNo)
        Case STATEINDEX   'Active/Dormant
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case COMMINDEX
            edcComm.Visible = False
            slStr = edcComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
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
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case DIGITRATINGINDEX   '! or 2 Digit Rating
            pbcDigitRating.Visible = False  'Set visibility
            If (tgSpf.sSGRPCPPCal = "A") Then
                If imDigitRating = 1 Then
                    slStr = "1"
                ElseIf imDigitRating = 2 Then
                    slStr = "2"
                Else
                    slStr = ""
                End If
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case REPCODEINDEX 'Rep Agency Code
            edcRepCode.Visible = False  'Set visibility
            If tgSpf.sARepCodes = "N" Then
                slStr = ""
            Else
                slStr = edcRepCode.Text
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case STNCODEINDEX 'Station Agency Code
            edcStnCode.Visible = False  'Set visibility
            If tgSpf.sAStnCodes = "N" Then
                slStr = ""
            Else
                slStr = edcStnCode.Text
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CREDITAPPROVALINDEX   'Credit Approval
            lbcCreditApproval.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcCreditApproval.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCreditApproval.List(lbcCreditApproval.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CREDITRESTRINDEX   'Credit Restriction
            lbcCreditRestr.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcCreditRestr.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCreditRestr.List(lbcCreditRestr.ListIndex)
            End If
            'gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            If lbcCreditRestr.ListIndex <> 1 Then
                flWidth = tmCtrls(ilBoxNo).fBoxW
                tmCtrls(ilBoxNo).fBoxW = 2 * tmCtrls(ilBoxNo).fBoxW
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
                tmCtrls(ilBoxNo).fBoxW = flWidth
            Else
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            End If
            If lbcCreditRestr.ListIndex <> 1 Then
                edcCreditLimit.Text = ""
                slStr = ""
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(CREDITRESTRINDEX + 1)
                gPaintArea pbcAgy(imPbcIndex), tmCtrls(CREDITRESTRINDEX + 1).fBoxX, tmCtrls(CREDITRESTRINDEX + 1).fBoxY, tmCtrls(CREDITRESTRINDEX + 1).fBoxW - 15, tmCtrls(CREDITRESTRINDEX + 1).fBoxH - 15, pbcAgy(imPbcIndex).BackColor     'WHITE
            End If
        Case CREDITRESTRINDEX + 1 'Credit Restriction
            edcCreditLimit.Visible = False
            If lbcCreditRestr.ListIndex = 1 Then
                slStr = edcCreditLimit.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case PAYMRATINGINDEX   'Payment rating
            lbcPaymRating.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcPaymRating.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcPaymRating.List(lbcPaymRating.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CREDITRATINGINDEX 'Name
            edcRating.Visible = False  'Set visibility
            slStr = edcRating.Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case ISCIINDEX   'ISCI on Invoices
            pbcISCI.Visible = False  'Set visibility
            If imISCI = 0 Then
                slStr = "Yes"
            ElseIf imISCI = 1 Then
                slStr = "No"
            ElseIf imISCI = 2 Then
                slStr = "W/O Leader"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case INVSORTINDEX   'Invoice sorting
            lbcInvSort.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcInvSort.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcInvSort.List(lbcInvSort.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case PACKAGEINDEX   'Package Invoice Show
            pbcPackage.Visible = False  'Set visibility
            If imPackage = 0 Then
                slStr = "Daypart"
            ElseIf imPackage = 1 Then
                slStr = "Time"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX 'Contract Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Contract Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case BUYERINDEX 'Buyer
            lbcBuyer.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcBuyer.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcBuyer.List(lbcBuyer.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case PAYABLEINDEX 'Buyer
            lbcPayable.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcPayable.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcPayable.List(lbcPayable.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            
        Case CRMIDINDEX
            edcCRMID.Visible = False
            tmCtrls(CRMIDINDEX).sShow = edcCRMID.Text
            
        Case LKBOXINDEX   'Lock box
            lbcLkBox.Visible = False  'Set visibility
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcLkBox.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcLkBox.List(lbcLkBox.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case REFIDINDEX 'L.Bianchi 04/15/2021
            edcRefId.Visible = False  'Set visibility
            slStr = edcRefId.Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case EDICINDEX   'EDI service for contract
            lbcEDI(0).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcEDI(0).ListIndex <= 0) Or (tgSpf.sAEDIC = "N") Then
                slStr = ""
            Else
                slStr = lbcEDI(0).List(lbcEDI(0).ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case EDIIINDEX   'EDI service for Invoices
            lbcEDI(1).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcEDI(1).ListIndex <= 0) Or (tgSpf.sAEDII = "N") Then
                slStr = ""
            Else
                slStr = lbcEDI(1).List(lbcEDI(1).ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case TERMSINDEX   'Invoice sorting
            lbcTerms.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTerms.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcTerms.List(lbcTerms.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case TAXINDEX   'Sales tax
            lbcTax.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcTax.ListIndex < 0) Or (Not imTaxDefined) Then
                slStr = ""
            Else
                slStr = lbcTax.List(lbcTax.ListIndex)
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            
        Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
             pbcSuppressNet.Visible = False
            If imSuppressNet = 0 Then
                slStr = "No"
            ElseIf imSuppressNet = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            edcSuppressNet.Text = ""
        Case EXPORTFORMINDEX   'Proposal export form
            pbcExport.Visible = False  'Set visibility
            If imExport = 0 Then
                slStr = "CSI"
            ElseIf imExport = 1 Then
                slStr = "OMD"
            ElseIf imExport = 2 Then
                slStr = "XML"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
            If imExport <> 2 Then
                edcXMLBand.Text = ""
                slStr = ""
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(XMLBANDINDEX)
                gPaintArea pbcAgy(imPbcIndex), tmCtrls(XMLBANDINDEX).fBoxX, tmCtrls(XMLBANDINDEX).fBoxY + fgBoxInsetY, tmCtrls(XMLBANDINDEX).fBoxW - 15, tmCtrls(XMLBANDINDEX).fBoxH - fgBoxInsetY - 15, pbcAgy(imPbcIndex).BackColor 'WHITE
                edcXMLCall.Text = ""
                slStr = ""
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(XMLCALLINDEX)
                gPaintArea pbcAgy(imPbcIndex), tmCtrls(XMLCALLINDEX).fBoxX, tmCtrls(XMLCALLINDEX).fBoxY + fgBoxInsetY, tmCtrls(XMLCALLINDEX).fBoxW - 15, tmCtrls(XMLCALLINDEX).fBoxH - fgBoxInsetY - 15, pbcAgy(imPbcIndex).BackColor 'WHITE
                imXMLDates = -1
                slStr = ""
                gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(XMLDATESINDEX)
                gPaintArea pbcAgy(imPbcIndex), tmCtrls(XMLDATESINDEX).fBoxX, tmCtrls(XMLDATESINDEX).fBoxY + fgBoxInsetY, tmCtrls(XMLDATESINDEX).fBoxW - 15, tmCtrls(XMLDATESINDEX).fBoxH - fgBoxInsetY - 15, pbcAgy(imPbcIndex).BackColor 'WHITE
            End If
'        Case PRTSTYLEINDEX   'Print style
'            pbcPrtStyle.Visible = False  'Set visibility
'            If imPrtStyle = 0 Then
'                slStr = "Wide"
'            ElseIf imPrtStyle = 1 Then
'                slStr = "Narrow"
'            Else
'                slStr = ""
'            End If
'            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case XMLCALLINDEX 'Name
            edcXMLCall.Visible = False  'Set visibility
            slStr = edcXMLCall.Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case XMLBANDINDEX 'Name
            edcXMLBand.Visible = False  'Set visibility
            slStr = edcXMLBand.Text
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case XMLDATESINDEX   'XML Dates
            pbcXMLDates.Visible = False  'Set visibility
            If imXMLDates = 0 Then
                slStr = "M-Su"
            ElseIf imXMLDates = 1 Then
                slStr = "Air"
            Else
                slStr = ""
            End If
            gSetShow pbcAgy(imPbcIndex), slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonBranch                  *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      office and process             *
'*                      communication back from sales  *
'*                      office                         *
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
'   ilRet = mSSourceBranch()
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
    'If (imDoubleClickName) And (lbcSPerson.ListIndex > Traffic!lbcSalesperson.ListCount) Then
    If (imDoubleClickName) And (lbcSPerson.ListIndex > UBound(tgSalesperson)) Then
        imCombo = True
        'igVsfCallSource = CALLSOURCEAGENCY
        'sgVsfName = slStr
        'sgVsfCallType = "S"
        'mSPersonBranch = True
'                Combo.Show vbModal
    Else
        'If Not gWinRoom(igNoLJWinRes(SALESPEOPLELIST)) Then
        '    imDoubleClickName = False
        '    mSPersonBranch = True
        '    mEnableBox imBoxNo
        '    Exit Function
        'End If
        imCombo = False
        igSlfCallSource = CALLSOURCEAGENCY
        If edcDropDown.Text = "[New]" Then
            sgSlfName = ""
        Else
            sgSlfName = slStr
        End If
        ilUpdateAllowed = imUpdateAllowed

        'igChildDone = False
        'edcLinkSrceDoneMsg.Text = ""
        'If (Not igStdAloneMode) And (imShowHelpMsg) Then
            If igTestSystem Then
                slStr = "Agency^Test\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
            Else
                slStr = "Agency^Prod\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
            End If
        'Else
        '    If igTestSystem Then
        '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
        '    Else
        '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
        '    End If
        'End If
        'lgShellRet = Shell(sgExePath & "SPerson.Exe " & slStr, 1)
        'Agency.Enabled = False
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        sgCommandStr = slStr
        SPerson.Show vbModal
        slStr = sgDoneMsg
        ilParse = gParseItem(slStr, 1, "\", sgSlfName)
        igSlfCallSource = Val(sgSlfName)
        ilParse = gParseItem(slStr, 2, "\", sgSlfName)
        'Agency.Enabled = True
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
'        gSetMenuState True
        lbcSPerson.Clear
        sgSalespersonTag = ""
        sgMSlfStamp = ""
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
    Dim ilDormant As Integer
    ilIndex = lbcSPerson.ListIndex
    If ilIndex > 1 Then
        slName = lbcSPerson.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSPersonComboBox(Agency, lbcSPerson, Traffic!lbcSalesperson, Traffic!lbcSPersonCombo, igSlfFirstNameFirst)
    'ilRet = gPopSalespersonBox(Agency, 0, False, True, lbcSPerson, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    If imSelectedIndex = 0 Then 'New selected
        ilDormant = False
    Else
        ilDormant = True
    End If
    ilRet = gPopSalespersonBox(Agency, 0, False, ilDormant, lbcSPerson, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gIMoveListBox)", Agency
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

    sgDoneMsg = Trim$(str$(igAgyCallSource)) & "\" & sgAgyName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload Agency
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
    If igWinStatus(AGENCIESLIST) <> 2 Then
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
    igMNmCallSource = CALLSOURCEAGENCY
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
            slStr = "Agency^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
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
    'ilRet = gIMoveListBox(Agency, lbcTerms, lbcTermsCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Agency, lbcTerms, tmTermsCode(), smTermsCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTermsPopErr
        gCPErrorMsg ilRet, "mTermsPop (gIMoveListBox)", Agency
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
        If gFieldDefinedStr(smName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
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
    If (ilCtrlNo = CITYINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smCity, "", "City ID must be specified", tmCtrls(CITYINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CITYINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
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
    If (ilCtrlNo = COMMINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcComm, "", "Agency commission must be specified", tmCtrls(COMMINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMINDEX
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
    If (ilCtrlNo = DIGITRATINGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imDigitRating = 1 Then
            slStr = "1"
        ElseIf imDigitRating = 2 Then
            slStr = "2"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "1 or 2 Digit Rating must be specified", tmCtrls(DIGITRATINGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DIGITRATINGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REPCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRepCode, "", "Rep Agency Code must be specified", tmCtrls(REPCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REPCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    slStr = Trim$(edcRepCode.Text)
    If Not gGPNoOk(slStr, 0, tmAgf.iCode, hmAdf, hmAgf) Then
        If (ilState And SHOWMSG) = SHOWMSG Then
            ilRet = MsgBox("Rep Agency Code must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
    End If
    If (ilCtrlNo = STNCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcStnCode, "", "Station Agency Code must be specified", tmCtrls(STNCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STNCODEINDEX
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
        If gFieldDefinedCtrl(edcAbbr, "", "Credit Rating must be specified", tmCtrls(CREDITRATINGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CREDITRATINGINDEX
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
    If (ilCtrlNo = EXPORTFORMINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imExport = 0 Then
            slStr = "CSI"
        ElseIf imExport = 1 Then
            slStr = "OMD"
        ElseIf imExport = 2 Then
            slStr = "XML"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Contract Export Form must be specified", tmCtrls(EXPORTFORMINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = EXPORTFORMINDEX
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
            'If ilState = ilState = (ALLMANDEFINED + SHOWMSG) Or ilState = SHOWMSG Then
                'imBoxNo = REFIDINDEX
            'End If
            'mTestFields = NO
            'Exit Function
        'End If
        
         'If gFieldDefinedGuidStr(edcRefId.Text, "", "Valid Ref Id identifier must be specified", ilState) = NO Then
            'If ilState = ilState = (ALLMANDEFINED + SHOWMSG) Or ilState = SHOWMSG Then
                'imBoxNo = REFIDINDEX
            'End If
            'mTestFields = NO
            'Exit Function
       ' End If
        
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

Private Sub pbcAgy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim flAdj As Single
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (ilBox = CADDRINDEX + 1) Or (ilBox = CADDRINDEX + 2) Or (ilBox = BADDRINDEX + 1) Or (ilBox = BADDRINDEX + 2) Then
                flAdj = fgBoxInsetY
            Else
                flAdj = 0
            End If
            If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                If (ilBox = DIGITRATINGINDEX) And (tgSpf.sSGRPCPPCal <> "A") Then
                    imDigitRating = 1
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = REPCODEINDEX) And (tgSpf.sARepCodes = "N") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = STNCODEINDEX) And (tgSpf.sAStnCodes = "N") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = ISCIINDEX) And (tgSpf.sAISCI = "A") Then
                    imISCI = 0
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = ISCIINDEX) And (tgSpf.sAISCI = "X") Then
                    imISCI = 1
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = EDICINDEX) And (tgSpf.sAEDIC = "N") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = EDIIINDEX) And (tgSpf.sAEDII = "N") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
'                If (ilBox = PRTSTYLEINDEX) And (tgSpf.sAPrtStyle = "W") Then
'                    imPrtStyle = 0
'                    mSetFocus imBoxNo
'                    Beep
'                    Exit Sub
'                End If
'                If (ilBox = PRTSTYLEINDEX) And (tgSpf.sAPrtStyle = "N") Then
'                    imPrtStyle = 1
'                    mSetFocus imBoxNo
'                    Beep
'                    Exit Sub
'                End If
                If (ilBox = TAXINDEX) And (Not imTaxDefined) Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX) And (tgUrf(0).sCredit <> "I") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX + 1) And (tgUrf(0).sCredit <> "I") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = PAYMRATINGINDEX) And (tgUrf(0).sPayRate <> "I") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITAPPROVALINDEX) And (tgUrf(0).sChgCrRt <> "I") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = PACKAGEINDEX) And (tgSpf.sCPkOrdered = "N") And (tgSpf.sCPkAired = "N") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = XMLCALLINDEX) And (imExport <> 2) Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = XMLBANDINDEX) And (imExport <> 2) Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = XMLDATESINDEX) And (imExport <> 2) Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = CREDITRESTRINDEX + 1) Then
                    ilBox = CREDITRESTRINDEX
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcAgy_Paint(Index As Integer)
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintAgyTitle (Index)
    If (tgSpf.sSGRPCPPCal = "A") And (Index = 0) Then
        llColor = pbcAgy(Index).ForeColor
        slFontName = pbcAgy(Index).FontName
        flFontSize = pbcAgy(Index).FontSize
        pbcAgy(Index).ForeColor = BLUE
        pbcAgy(Index).FontBold = True   'False
        pbcAgy(Index).FontSize = 6  '7
        pbcAgy(Index).FontName = "Arial"    '"Lucida Sans"  '"Arial"
        pbcAgy(Index).FontSize = 6  '7  'Font size done twice as indicated in FontSize property area in manual
        pbcAgy(Index).CurrentX = tmCtrls(DIGITRATINGINDEX).fBoxX + 15  'fgBoxInsetX
        pbcAgy(Index).CurrentY = tmCtrls(DIGITRATINGINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcAgy(Index).Print "1 or 2 Place Rating"
        pbcAgy(Index).FontSize = flFontSize
        pbcAgy(Index).FontName = slFontName
        pbcAgy(Index).FontSize = flFontSize
        pbcAgy(Index).ForeColor = llColor
        pbcAgy(Index).FontBold = True
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (ilBox <> CREDITAPPROVALINDEX) Or (tgUrf(0).sChgCrRt <> "H") Then
            If (ilBox <> PAYMRATINGINDEX) Or (tgUrf(0).sPayRate <> "H") Then
                If (ilBox <> CREDITRESTRINDEX) Or (tgUrf(0).sCredit <> "H") Then
                    If (ilBox <> CREDITRESTRINDEX + 1) Or (tgUrf(0).sCredit <> "H") Then
                        pbcAgy(Index).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                        pbcAgy(Index).CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
                        pbcAgy(Index).Print tmCtrls(ilBox).sShow
                    End If
                End If
            End If
        End If
    Next ilBox
'    llColor = pbcAgy(Index).ForeColor
'    pbcAgy(Index).ForeColor = BLUE
    'pbcAgy(Index).CurrentX = 30 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smPct90
    'pbcAgy(Index).CurrentX = 1560 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smCurrAR
    'pbcAgy(Index).CurrentX = 3090 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smUnbilled
    'pbcAgy(Index).CurrentX = 4620 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smHiCredit
    'pbcAgy(Index).CurrentX = 6150 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smTotalGross
    'pbcAgy(Index).CurrentX = 7440 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 2820 + fgBoxInsetY
    'pbcAgy(Index).Print smDateEntrd
    'pbcAgy(Index).CurrentX = 30 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
    'pbcAgy(Index).Print smNSFChks
    'pbcAgy(Index).CurrentX = 1560 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
    'pbcAgy(Index).Print smDateLstInv
    'pbcAgy(Index).CurrentX = 3090 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
    'pbcAgy(Index).Print smDateLstPaym
    'pbcAgy(Index).CurrentX = 4620 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
    'pbcAgy(Index).Print smAvgToPay
    'pbcAgy(Index).CurrentX = 6150 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
    'pbcAgy(Index).Print smLstToPay
    'pbcAgy(Index).CurrentX = 7440 + fgBoxInsetX
    'pbcAgy(Index).CurrentY = 3160 + fgBoxInsetY
'    pbcAgy(Index).ForeColor = llColor
    pbcAgy(Index).CurrentX = tmARCtrls(PCT90INDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(PCT90INDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smPct90
    pbcAgy(Index).CurrentX = tmARCtrls(CURRARINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(CURRARINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smCurrAR
    pbcAgy(Index).CurrentX = tmARCtrls(UNBILLEDINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(UNBILLEDINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smUnbilled
    pbcAgy(Index).CurrentX = tmARCtrls(HICREDITINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(HICREDITINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smHiCredit
    pbcAgy(Index).CurrentX = tmARCtrls(TOTALGROSSINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(TOTALGROSSINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smTotalGross
    pbcAgy(Index).CurrentX = tmARCtrls(DATEENTRDINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(DATEENTRDINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smDateEntrd
    pbcAgy(Index).CurrentX = tmARCtrls(NSFCHKSINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(NSFCHKSINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smNSFChks
    pbcAgy(Index).CurrentX = tmARCtrls(DATELSTINVINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(DATELSTINVINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smDateLstInv
    pbcAgy(Index).CurrentX = tmARCtrls(DATELSTPAYMINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(DATELSTPAYMINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smDateLstPaym
    pbcAgy(Index).CurrentX = tmARCtrls(AVGTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(AVGTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smAvgToPay
    pbcAgy(Index).CurrentX = tmARCtrls(LSTTOPAYINDEX).fBoxX + fgBoxInsetX
    pbcAgy(Index).CurrentY = tmARCtrls(LSTTOPAYINDEX).fBoxY + fgBoxInsetY
    pbcAgy(Index).Print smLstToPay

End Sub
Private Sub pbcClickFocus_GotFocus()
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

Private Sub pbcDigitRating_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcDigitRating_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("1") Then
        If imDigitRating <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imDigitRating = 1
        pbcDigitRating_Paint
    ElseIf KeyAscii = Asc("2") Then
        If imDigitRating <> 2 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imDigitRating = 2
        pbcDigitRating_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imDigitRating = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imDigitRating = 2
            pbcDigitRating_Paint
        ElseIf imDigitRating = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imDigitRating = 1
            pbcDigitRating_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcDigitRating_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDigitRating = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imDigitRating = 2
    ElseIf imDigitRating = 2 Then
        tmCtrls(imBoxNo).iChg = True
        imDigitRating = 1
    End If
    pbcDigitRating_Paint
    mSetCommands
End Sub

Private Sub pbcDigitRating_Paint()
    pbcDigitRating.Cls
    pbcDigitRating.CurrentX = fgBoxInsetX
    pbcDigitRating.CurrentY = 0 'fgBoxInsetY
    If imDigitRating = 1 Then
        pbcDigitRating.Print "1"
    ElseIf imDigitRating = 2 Then
        pbcDigitRating.Print "2"
    Else
        pbcDigitRating.Print "   "
    End If
End Sub

Private Sub pbcExport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub pbcExport_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
        If imExport <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imExport = 0
        pbcExport_Paint
    ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imExport <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imExport = 1
        pbcExport_Paint
    ElseIf KeyAscii = Asc("X") Or (KeyAscii = Asc("x")) Then
        If (Asc(tgSpf.sUsingFeatures9) And PROPOSALXML) = PROPOSALXML Then
            If imExport <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imExport = 2
            pbcExport_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imExport = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imExport = 1
            pbcExport_Paint
        ElseIf imExport = 1 Then
            tmCtrls(imBoxNo).iChg = True
            If (Asc(tgSpf.sUsingFeatures9) And PROPOSALXML) = PROPOSALXML Then
                imExport = 2
            Else
                imExport = 0
            End If
            pbcExport_Paint
        ElseIf imExport = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imExport = 0
            pbcExport_Paint
        End If
    End If
    mSetCommands
End Sub


Private Sub pbcSuppressNet_GotFocus()
' TTP 10622 - 2023-03-08 JJB
    gCtrlGotFocus ActiveControl
    If imSuppressNet = -1 Then
        imSuppressNet = 0
        pbcSuppressNet_Paint
    End If
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

Private Sub pbcExport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imExport = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imExport = 1
    ElseIf imExport = 1 Then
        tmCtrls(imBoxNo).iChg = True
        If (Asc(tgSpf.sUsingFeatures9) And PROPOSALXML) = PROPOSALXML Then
            imExport = 2
        Else
            imExport = 0
        End If
    ElseIf imExport = 2 Then
        tmCtrls(imBoxNo).iChg = True
        imExport = 0
    End If
    pbcExport_Paint
    mSetCommands

End Sub

Private Sub pbcExport_Paint()
    pbcExport.Cls
    pbcExport.CurrentX = fgBoxInsetX
    pbcExport.CurrentY = 0 'fgBoxInsetY
    If imExport = 0 Then
        pbcExport.Print "CSI"
    ElseIf imExport = 1 Then
        pbcExport.Print "OMD"
    ElseIf imExport = 2 Then
        pbcExport.Print "XML"
    Else
        pbcExport.Print "   "
    End If
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
    ElseIf KeyAscii = Asc("W") Or (KeyAscii = Asc("w")) Then     '3-10-15 addl flag to show isci but remove hard-coded WW_ (W/O Leader)
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
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
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
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
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
            Case -1
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
                ilBox = 1
            Case REPCODEINDEX
                If (tgSpf.sSGRPCPPCal <> "A") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = DIGITRATINGINDEX
            Case STNCODEINDEX
                If (tgSpf.sARepCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = REPCODEINDEX
            Case CREDITAPPROVALINDEX
                If (tgSpf.sAStnCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = STNCODEINDEX
            Case CREDITRESTRINDEX
                If tgUrf(0).sChgCrRt <> "I" Then
                    If tgSpf.sAStnCodes = "N" Then
                        ilFound = False
                    End If
                    ilBox = STNCODEINDEX
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
            Case INVSORTINDEX
                If tgSpf.sAISCI = "A" Then
                    If imISCI <> 0 Then
                        imISCI = 0
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                ElseIf tgSpf.sAISCI = "X" Then
                    If imISCI <> 1 Then
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
            Case PACKAGEINDEX
                If (lbcInvSort.ListCount = 2) Or imShortForm Then
                    imChgMode = True
                    lbcInvSort.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = INVSORTINDEX
            Case CADDRINDEX
                If imShortForm Or ((tgSpf.sCPkOrdered = "N") And (tgSpf.sCPkAired = "N")) Then
                    If (lbcInvSort.ListCount = 2) Or imShortForm Then
                        imChgMode = True
                        lbcInvSort.ListIndex = 1
                        imChgMode = False
                        ilFound = False
                    End If
                    ilBox = INVSORTINDEX
                Else
                    ilBox = PACKAGEINDEX
                End If
            Case BADDRINDEX
                If edcCAddr(2).Text <> "" Then
                    ilBox = CADDRINDEX + 2
                ElseIf edcCAddr(1).Text <> "" Then
                    ilBox = CADDRINDEX + 1
                Else
                    ilBox = CADDRINDEX
                End If
            Case BUYERINDEX
                If edcBAddr(2).Text <> "" Then
                    ilBox = BADDRINDEX + 2
                ElseIf edcBAddr(1).Text <> "" Then
                    ilBox = BADDRINDEX + 1
                Else
                    ilBox = BADDRINDEX
                End If
            Case CRMIDINDEX
                ilBox = PAYABLEINDEX
                
            Case REFIDINDEX 'L.Bianchi 06/02/2021
                ilBox = CRMIDINDEX
'                If lbcLkBox.ListCount <= 2 Then
'                    ilBox = PAYABLEINDEX
'                Else
'                    ilBox = LKBOXINDEX 'L.Bianchi 06/02/2021
'                End If
'                If imShortForm Then
'                    ilFound = False
'                End If
            Case EDICINDEX
                ilBox = REFIDINDEX 'L.Bianchi 06/16/2021
                If imShortForm Then
                    ilFound = False
                End If
            Case EDIIINDEX
                If (tgSpf.sAEDIC = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDICINDEX
            Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
'                If (Not imTaxDefined) Or imShortForm Then
'                    ilFound = False
'                End If
                 ilBox = EXPORTFORMINDEX
            Case EXPORTFORMINDEX
                'If (tgSpf.sAEDII = "N") Or imShortForm Then
                '    ilFound = False
                'End If
                'ilBox = EDIIINDEX
                If (Not imTaxDefined) Or imShortForm Then
                    ilFound = False
                End If
                ilBox = TAXINDEX
            Case TERMSINDEX
                If (tgSpf.sAEDII = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = EDIIINDEX
                
            Case TAXINDEX
                ilBox = TERMSINDEX
            Case XMLCALLINDEX
                If imShortForm Then
                    If imExport < 0 Then
                        tmCtrls(EXPORTFORMINDEX).iChg = True
                        imExport = 0
                    End If
                    ilFound = False
                End If
                ilBox = EXPORTFORMINDEX
            Case XMLBANDINDEX
                If imShortForm Then
                    If imExport < 0 Then
                        tmCtrls(EXPORTFORMINDEX).iChg = True
                        imExport = 0
                    End If
                    ilFound = False
                End If
                ilBox = XMLCALLINDEX    'EXPORTFORMINDEX
            Case Else
                If imShortForm Then
                    If (ilBox >= SPERSONINDEX) And (ilBox <= PACKAGEINDEX) Then
                        ilFound = False
                    End If
                    If (ilBox >= LKBOXINDEX) And (ilBox <= XMLDATESINDEX) Then  'TAXINDEX) Then
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
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
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
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
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
            Case -1
                imTabDirection = -1  'Set-Right to left
                If imShortForm Then
                    ilBox = BUYERINDEX
                Else
                    'If Not imTaxDefined Then
                    '    imBoxNo = TAXINDEX
                    '    pbcSTab.SetFocus
                    '    Exit Sub
                    'End If
                    'ilBox = TAXINDEX
                    ilBox = EXPORTFORMINDEX
                End If
            Case COMMINDEX
                If imShortForm Then
                    If edcComm.Text = "" Then
                        edcComm.Text = 15
                    End If
                    ilFound = False
                End If
                ilBox = SPERSONINDEX
            Case SPERSONINDEX
                If (tgSpf.sSGRPCPPCal <> "A") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = DIGITRATINGINDEX
            Case DIGITRATINGINDEX
                If (tgSpf.sARepCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = REPCODEINDEX
            Case REPCODEINDEX
                If (tgSpf.sAStnCodes = "N") Or imShortForm Then
                    ilFound = False
                End If
                ilBox = STNCODEINDEX
            Case STNCODEINDEX
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
            Case CREDITRATINGINDEX
                If tgSpf.sAISCI = "A" Then
                    If imISCI <> 0 Then
                        imISCI = 0
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                ElseIf tgSpf.sAISCI = "X" Then
                    If imISCI <> 1 Then
                        imISCI = 1
                        tmCtrls(ISCIINDEX).iChg = True
                    End If
                    ilFound = False
                End If
                If imShortForm Then
                    ilFound = False
                    If imISCI < 0 Then
                        tmCtrls(ISCIINDEX).iChg = True
                        If tgSpf.sAISCI = "Y" Then
                            imISCI = 0
                        Else
                            imISCI = 1
                        End If
                    End If
                End If
                ilBox = ISCIINDEX
            Case ISCIINDEX
                If (lbcInvSort.ListCount = 2) Or imShortForm Then
                    imChgMode = True
                    lbcInvSort.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = INVSORTINDEX
            Case INVSORTINDEX
                If (imShortForm) Or ((tgSpf.sCPkOrdered = "N") And (tgSpf.sCPkAired = "N")) Then
                    ilBox = CADDRINDEX
                Else
                    ilBox = PACKAGEINDEX
                End If
            Case BADDRINDEX
                If edcBAddr(0).Text = "" Then
                    If edcBAddr(1).Text = "" Then
                        If edcBAddr(2).Text = "" Then
                            mSetShow imBoxNo
                            ilBox = BUYERINDEX
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
                        ilBox = BUYERINDEX
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
                ilBox = BADDRINDEX + 2
            Case PAYABLEINDEX
                If imShortForm Then
                    ilFound = False
                    ilBox = LKBOXINDEX
                Else
                    ilBox = CRMIDINDEX
                End If
            Case CRMIDINDEX
                ilBox = REFIDINDEX
                
            Case LKBOXINDEX
                'L.Bianchi 06/16/2021
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = REFIDINDEX 'EDICINDEX
            Case REFIDINDEX 'L.Bianchi 04/15/2021
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
                If imShortForm Then
                    If imExport < 0 Then
                        tmCtrls(EXPORTFORMINDEX).iChg = True
                        imExport = 0
                    End If
                    ilFound = False
                End If
                ilBox = EXPORTFORMINDEX
            
            Case TERMSINDEX
                If (Not imTaxDefined) Or imShortForm Then
                    ilFound = False
                End If
                ilBox = TAXINDEX
            Case EXPORTFORMINDEX
                If imShortForm Then
                    ilFound = False
                End If
                If imExport <> 2 Then
                    'mSetShow imBoxNo
                    ilBox = SUPPRESSNETINDEX
                Else
                    ilBox = XMLCALLINDEX
                End If
            Case XMLCALLINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = XMLBANDINDEX
            Case XMLBANDINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = XMLDATESINDEX
            Case XMLDATESINDEX  'TAXINDEX
                If imShortForm Then
                    ilFound = False
                End If
                ilBox = SUPPRESSNETINDEX
            Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igAgyCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case Else
                If imShortForm Then
                    If (ilBox >= STATEINDEX) And (ilBox <= PACKAGEINDEX) Then
                        ilFound = False
                    End If
                    If (ilBox >= BUYERINDEX) And (ilBox <= XMLDATESINDEX) Then  'TAXINDEX) Then
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

Private Sub pbcXMLDates_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcXMLDates_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("M") Or (KeyAscii = Asc("m")) Then
        If imXMLDates <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imXMLDates = 0
        pbcXMLDates_Paint
    ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imXMLDates <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imXMLDates = 1
        pbcXMLDates_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imXMLDates = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imXMLDates = 1
            pbcXMLDates_Paint
        ElseIf imXMLDates = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imXMLDates = 0
            pbcXMLDates_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcXMLDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imXMLDates = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imXMLDates = 1
    ElseIf imXMLDates = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imXMLDates = 0
    End If
    pbcXMLDates_Paint
    mSetCommands
End Sub

Private Sub pbcXMLDates_Paint()
    pbcXMLDates.Cls
    pbcXMLDates.CurrentX = fgBoxInsetX
    pbcXMLDates.CurrentY = 0 'fgBoxInsetY
    If imXMLDates = 0 Then
        pbcXMLDates.Print "M-Su"
    ElseIf imXMLDates = 1 Then
        pbcXMLDates.Print "Air"
    Else
        pbcXMLDates.Print "   "
    End If
End Sub

Private Sub plcBkgd_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case SPERSONINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSPerson, edcDropDown, imChgMode, imLbcArrowSetting
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
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Agency"
End Sub

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
    tmPdfSrchKey2.iCode = tmAgf.iCode
    ilRet = btrGetEqual(hmPDF, tmPdf(ilRow), imPdfRecLen, tmPdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE)
        If tmAgf.iCode <> tmPdf(ilRow).iAgfCode Then
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
    
    tmPdfSrchKey2.iCode = tmAgf.iCode
    ilRet = btrGetEqual(hmPDF, tlPdf, imPdfRecLen, tmPdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE)
        ilRet = btrDelete(hmPDF)
        tmPdfSrchKey2.iCode = tmAgf.iCode
        ilRet = btrGetEqual(hmPDF, tlPdf, imPdfRecLen, tmPdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
End Function

Private Function mAddPDFEMail() As Integer
    Dim ilRet As Integer
    Dim ilPdf As Integer
    For ilPdf = 0 To UBound(tmPdf) - 1 Step 1
        tmPdf(ilPdf).lCode = 0
        tmPdf(ilPdf).iAdfCode = 0
        tmPdf(ilPdf).iAgfCode = tmAgf.iCode
        tmPdf(ilPdf).sUnused = ""
        ilRet = btrInsert(hmPDF, tmPdf(ilPdf), imPdfRecLen, INDEXKEY0)
    Next ilPdf
End Function
Private Sub mPaintAgyTitle(ilIndex As Integer)
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcAgy(ilIndex).ForeColor
    slFontName = pbcAgy(ilIndex).FontName
    flFontSize = pbcAgy(ilIndex).FontSize
    pbcAgy(ilIndex).ForeColor = BLUE
    pbcAgy(ilIndex).FontBold = False
    pbcAgy(ilIndex).FontSize = 7
    pbcAgy(ilIndex).FontName = "Arial"
    pbcAgy(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    ''For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    For ilLoop = NAMEINDEX To PACKAGEINDEX Step 1
        If ilLoop <> CREDITRESTRINDEX + 1 Then
            If ilLoop = CREDITRESTRINDEX Then
                pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + tmCtrls(ilLoop + 1).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            Else
                pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            End If
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            Select Case ilLoop
                Case NAMEINDEX
                    pbcAgy(ilIndex).Print "Name"
                Case ABBRINDEX
                    pbcAgy(ilIndex).Print "Abbreviation"
                Case CITYINDEX
                    pbcAgy(ilIndex).Print "City ID"
                Case STATEINDEX
                    pbcAgy(ilIndex).Print "Active/Dormant"
                Case COMMINDEX
                    pbcAgy(ilIndex).Print "% Commission"
                    If ilIndex = 1 Then
                        pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - (pbcAgy(imPbcIndex).height + 100) - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                    End If
                Case SPERSONINDEX
                    pbcAgy(ilIndex).Print "default Salesperson"
                Case DIGITRATINGINDEX
                    'pbcAgy(ilIndex).Print "1 or 2 Place Rating"
                Case REPCODEINDEX
                    pbcAgy(ilIndex).Print "G/L Agency Code"
                Case STNCODEINDEX
                    pbcAgy(ilIndex).Print "Station Agy Code"
                Case CREDITAPPROVALINDEX
                    pbcAgy(ilIndex).Print "Credit Approval"
                Case CREDITRESTRINDEX
                    pbcAgy(ilIndex).Print "Credit Restriction"
                Case PAYMRATINGINDEX
                    pbcAgy(ilIndex).Print "Payment Rating"
                Case CREDITRATINGINDEX
                    pbcAgy(ilIndex).Print "Credit Rating"
                Case ISCIINDEX
                    pbcAgy(ilIndex).Print "ISCI on Invoice"
                Case INVSORTINDEX
                    pbcAgy(ilIndex).Print "Invoice Sort"
                Case PACKAGEINDEX
                    pbcAgy(ilIndex).Print "Package Inv Show"
                Case PACKAGEINDEX
                    pbcAgy(ilIndex).Print "Package Inv Show"
            End Select
        End If
    Next ilLoop
    
    pbcAgy(ilIndex).Line (tmCtrls(CADDRINDEX).fBoxX - 15, tmCtrls(CADDRINDEX).fBoxY - 15)-(tmCtrls(CADDRINDEX).fBoxW + 30, tmCtrls(BUYERINDEX).fBoxY - 15), BLUE, B
    pbcAgy(ilIndex).CurrentX = tmCtrls(CADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcAgy(ilIndex).CurrentY = tmCtrls(CADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcAgy(ilIndex).Print "Contract Address"
    pbcAgy(ilIndex).Line (tmCtrls(BADDRINDEX).fBoxX - 15, tmCtrls(BADDRINDEX).fBoxY - 15)-(tmCtrls(BADDRINDEX).fBoxX + tmCtrls(BADDRINDEX).fBoxW, tmCtrls(BUYERINDEX).fBoxY - 15), BLUE, B
    pbcAgy(ilIndex).CurrentX = tmCtrls(BADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcAgy(ilIndex).CurrentY = tmCtrls(BADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcAgy(ilIndex).Print "Billing Address"

    For ilLoop = BUYERINDEX To XMLDATESINDEX Step 1
        pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case BUYERINDEX
                pbcAgy(ilIndex).Print "default Buyer/Phone # and Extension/Fax #"
                If ilIndex = 1 Then
                    pbcAgy(ilIndex).Line (tmCtrls(LKBOXINDEX).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(LKBOXINDEX).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B 'L.Bianchi 04/15/2021
                Else
                    pbcAgy(ilIndex).Line (tmCtrls(LKBOXINDEX).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(REFIDINDEX).fBoxW + tmCtrls(LKBOXINDEX).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B 'L.Bianchi 04/15/2021
                End If
                
            Case PAYABLEINDEX
                pbcAgy(ilIndex).Print "Payables Contact/Phone # and Extension/Fax #"
            Case CRMIDINDEX ' JD 09/16/2022
                pbcAgy(ilIndex).Print "CRM ID"
            Case LKBOXINDEX
                pbcAgy(ilIndex).Print "Lock Box Name"
               
            Case REFIDINDEX 'L.Bianchi 04/15/2021
                pbcAgy(ilIndex).Print "Ref Id"
            Case EDICINDEX
                pbcAgy(ilIndex).Print "EDI for Contracts"
            Case EDIIINDEX
                pbcAgy(ilIndex).Print "EDI for Invoices"
            Case TERMSINDEX
                pbcAgy(ilIndex).Print "Terms"
            Case TAXINDEX
                pbcAgy(ilIndex).Print "Commercial Tax"
        
            Case EXPORTFORMINDEX
                pbcAgy(ilIndex).Print "Form"
            Case XMLCALLINDEX
                pbcAgy(ilIndex).Print "Call-L"
            Case XMLBANDINDEX
                pbcAgy(ilIndex).Print "Band"
            Case XMLDATESINDEX
                pbcAgy(ilIndex).Print "Dates"
        End Select

        If ilLoop = TAXINDEX Then
            pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 30, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcAgy(ilIndex).Print "Proposal"
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY + tmCtrls(ilLoop).fBoxH / 2 - 30
            pbcAgy(ilIndex).Print "Export:"
        End If
       
    Next ilLoop
        
    For ilLoop = SUPPRESSNETINDEX To UNUSEDINDEX Step 1 ' TTP 10622 - 2023-03-08 JJB
        Select Case ilLoop
            Case SUPPRESSNETINDEX
                pbcAgy(ilIndex).Line (tmCtrls(SUPPRESSNETINDEX).fBoxX - 15, tmCtrls(SUPPRESSNETINDEX).fBoxY - 15)-Step(tmCtrls(SUPPRESSNETINDEX).fBoxW / 2, tmCtrls(SUPPRESSNETINDEX).fBoxH + 15), BLUE, B
                pbcAgy(ilIndex).CurrentX = tmCtrls(SUPPRESSNETINDEX).fBoxX + 15
                pbcAgy(ilIndex).CurrentY = tmCtrls(SUPPRESSNETINDEX).fBoxY - 15
                pbcAgy(ilIndex).Print "Suppress Net Amount for Trade Invoices"
            Case UNUSEDINDEX
                pbcAgy(ilIndex).Line (2730, 2765)-Step(9345, 345), BLUE, B
                pbcAgy(ilIndex).CurrentX = 3650 'tmCtrls(UNUSEDINDEX).fBoxX + 15
                pbcAgy(ilIndex).CurrentY = 2765 'tmCtrls(UNUSEDINDEX).fBoxY - 15
                pbcAgy(ilIndex).Print ""
        End Select
    Next ilLoop
   
    For ilLoop = PCT90INDEX To LSTTOPAYINDEX Step 1
        If ilLoop = TOTALGROSSINDEX Then
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX - 15, tmARCtrls(ilLoop).fBoxY - 15)-Step(tmARCtrls(ilLoop).fBoxW + tmARCtrls(ilLoop + 1).fBoxW + 45, tmARCtrls(ilLoop).fBoxH + 15), BLUE, B
        ElseIf ilLoop = DATEENTRDINDEX Then
            'Box part of TOTALGROSSINDEX
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 60), LIGHTERYELLOW, BF
        Else
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 30), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX - 15, tmARCtrls(ilLoop).fBoxY - 15)-Step(tmARCtrls(ilLoop).fBoxW + 15, tmARCtrls(ilLoop).fBoxH + 15), BLUE, B
        End If
        pbcAgy(ilIndex).CurrentX = tmARCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcAgy(ilIndex).CurrentY = tmARCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case PCT90INDEX
                pbcAgy(ilIndex).Print "% Over 90"
            Case CURRARINDEX
                pbcAgy(ilIndex).Print "Current A/R Total"
            Case UNBILLEDINDEX
                pbcAgy(ilIndex).Print "Unbilled + Projected"
            Case HICREDITINDEX
                pbcAgy(ilIndex).Print "Highest A/R Total"
            Case TOTALGROSSINDEX
                pbcAgy(ilIndex).Print "Total Gross"
            Case DATEENTRDINDEX
                pbcAgy(ilIndex).Print "since"
            Case NSFCHKSINDEX
                pbcAgy(ilIndex).Print "# NSF Checks"
            Case DATELSTINVINDEX
                pbcAgy(ilIndex).Print "Date Last Billed"
            Case DATELSTPAYMINDEX
                pbcAgy(ilIndex).Print "Date Last Payment"
            Case AVGTOPAYINDEX
                pbcAgy(ilIndex).Print "Avg # Days to Pay"
            Case LSTTOPAYINDEX
                pbcAgy(ilIndex).Print "# Days to Pay Last Payment"
        End Select
    Next ilLoop
    
    
    pbcAgy(ilIndex).FontSize = flFontSize
    pbcAgy(ilIndex).FontName = slFontName
    pbcAgy(ilIndex).FontSize = flFontSize
    pbcAgy(ilIndex).ForeColor = llColor
    pbcAgy(ilIndex).FontBold = True


End Sub

Private Sub mPaintAgyTitle2(ilIndex As Integer)
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcAgy(ilIndex).ForeColor
    slFontName = pbcAgy(ilIndex).FontName
    flFontSize = pbcAgy(ilIndex).FontSize
    pbcAgy(ilIndex).ForeColor = BLUE
    pbcAgy(ilIndex).FontBold = False
    pbcAgy(ilIndex).FontSize = 7
    pbcAgy(ilIndex).FontName = "Arial"
    pbcAgy(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    ''For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    For ilLoop = NAMEINDEX To PACKAGEINDEX Step 1
        If ilLoop <> CREDITRESTRINDEX + 1 Then
            If ilLoop = CREDITRESTRINDEX Then
                pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + tmCtrls(ilLoop + 1).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            Else
                pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            End If
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            Select Case ilLoop
                Case NAMEINDEX
                    pbcAgy(ilIndex).Print "Name"
                Case ABBRINDEX
                    pbcAgy(ilIndex).Print "Abbreviation"
                Case CITYINDEX
                    pbcAgy(ilIndex).Print "City ID"
                Case STATEINDEX
                    pbcAgy(ilIndex).Print "Active/Dormant"
                Case COMMINDEX
                    pbcAgy(ilIndex).Print "% Commission"
                    If ilIndex = 1 Then
                        pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - (pbcAgy(imPbcIndex).height + 100) - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                    End If
                Case SPERSONINDEX
                    pbcAgy(ilIndex).Print "default Salesperson"
                Case DIGITRATINGINDEX
                    'pbcAgy(ilIndex).Print "1 or 2 Place Rating"
                Case REPCODEINDEX
                    pbcAgy(ilIndex).Print "G/L Agency Code"
                Case STNCODEINDEX
                    pbcAgy(ilIndex).Print "Station Agy Code"
                Case CREDITAPPROVALINDEX
                    pbcAgy(ilIndex).Print "Credit Approval"
                Case CREDITRESTRINDEX
                    pbcAgy(ilIndex).Print "Credit Restriction"
                Case PAYMRATINGINDEX
                    pbcAgy(ilIndex).Print "Payment Rating"
                Case CREDITRATINGINDEX
                    pbcAgy(ilIndex).Print "Credit Rating"
                Case ISCIINDEX
                    pbcAgy(ilIndex).Print "ISCI on Invoice"
                Case INVSORTINDEX
                    pbcAgy(ilIndex).Print "Invoice Sort"
                Case PACKAGEINDEX
                    pbcAgy(ilIndex).Print "Package Inv Show"
                Case PACKAGEINDEX
                    pbcAgy(ilIndex).Print "Package Inv Show"
            End Select
        End If
    Next ilLoop
    pbcAgy(ilIndex).Line (tmCtrls(CADDRINDEX).fBoxX - 15, tmCtrls(CADDRINDEX).fBoxY - 15)-(tmCtrls(CADDRINDEX).fBoxW + 30, tmCtrls(BUYERINDEX).fBoxY - 15), BLUE, B
    pbcAgy(ilIndex).CurrentX = tmCtrls(CADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcAgy(ilIndex).CurrentY = tmCtrls(CADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcAgy(ilIndex).Print "Contract Address"
    pbcAgy(ilIndex).Line (tmCtrls(BADDRINDEX).fBoxX - 15, tmCtrls(BADDRINDEX).fBoxY - 15)-(tmCtrls(BADDRINDEX).fBoxX + tmCtrls(BADDRINDEX).fBoxW, tmCtrls(BUYERINDEX).fBoxY - 15), BLUE, B
    pbcAgy(ilIndex).CurrentX = tmCtrls(BADDRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcAgy(ilIndex).CurrentY = tmCtrls(BADDRINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcAgy(ilIndex).Print "Billing Address"
 
    For ilLoop = BUYERINDEX To SUPPRESSNETINDEX Step 1 ' TTP 10622 - 2023-03-08 JJB
        pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case BUYERINDEX
                pbcAgy(ilIndex).Print "default Buyer/Phone # and Extension/Fax #"
                If ilIndex = 1 Then
                    pbcAgy(ilIndex).Line (tmCtrls(LKBOXINDEX).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(LKBOXINDEX).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B 'L.Bianchi 04/15/2021
                Else
                    pbcAgy(ilIndex).Line (tmCtrls(LKBOXINDEX).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(REFIDINDEX).fBoxW + tmCtrls(LKBOXINDEX).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B 'L.Bianchi 04/15/2021
                End If
                
            Case PAYABLEINDEX
                pbcAgy(ilIndex).Print "Payables Contact/Phone # and Extension/Fax #"
            Case CRMIDINDEX ' JD 09/16/2022
                pbcAgy(ilIndex).Print "CRM ID"
            Case LKBOXINDEX
                pbcAgy(ilIndex).Print "Lock Box Name"
               
            Case REFIDINDEX 'L.Bianchi 04/15/2021
                pbcAgy(ilIndex).Print "Ref Id"
            Case EDICINDEX
                pbcAgy(ilIndex).Print "EDI for Contracts"
            Case EDIIINDEX
                pbcAgy(ilIndex).Print "EDI for Invoices"
            Case TERMSINDEX
                pbcAgy(ilIndex).Print "Terms"
            Case TAXINDEX
                pbcAgy(ilIndex).Print "Commercial Tax"
        
            Case EXPORTFORMINDEX
                pbcAgy(ilIndex).Print "Form"
            Case XMLCALLINDEX
                pbcAgy(ilIndex).Print "Call-L"
            Case XMLBANDINDEX
                pbcAgy(ilIndex).Print "Band"
            Case XMLDATESINDEX
                pbcAgy(ilIndex).Print "Dates"
            Case SUPPRESSNETINDEX ' TTP 10622 - 2023-03-08 JJB
                pbcAgy(ilIndex).Print "Suppress Net Trade on Inv"
        End Select

        If ilLoop = TAXINDEX Then
            pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 30, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop + 1).fBoxX - (tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW) - 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcAgy(ilIndex).Print "Proposal"
            pbcAgy(ilIndex).CurrentX = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 45
            pbcAgy(ilIndex).CurrentY = tmCtrls(ilLoop).fBoxY + tmCtrls(ilLoop).fBoxH / 2 - 30
            pbcAgy(ilIndex).Print "Export:"
        End If
       
    Next ilLoop
    
    For ilLoop = PCT90INDEX To LSTTOPAYINDEX Step 1
        If ilLoop = TOTALGROSSINDEX Then
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX - 15, tmARCtrls(ilLoop).fBoxY - 15)-Step(tmARCtrls(ilLoop).fBoxW + tmARCtrls(ilLoop + 1).fBoxW + 45, tmARCtrls(ilLoop).fBoxH + 15), BLUE, B
        ElseIf ilLoop = DATEENTRDINDEX Then
            'Box part of TOTALGROSSINDEX
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 60), LIGHTERYELLOW, BF
        Else
        pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX, tmARCtrls(ilLoop).fBoxY)-Step(tmARCtrls(ilLoop).fBoxW - 15, tmARCtrls(ilLoop).fBoxH - 30), LIGHTERYELLOW, BF
            pbcAgy(ilIndex).Line (tmARCtrls(ilLoop).fBoxX - 15, tmARCtrls(ilLoop).fBoxY - 15)-Step(tmARCtrls(ilLoop).fBoxW + 15, tmARCtrls(ilLoop).fBoxH + 15), BLUE, B
        End If
        pbcAgy(ilIndex).CurrentX = tmARCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcAgy(ilIndex).CurrentY = tmARCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case PCT90INDEX
                pbcAgy(ilIndex).Print "% Over 90"
            Case CURRARINDEX
                pbcAgy(ilIndex).Print "Current A/R Total"
            Case UNBILLEDINDEX
                pbcAgy(ilIndex).Print "Unbilled + Projected"
            Case HICREDITINDEX
                pbcAgy(ilIndex).Print "Highest A/R Total"
            Case TOTALGROSSINDEX
                pbcAgy(ilIndex).Print "Total Gross"
            Case DATEENTRDINDEX
                pbcAgy(ilIndex).Print "since"
            Case NSFCHKSINDEX
                pbcAgy(ilIndex).Print "# NSF Checks"
            Case DATELSTINVINDEX
                pbcAgy(ilIndex).Print "Date Last Billed"
            Case DATELSTPAYMINDEX
                pbcAgy(ilIndex).Print "Date Last Payment"
            Case AVGTOPAYINDEX
                pbcAgy(ilIndex).Print "Avg # Days to Pay"
            Case LSTTOPAYINDEX
                pbcAgy(ilIndex).Print "# Days to Pay Last Payment"
        End Select
    Next ilLoop
    
    
    pbcAgy(ilIndex).FontSize = flFontSize
    pbcAgy(ilIndex).FontName = slFontName
    pbcAgy(ilIndex).FontSize = flFontSize
    pbcAgy(ilIndex).ForeColor = llColor
    pbcAgy(ilIndex).FontBold = True


End Sub
Public Function mXmlChoicesOk() As Boolean
    Dim blRet As Boolean
    
    blRet = True
    If imExport = 2 Then
        If Len(edcXMLCall.Text) >= 1 And Len(edcXMLCall.Text) < 3 Then
            blRet = False
            MsgBox "Call letters must be 3 or 4 letters", vbOKOnly + vbExclamation, "Invalid"
        End If
        If Len(edcXMLBand.Text) > 0 And Not (InStr(1, edcXMLBand.Text, "AM", vbBinaryCompare) > 0 Or InStr(1, edcXMLBand.Text, "FM", vbBinaryCompare) > 0 Or InStr(1, edcXMLBand.Text, "DV", vbBinaryCompare) > 0 Or InStr(1, edcXMLBand.Text, "SM", vbBinaryCompare) > 0 Or InStr(1, edcXMLBand.Text, "N", vbBinaryCompare) > 0) Then
            blRet = False
            MsgBox "Band may only be 'AM','FM','DV','SM', or 'N'", vbOKOnly + vbExclamation, "Invalid"
        End If
        If Len(edcXMLBand.Text) = 2 And InStr(1, edcXMLBand.Text, "N", vbBinaryCompare) > 0 Then
            blRet = False
            MsgBox "Band may only be 'AM','FM','DV','SM', or 'N'", vbOKOnly + vbExclamation, "Invalid"
        End If
    End If
    mXmlChoicesOk = blRet
End Function
