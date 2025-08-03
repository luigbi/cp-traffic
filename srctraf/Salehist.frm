VERSION 5.00
Begin VB.Form SaleHist 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9435
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   9435
   Begin VB.Timer tmcCntrNo 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1110
      Top             =   5340
   End
   Begin VB.Frame frcSelection 
      Height          =   540
      Left            =   3075
      TabIndex        =   0
      Top             =   -30
      Width           =   6285
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
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   4050
      End
      Begin VB.TextBox edcCntrNo 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lacCntrNo 
         Caption         =   "or Contract #"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   195
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmcYear 
      Appearance      =   0  'Flat
      Caption         =   "&Year"
      Height          =   285
      Left            =   5985
      TabIndex        =   45
      Top             =   5250
      Width           =   1050
   End
   Begin VB.ListBox lbcTranType 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Salehist.frx":0000
      Left            =   1410
      List            =   "Salehist.frx":0002
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcKey 
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
      Height          =   1755
      Left            =   165
      Picture         =   "Salehist.frx":0004
      ScaleHeight     =   1725
      ScaleWidth      =   2595
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   285
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ListBox lbcTax 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Salehist.frx":E9DE
      Left            =   5190
      List            =   "Salehist.frx":E9E0
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3540
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   570
      Top             =   5325
   End
   Begin VB.PictureBox pbcTT 
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
      Height          =   210
      Left            =   4800
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   345
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox lbcSSPart 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Salehist.frx":E9E2
      Left            =   3195
      List            =   "Salehist.frx":E9E4
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4185
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcNTRType 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Salehist.frx":E9E6
      Left            =   1380
      List            =   "Salehist.frx":E9E8
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcBVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1995
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox plcType 
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
      Height          =   240
      Left            =   150
      ScaleHeight     =   240
      ScaleWidth      =   3120
      TabIndex        =   41
      Top             =   360
      Width           =   3120
      Begin VB.OptionButton rbcType 
         Caption         =   "Receivables"
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
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   43
         Top             =   0
         Width           =   1380
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Sales History"
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
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9300
      Top             =   4140
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
      Left            =   15
      Picture         =   "Salehist.frx":E9EA
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9315
      Top             =   3705
   End
   Begin VB.ListBox lbcAdvtAgyCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   6645
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1515
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
         TabIndex        =   26
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
         TabIndex        =   23
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
         Picture         =   "Salehist.frx":ECF4
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   24
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
            TabIndex        =   25
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
         TabIndex        =   22
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Salehist.frx":11B0E
      Left            =   2370
      List            =   "Salehist.frx":11B10
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcSPerson 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3090
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcProduct 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3915
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1710
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcAdvertiser 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   645
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2820
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcAgency 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   645
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox pbcCT 
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
      Height          =   210
      Left            =   4110
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   180
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
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1620
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
      Left            =   2385
      Picture         =   "Salehist.frx":11B12
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1635
      Visible         =   0   'False
      Width           =   195
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
      Left            =   9180
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3225
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
      Left            =   9210
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2865
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
      Left            =   9060
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   7680
      TabIndex        =   33
      Top             =   5580
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5355
      Width           =   75
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   4830
      TabIndex        =   32
      Top             =   5250
      Width           =   945
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
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   28
      Top             =   5070
      Width           =   60
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
      Height          =   120
      Left            =   30
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   435
      Width           =   105
   End
   Begin VB.VScrollBar vbcHist 
      Height          =   4260
      LargeChange     =   19
      Left            =   9060
      Min             =   1
      TabIndex        =   29
      Top             =   705
      Value           =   1
      Width           =   240
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3555
      TabIndex        =   31
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1260
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   1260
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2265
      TabIndex        =   30
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox pbcHist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   4275
      Left            =   150
      Picture         =   "Salehist.frx":11C0C
      ScaleHeight     =   4275
      ScaleWidth      =   8895
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   705
      Width           =   8895
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
         Height          =   240
         Left            =   0
         TabIndex        =   39
         Top             =   390
         Visible         =   0   'False
         Width           =   8895
      End
   End
   Begin VB.PictureBox plcHist 
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
      Height          =   4395
      Left            =   135
      ScaleHeight     =   4335
      ScaleWidth      =   9195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Width           =   9255
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8670
      Picture         =   "Salehist.frx":8E55E
      Top             =   5085
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1455
      Picture         =   "Salehist.frx":8E868
      Top             =   45
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   5145
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SaleHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Salehist.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SaleHist.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Commission input screen code
Option Explicit
Option Compare Text
Dim tmBVehicleCode() As SORTCODE
Dim smBVehicleCodeTag As String
Dim tmSSPart() As SSPART
'If tmCtrls is changed, then smShow must be changed
Dim tmCtrls(0 To 18)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim lmRowNo As Long
Dim smSave() As String * 60  '1=Agency Name; 2=Advertiser Name; 3=Product; 4=Salesperson;
                        '5=Invoice #; 6=Contract #; 7=Bill Vehicle; 8=Air Vehicle;
                        '9=Package Line #; 10=Transaction Date; 11=Gross Amount; 12=Net; \
                        '13=Transaction Type; 14=Acquisition Cost; 15=Total Tax Amount
Dim imSave() As Integer '1=Cash(0)/Trade(1); 2=NTR Index, 3=Participant; 4=NTR Tax Index
Dim smShow() As String * 50  '1=Cash/Trade; 2=Advertiser; 3=Product; 4=Agency Name; 5=Salesperson;
                        '6=Invoice #; 7=Contract #; 8=Bill Vehicle; 9=Air Vehicle;
                        '10=Package Line #; 11=Transaction Date; 12=NTR Type; 13=Transaction Type;
                        '14=Gross Amount; 15=Net; 16=SS/Participant
Dim lmSave() As Long    '1=SbfCode
Dim lmSalesHistStartDate As Long
Dim lmSalesHistEndDate As Long
Dim imWarningShown As Integer
Dim imPriceChgd As Integer
Dim tmAdvertiser() As SORTCODE
Dim smAdvertiserTag As String
Dim tmRvfPhf As RVF        'Phf record image
Dim tmRvfPhfSrchKey0 As RVFKEY0            'RVF record image (agency code)
Dim tmRvfPhfSrchKey1 As RVFKEY1            'RVF record image (Advertiser code)
Dim tmRvfPhfSrchKey4 As RVFKEY4            'RVF record image (Advertiser code)
Dim tmRvfSrchKey2 As LONGKEY0              'PRF record image
Dim hmPhf As Integer    'Payment History file handle
Dim hmRvf As Integer    'Receivable file handle
Dim imRvfPhfRecLen As Integer        'SLF record length
Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey0 As LONGKEY0            'PRF record image
Dim imPrfRecLen As Integer        'PRF record length
Dim tmPrf As PRF
Dim hmAdf As Integer            'Advertiser file handle
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAdf As ADF
Dim hmAgf As Integer            'Agency file handle
Dim tmAgfSrchKey As INTKEY0            'AGF record image
Dim imAgfRecLen As Integer        'AGF record length
Dim tmAgf As AGF
Dim hmVef As Integer            'Vehicle file handle
Dim imVefRecLen As Integer        'VEF record length
Dim tmVef As VEF
'Dim tmRec As LPOPREC
Dim hmSlf As Integer            'Salesperson file handle
Dim tmSlfSrchKey As INTKEY0            'SLF record image
Dim imSlfRecLen As Integer        'SLF record length
Dim tmSlf As SLF
Dim hmSof As Integer            'Sales Office file handle
Dim tmSofSrchKey As INTKEY0            'SOF record image
Dim imSofRecLen As Integer        'SOF record length
Dim tmSof As SOF
Dim imAdfCode As Integer
Dim imAgfCode As Integer
Dim imTaxDefined As Integer

Dim hmSbf As Integer            'Special Bill file handle
Dim tmSbfSrchKey1 As LONGKEY0            'SBF record image
Dim imSbfRecLen As Integer        'SBF record length
Dim tmSbf As SBF

Dim hmPif As Integer
Dim tmPif() As PIF

Dim tmTranTypeCode() As SORTCODE
Dim smTranTypeCodeTag As String

Dim tmNTRTypeCode() As SORTCODE
Dim smNTRTypeCodeTag As String

Dim tmTaxSortCode() As SORTCODE
Dim smTaxSortCodeTag As String

Dim tmPhfSplit() As PHFREC

Dim smSMnfStamp As String
Dim smHMnfStamp As String
Dim tmSMnf() As MNF
Dim tmHMnf() As MNF
Dim imSSPartSplit As Integer
Dim imPhfChg As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imSettingValue As Integer
Dim smNowDate As String
Dim inGenMsgRet As Integer

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1

Const CTINDEX = 1
Const ADVTINDEX = 2
Const PRODINDEX = 3
Const AGYINDEX = 4
Const SPERSONINDEX = 5
Const INVNOINDEX = 6
Const CNTRNOINDEX = 7
Const BVEHINDEX = 8
Const AVEHINDEX = 9
Const PKLNINDEX = 10
Const TRANDATEINDEX = 11
Const NTRTYPEINDEX = 12
Const NTRTAXINDEX = 13
Const TRANTYPEINDEX = 14
Const GROSSINDEX = 15
Const NETINDEX = 16
Const ACQCOSTINDEX = 17
Const SSPARTINDEX = 18

'*******************************************************
'*                                                     *
'*      Procedure Name:mNTRTypePop                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate NTR Types             *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mNTRTypePop()
'
'   mAgyDPPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcNTRType.ListIndex
    If ilIndex > 0 Then
        slName = lbcNTRType.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(Collect, lbcNTRType, lbcNTRTypeCode, "YW")
    ilRet = gPopMnfPlusFieldsBox(SaleHist, lbcNTRType, tmNTRTypeCode(), smNTRTypeCodeTag, "I")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mNTRTypePopErr
        gCPErrorMsg ilRet, "mNTRTypePop (gPopMnfPlusFieldsBox)", SaleHist
        On Error GoTo 0
        lbcNTRType.AddItem "[None]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcNTRType
            If gLastFound(lbcNTRType) > 0 Then
                lbcNTRType.ListIndex = gLastFound(lbcNTRType)
            Else
                lbcNTRType.ListIndex = -1
            End If
        Else
            lbcNTRType.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mNTRTypePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub




Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        Screen.MousePointer = vbHourglass  'Wait
        ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
        If ilRet = 0 Then
            ilIndex = cbcSelect.ListIndex   'cbcSelect.ListCount - cbcSelect.ListIndex
            If ilIndex > 1 Then
                slNameCode = lbcAdvtAgyCode.List(ilIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If (InStr(slNameCode, "/Direct") > 0) Or (InStr(slNameCode, "/Non-Payee") > 0) Then
                    If Not mReadRvfPhfRec(Val(slCode), 0, "") Then
                        GoTo cbcSelectErr
                    End If
                Else
                    If Not mReadRvfPhfRec(0, Val(slCode), "") Then
                        GoTo cbcSelectErr
                    End If
                End If
            ElseIf ilIndex = 1 Then
                If Not mReadRvfPhfRec(0, 0, "") Then
                    GoTo cbcSelectErr
                End If
            Else
                ilRet = 1
            End If
        Else
            If ilRet = 1 Then
                cbcSelect.ListIndex = 0
            End If
            ilRet = 1   'Clear fields as no match name found
        End If
        pbcHist.Cls
        edcCntrNo.Text = ""
        If ilRet = 0 Then
            imSelectedIndex = cbcSelect.ListIndex
            mMoveRecToCtrl
            mInitShow
        Else
            imSelectedIndex = 0
            mClearCtrlFields
        End If
        'pbcHist_Paint
        mSetMinMax
        mSetCommands
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault    'Default
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
Private Sub cbcSelect_GotFocus()

    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
    If imFirstFocus Then
        imFirstFocus = False
        'cbcSelect.ListIndex = 0
    End If
    imComboBoxIndex = imSelectedIndex
    gCtrlGotFocus cbcSelect
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
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDone_Click()
    If (imUpdateAllowed) And (igPasswordOk) Then
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
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case CTINDEX
        Case ADVTINDEX
            lbcAdvertiser.Visible = Not lbcAdvertiser.Visible
        Case PRODINDEX
            lbcProduct.Visible = Not lbcProduct.Visible
        Case AGYINDEX
            lbcAgency.Visible = Not lbcAgency.Visible
        Case SPERSONINDEX
            lbcSPerson.Visible = Not lbcSPerson.Visible
        Case INVNOINDEX
        Case CNTRNOINDEX
        Case BVEHINDEX
            lbcBVehicle.Visible = Not lbcBVehicle.Visible
        Case AVEHINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case PKLNINDEX
        Case TRANDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case NTRTYPEINDEX  'NTR Type
            lbcNTRType.Visible = Not lbcNTRType.Visible
        Case NTRTAXINDEX  'NTR Tax
            lbcTax.Visible = Not lbcTax.Visible
        Case TRANTYPEINDEX  'NTR Type
            lbcTranType.Visible = Not lbcTranType.Visible
        Case GROSSINDEX
        Case NETINDEX
        Case ACQCOSTINDEX
        Case SSPARTINDEX  'Vehicle
            lbcSSPart.Visible = Not lbcSSPart.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcReport_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub
Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSave_Click()
    Dim ilLoop As Integer

    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    ReDim tgPhfDel(0 To 0) As PHFREC
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imPhfChg = False
    If imSSPartSplit Then
        'Merge tmPhfSplit into tgPhfRec
        For ilLoop = LBound(tmPhfSplit) To UBound(tmPhfSplit) - 1 Step 1
            tgPhfRec(UBound(tgPhfRec)).tPhf = tmPhfSplit(ilLoop).tPhf
            tgPhfRec(UBound(tgPhfRec)).iStatus = tmPhfSplit(ilLoop).iStatus
            tgPhfRec(UBound(tgPhfRec)).lRecPos = tmPhfSplit(ilLoop).lRecPos
            ReDim Preserve tgPhfRec(0 To UBound(tgPhfRec) + 1) As PHFREC
        Next ilLoop
    End If
    pbcHist.Cls
    mBuildKey
    mMoveRecToCtrl
    mInitShow
    mSetMinMax
    mSetCommands
    '1/3/18: Client was failing with this call sometimes.
    'pbcSTab.SetFocus
End Sub
Private Sub cmcSave_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub
Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcYear_Click()
    Dim slYear As String
    Dim slDate As String

    If (rbcType(0).Value) And (imSelectedIndex = 1) And (cmcSave.Enabled = False) Then
        sgGenMsg = "Sales History Year to View"
        sgCMCTitle(0) = "Done"
        sgCMCTitle(1) = "Cancel"
        sgCMCTitle(2) = ""
        sgCMCTitle(3) = ""
        igDefCMC = 0
        igEditBox = 1
        sgEditValue = Year(gAdjYear(Format$(lmSalesHistEndDate, "m/d/yy")))
        slYear = sgEditValue
        GenMsg.Show vbModal
        If igAnsCMC = 0 Then
            If Val(sgEditValue) <= Val(slYear) Then
                Screen.MousePointer = vbHourglass
                slDate = "6/1/" & sgEditValue
                lmSalesHistStartDate = gDateValue(gObtainYearStartDate(0, slDate))
                lmSalesHistEndDate = gDateValue(gObtainYearEndDate(0, slDate))
                If Not mReadRvfPhfRec(0, 0, "") Then
                End If
                pbcHist.Cls
                mMoveRecToCtrl
                mInitShow
                mSetMinMax
                mSetCommands
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
End Sub

Private Sub edcCntrNo_Change()
    If imChgMode = True Then
        Exit Sub
    End If
    tmcCntrNo.Enabled = False
    If imSelectedIndex <> -1 Then
        imChgMode = True
        cbcSelect.ListIndex = -1
        imSelectedIndex = -1
    End If
    mClearCtrlFields
    pbcHist.Cls
    pbcHist_Paint
    imChgMode = False
    tmcCntrNo.Enabled = True
End Sub

Private Sub edcCntrNo_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
    If imSelectedIndex <> -1 Then
        tmcCntrNo.Enabled = True
    End If
End Sub

Private Sub edcCntrNo_LostFocus()
    tmcCntrNo.Enabled = False
    tmcCntrNo_Timer
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case CTINDEX
        Case ADVTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAdvertiser, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAdvertiser.ListIndex = 0
            End If
        Case PRODINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProduct, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcProduct.ListIndex = 0
            End If
        Case AGYINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAgency, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAgency.ListIndex = 0
            End If
        Case SPERSONINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSPerson, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSPerson.ListIndex = 0
            End If
        Case INVNOINDEX
        Case CNTRNOINDEX
        Case BVEHINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcBVehicle, imBSMode, slStr)
            If ilRet = 1 Then
                'lbcBVehicle.ListIndex = 0
            End If
        Case AVEHINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case PKLNINDEX
        Case TRANDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case NTRTYPEINDEX  'NTR Type
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcNTRType, imBSMode, slStr)
            If ilRet = 1 Then
                lbcNTRType.ListIndex = 0
            End If
        Case NTRTAXINDEX  'NTR Tax
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTax, imBSMode, imComboBoxIndex
        Case TRANTYPEINDEX  'NTR Tax
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTranType, imBSMode, imComboBoxIndex
        Case GROSSINDEX
        Case NETINDEX
        Case ACQCOSTINDEX
        Case SSPARTINDEX  'Vehicle
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSSPart, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSSPart.ListIndex = 0
            End If
    End Select
    'edcDropDown.SelStart = 0
    'edcDropDown.SelLength = Len(edcDropDown.Text)
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case CTINDEX
        Case ADVTINDEX
        Case PRODINDEX
        Case AGYINDEX
        Case SPERSONINDEX
        Case INVNOINDEX
        Case CNTRNOINDEX
        Case BVEHINDEX
        Case AVEHINDEX
            imComboBoxIndex = lbcVehicle.ListIndex
        Case PKLNINDEX
        Case TRANDATEINDEX
        Case GROSSINDEX
        Case NETINDEX
        Case ACQCOSTINDEX
    End Select
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
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    Select Case imBoxNo
        Case CTINDEX
        Case ADVTINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case PRODINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case AGYINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case SPERSONINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case INVNOINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case CNTRNOINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case BVEHINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case AVEHINDEX
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case PKLNINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case TRANDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case NTRTYPEINDEX  'NTR Type
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
        Case NTRTAXINDEX  'NTR Tax
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
         Case TRANTYPEINDEX  'NTR Tax
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
       Case GROSSINDEX
            If Not mDropDownKeyPress(ilKey, True) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case NETINDEX
            If Not mDropDownKeyPress(ilKey, True) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case ACQCOSTINDEX
            If Not mDropDownKeyPress(ilKey, True) Then
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case CTINDEX
            Case ADVTINDEX
                gProcessArrowKey Shift, KeyCode, lbcAdvertiser, imLbcArrowSetting
            Case PRODINDEX
                gProcessArrowKey Shift, KeyCode, lbcProduct, imLbcArrowSetting
            Case AGYINDEX
                gProcessArrowKey Shift, KeyCode, lbcAgency, imLbcArrowSetting
            Case SPERSONINDEX
                gProcessArrowKey Shift, KeyCode, lbcSPerson, imLbcArrowSetting
            Case INVNOINDEX
            Case CNTRNOINDEX
            Case BVEHINDEX
                gProcessArrowKey Shift, KeyCode, lbcBVehicle, imLbcArrowSetting
            Case AVEHINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case PKLNINDEX
            Case TRANDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case NTRTYPEINDEX  'NTR Type
                gProcessArrowKey Shift, KeyCode, lbcNTRType, imLbcArrowSetting
            Case NTRTAXINDEX  'NTR Tax
                gProcessArrowKey Shift, KeyCode, lbcTax, imLbcArrowSetting
            Case TRANTYPEINDEX  'NTR Tax
                gProcessArrowKey Shift, KeyCode, lbcTranType, imLbcArrowSetting
            Case GROSSINDEX
            Case NETINDEX
            Case ACQCOSTINDEX
            Case SSPARTINDEX 'Vehicle
                gProcessArrowKey Shift, KeyCode, lbcSSPart, imLbcArrowSetting
        End Select
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case CTINDEX
            Case ADVTINDEX
            Case PRODINDEX
            Case AGYINDEX
            Case SPERSONINDEX
            Case INVNOINDEX
            Case CNTRNOINDEX
            Case BVEHINDEX
            Case AVEHINDEX
            Case PKLNINDEX
            Case TRANDATEINDEX
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
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case NTRTYPEINDEX  'NTR Type
            Case NTRTAXINDEX  'NTR Tax
            Case TRANTYPEINDEX  'NTR Tax
            Case GROSSINDEX
            Case NETINDEX
            Case ACQCOSTINDEX
            Case SSPARTINDEX
        End Select
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case ADVTINDEX, PRODINDEX, AGYINDEX, SPERSONINDEX, NTRTYPEINDEX
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
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If ((igWinStatus(INVOICESJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName)) Or (igPasswordOk = False) Then
        pbcHist.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcHist.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    'gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.Refresh
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
    Dim llDate As Long
    Dim ilRet As Integer

    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If (igWinStatus(INVOICESJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        igPasswordOk = False
        imWarningShown = True
    'ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
    ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Then
        igPasswordOk = True
    Else
        gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llDate
        If llDate > 0 Then
            CSPWord.Show vbModal
        Else
            igPasswordOk = True
            imWarningShown = True
        End If
    End If
    If Not igPasswordOk Then
        pbcHist.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        imWarningShown = True
    End If
    If imTerminate Then
        tmcTerminate.Enabled = True
        'cmcCancel_Click
    Else
        If imWarningShown = False Then
            ilRet = MsgBox("Warning:  Changing Amounts will Affect Accounts Receivables!, Continue?", vbYesNo + vbQuestion, "Backlog")
            If ilRet = vbNo Then
                tmcTerminate.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmBVehicleCode
    Erase tgPhfRec
    Erase tgPhfDel
    Erase tmPhfSplit
    Erase smSave
    Erase imSave
    Erase smShow
    Erase tmHMnf
    Erase tmSMnf
    Erase tmSSPart
    Erase tmPif
    Erase tmTranTypeCode
    Erase tmNTRTypeCode
    Erase tmTaxSortCode
    btrExtClear hmRvf   'Clear any previous extend operation
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    btrExtClear hmPhf   'Clear any previous extend operation
    ilRet = btrClose(hmPhf)
    btrDestroy hmPhf
    btrExtClear hmPrf   'Clear any previous extend operation
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmSlf   'Clear any previous extend operation
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    btrExtClear hmSof   'Clear any previous extend operation
    ilRet = btrClose(hmSof)
    btrDestroy hmSof
    btrExtClear hmSbf   'Clear any previous extend operation
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    btrExtClear hmPif   'Clear any previous extend operation
    ilRet = btrClose(hmPif)
    btrDestroy hmPif
    
    Set SaleHist = Nothing
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub imcTrash_Click()
    Dim llLoop As Long
    Dim ilIndex As Integer
    Dim llUpperBound As Long
    Dim llRowNo As Long
    Dim llPhf As Long
    If (lmRowNo < vbcHist.Value) Or (lmRowNo > vbcHist.Value + vbcHist.LargeChange) Then
        Exit Sub
    End If
    llRowNo = lmRowNo
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    llUpperBound = UBound(smSave, 2)
    llPhf = llRowNo
    If llPhf = llUpperBound Then
        mInitNew llPhf
    Else
        If llPhf > 0 Then
            If tgPhfRec(llPhf).iStatus = 1 Then
                tgPhfDel(UBound(tgPhfDel)).tPhf = tgPhfRec(llPhf).tPhf
                tgPhfDel(UBound(tgPhfDel)).iStatus = tgPhfRec(llPhf).iStatus
                tgPhfDel(UBound(tgPhfDel)).lRecPos = tgPhfRec(llPhf).lRecPos
                ReDim Preserve tgPhfDel(0 To UBound(tgPhfDel) + 1) As PHFREC
            End If
            llPhf = llRowNo
            'Remove record from tgRjf1Rec- Leave tgPjf2Rec
            For llLoop = llRowNo To llUpperBound - 1 Step 1
                tgPhfRec(llLoop) = tgPhfRec(llLoop + 1)
            Next llLoop
            ReDim Preserve tgPhfRec(0 To UBound(tgPhfRec) - 1) As PHFREC
        End If
        For llLoop = llRowNo To llUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, llLoop) = smSave(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imSave, 1) Step 1
                imSave(ilIndex, llLoop) = imSave(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, llLoop) = smShow(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(lmSave, 1) Step 1
                lmSave(ilIndex, llLoop) = lmSave(ilIndex, llLoop + 1)
            Next ilIndex
        Next llLoop
        llUpperBound = UBound(smSave, 2)
        ReDim Preserve smShow(0 To 18, 0 To llUpperBound - 1) As String * 50 'Values shown in program area
        ReDim Preserve smSave(0 To 15, 0 To llUpperBound - 1) As String * 60    'Values saved (program name) in program area
        ReDim Preserve imSave(0 To 4, 0 To llUpperBound - 1) As Integer 'Values saved (program name) in program area
        ReDim Preserve lmSave(0 To 1, 0 To llUpperBound - 1) As Long
        imPhfChg = True
    End If
    mSetCommands
    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSettingValue = True
    vbcHist.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcHist.LargeChange + 1 Then ' + 1 Then
        vbcHist.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcHist.Max = UBound(smShow, 2) - vbcHist.LargeChange
    End If
    imSettingValue = True
    vbcHist.Value = vbcHist.Min
    imSettingValue = True
    pbcHist.Cls
    pbcHist_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcAdvertiser_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcAdvertiser, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcAdvertiser_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcAdvertiser_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcAdvertiser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcAdvertiser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAdvertiser, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcAgency_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
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
Private Sub lbcBVehicle_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcBVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub
Private Sub lbcBVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcBVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcBVehicle, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcNTRType_Click()
    imLbcArrowSetting = False
    gProcessLbcClick lbcNTRType, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcNTRType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcProduct_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcProduct, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcProduct_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcProduct_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcProduct_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcProduct_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcProduct, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcSPerson_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
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

Private Sub lbcSSPart_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        'pbcEatTab(1).Enabled = True
        'pbcEatTab(0).Enabled = True
        'pbcEatTab(0).SetFocus
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcSSPart, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcSSPart_DblClick()
    tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcSSPart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcTax_Click()
    gProcessLbcClick lbcTax, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcTax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcTranType_Click()
    gProcessLbcClick lbcTranType, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcTranType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehicle_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddProd                        *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Product if required        *
'*                                                     *
'*******************************************************
Private Function mAddProd(slProduct As String, ilAdfCode As Integer) As Long
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    If (Trim$(slProduct) <> "") And (Trim$(slProduct) <> "[None]") Then
        mProdPop ilAdfCode
        gFindMatch slProduct, 2, lbcProduct
        If gLastFound(lbcProduct) < 0 Then
            gGetSyncDateTime slSyncDate, slSyncTime
            tmPrf.lCode = 0
            tmPrf.iAdfCode = ilAdfCode
            tmPrf.sName = slProduct
            tmPrf.iMnfComp(0) = 0
            tmPrf.iMnfComp(1) = 0
            tmPrf.iMnfExcl(0) = 0
            tmPrf.iMnfExcl(1) = 0
            tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
            tmPrf.lAutoCode = tmPrf.lCode
            ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                mAddProd = 0
                Exit Function
            End If
            mAddProd = tmPrf.lCode
            Do
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                gPackTime slSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        Else
            If gLastFound(lbcProduct) > 1 Then
                slNameCode = tgProdCode(gLastFound(lbcProduct) - 2).sKey 'lbcProdCode.List(gLastFound(lbcProduct) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    slCode = Trim$(slCode)
                    mAddProd = Val(slCode)
                Else
                    mAddProd = 0
                End If
            Else
                mAddProd = 0
            End If
        End If
    Else
        mAddProd = 0
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdjSSPart                      *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update Participant records      *
'*                                                     *
'*******************************************************
Private Function mAdjSSPart(tlPhf As RVF) As Integer
    Dim slTranType As String
    Dim ilLoop As Integer
    Dim slPct As String
    Dim slDollar As String
    Dim slGross As String
    Dim slNet As String
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim hlFile As Integer
    Dim slTranDate As String

    If rbcType(0).Value Then
        hlFile = hmPhf
    Else
        hlFile = hmRvf
    End If
    ilRet = 0
    gUnpackDate tlPhf.iTranDate(0), tlPhf.iTranDate(1), slTranDate
    mSSPartPop -1, tlPhf.iAirVefCode, tlPhf.iSlfCode, tlPhf.iAdfCode, 0, slTranDate
    'If (UBound(tmSSPart) > LBound(tmSSPart)) And ((tlPhf.sTranType = "HI") Or (tlPhf.sTranType = "IN")) Then
    If (UBound(tmSSPart) > LBound(tmSSPart)) And (tlPhf.iMnfGroup = -1) Then
        slTranType = tlPhf.sTranType
        gPDNToStr tlPhf.sGross, 2, slGross
        gPDNToStr tlPhf.sNet, 2, slNet
        'slRunGross = ".00"
        'slRunNet = ".00"
        For ilLoop = LBound(tmSSPart) To UBound(tmSSPart) - 1 Step 1
            'If ((tmSSPart(ilLoop).sUpdateRvf = "N") And (slTranType = "HI")) Or ((tmSSPart(ilLoop).sUpdateRvf <> "N") And (slTranType = "IN")) Then
                If rbcType(0).Value Then
                    hlFile = hmPhf
                Else
                    If (tmSSPart(ilLoop).sUpdateRVF = "N") Or (tmSSPart(ilLoop).sUpdateRVF = "E") Then
                        hlFile = hmPhf
                        If tlPhf.sTranType = "IN" Then
                            tlPhf.sTranType = "HI"
                        End If
                    Else
                        hlFile = hmRvf
                    End If
                End If
                imSSPartSplit = True
                tlPhf.sInvoiceUndone = "N"
                tlPhf.lCode = 0
                tlPhf.iMnfGroup = tmSSPart(ilLoop).iMnfGroup
                slPct = gIntToStrDec(tmSSPart(ilLoop).iProdPct, 2)
                slDollar = gDivStr(gMulStr(slGross, slPct), "100.00")
                'slRunGross = gAddStr(slRunGross, slDollar)
                gStrToPDN slDollar, 2, 6, tlPhf.sGross
                slDollar = gDivStr(gMulStr(slNet, slPct), "100.00")
                'slRunNet = gAddStr(slRunNet, slDollar)
                gStrToPDN slDollar, 2, 6, tlPhf.sNet
                ilRet = btrInsert(hlFile, tlPhf, imRvfPhfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit For
                End If
                If ilLoop <> UBound(tmSSPart) - 1 Then
                    If (rbcType(0).Value) Or ((rbcType(1).Value) And (hlFile = hmRvf)) Then
                        ilUpper = UBound(tmPhfSplit)
                        tmPhfSplit(ilUpper).tPhf = tlPhf
                        ilRet = btrGetPosition(hlFile, tmPhfSplit(ilUpper).lRecPos)
                        tmPhfSplit(ilUpper).iStatus = 1
                        ReDim Preserve tmPhfSplit(0 To UBound(tmPhfSplit) + 1) As PHFREC
                    End If
                Else
                End If
            'End If
        Next ilLoop
    Else
        'tlPhf.iMnfGroup = 0
        tlPhf.sInvoiceUndone = "N"
        ilRet = btrInsert(hlFile, tlPhf, imRvfPhfRecLen, INDEXKEY0)
    End If
    mAdjSSPart = ilRet
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      advertiser and process         *
'*                      communication back from        *
'*                      advertiser                     *
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
Private Function mAdvtBranch() As Integer
'
'   ilRet = mAdvtBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcAdvertiser, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) And (edcDropDown.Text <> "[New]") Then
        mAdvtBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(ADVERTISERSLIST)) Then
    '    imDoubleClickName = False
    '    mAdvtBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igAdvtCallSource = CALLSOURCEPOSTITEM
    If edcDropDown.Text = "[New]" Then
        sgAdvtName = ""
    Else
        sgAdvtName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & sgAdvtName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & sgAdvtName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Advt.Exe " & slStr, 1)
    'SaleHist.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Advt.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtName)
    igAdvtCallSource = Val(sgAdvtName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtName)
    'SaleHist.Enabled = True
    'Invoice!edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAdvtBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'gShowBranner imUpdateAllowed
    If igAdvtCallSource = CALLDONE Then  'Done
        igAdvtCallSource = CALLNONE
        lbcAdvertiser.Clear
        smAdvertiserTag = ""
'        sgCommAdfStamp = ""
        mAdvtPop
        If imTerminate Then
            mAdvtBranch = False
            Exit Function
        End If
'        End If
        gFindMatch sgAdvtName, 1, lbcAdvertiser
        sgAdvtName = ""
        If gLastFound(lbcAdvertiser) > 0 Then
            imChgMode = True
            lbcAdvertiser.ListIndex = gLastFound(lbcAdvertiser)
            edcDropDown.Text = lbcAdvertiser.List(lbcAdvertiser.ListIndex)
            imChgMode = False
            mAdvtBranch = False
        Else
            imChgMode = True
            lbcAdvertiser.ListIndex = 0
            edcDropDown.Text = lbcAdvertiser.List(lbcAdvertiser.ListIndex)
            imChgMode = False
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igAdvtCallSource = CALLCANCELLED Then  'Cancelled
        igAdvtCallSource = CALLNONE
        sgAdvtName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtCallSource = CALLTERMINATED Then
        igAdvtCallSource = CALLNONE
        sgAdvtName = ""
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
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilPass As Integer
    Dim hlFile As Integer

    ilIndex = lbcAdvertiser.ListIndex
    If ilIndex >= 1 Then
        slName = lbcAdvertiser.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(SaleHist, lbcAdvertiser, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(SaleHist, lbcAdvertiser, tmAdvertiser(), smAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", SaleHist
        On Error GoTo 0
        'Check for Advertiser that were one time Direct but are not now
        For ilPass = 0 To 1 Step 1
            If ilPass = 0 Then
                hlFile = hmRvf
            Else
                hlFile = hmPhf
            End If
            tmRvfPhfSrchKey0.iAgfCode = 0
            ilRet = btrGetGreaterOrEqual(hlFile, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet <> BTRV_ERR_END_OF_FILE) And (tmRvfPhf.iAgfCode = 0)
                ilIndex = gBinarySearchAdf(tmRvfPhf.iAdfCode)
                If ilIndex <> -1 Then
                    If tgCommAdf(ilIndex).sBillAgyDir <> "D" Then
                        For ilLoop = LBound(tmAdvertiser) To UBound(tmAdvertiser) - 1 Step 1
                            slNameCode = tmAdvertiser(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tgCommAdf(ilIndex).iCode Then
                                If InStr(slNameCode, "/Non-Payee") <= 0 Then
                                    If Trim$(tgCommAdf(ilIndex).sAddrID) <> "" Then
                                        slName = Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Non-Payee"
                                        slName = slName & "\" & Trim$(str$(tgCommAdf(ilIndex).iCode))
                                        tmAdvertiser(ilLoop).sKey = slName
                                        lbcAdvertiser.List(ilLoop) = Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Non-Payee"
                                    Else
                                        slName = Trim$(tgCommAdf(ilIndex).sName) & "/Non-Payee"
                                        slName = slName & "\" & Trim$(str$(tgCommAdf(ilIndex).iCode))
                                        tmAdvertiser(ilLoop).sKey = slName
                                        lbcAdvertiser.List(ilLoop) = Trim$(tgCommAdf(ilIndex).sName) & "/Non-Payee"
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
                ilRet = btrGetNext(hlFile, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass
        lbcAdvertiser.AddItem "[None]", 0  'Force as first item on list
        lbcAdvertiser.AddItem "[New]", 0  'Force as first item on list
        If ilIndex >= 1 Then
            gFindMatch slName, 1, lbcAdvertiser
            If gLastFound(lbcAdvertiser) >= 1 Then
                lbcAdvertiser.ListIndex = gLastFound(lbcAdvertiser)
            Else
                lbcAdvertiser.ListIndex = -1
            End If
        Else
            lbcAdvertiser.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    If ((ilRet = 0) And (Not imDoubleClickName) And (edcDropDown.Text <> "[New]")) Or (edcDropDown.Text = "[Direct]") Then
        mAgencyBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(AGENCIESLIST)) Then
    '    imDoubleClickName = False
    '    mAgencyBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igAgyCallSource = CALLSOURCEPOSTITEM
    If lbcAgency.Text = "[New]" Then
        sgAgyName = ""
    Else
        sgAgyName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & Trim$(str$(igAgyCallSource)) & "\" & sgAgyName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & Trim$(str$(igAgyCallSource)) & "\" & sgAgyName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAgyCallSource)) & "\" & sgAgyName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAgyCallSource)) & "\" & sgAgyName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Agency.Exe " & slStr, 1)
    'SaleHist.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Agency.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAgyName)
    igAgyCallSource = Val(sgAgyName)
    ilParse = gParseItem(slStr, 2, "\", sgAgyName)
    'SaleHist.Enabled = True
    'Invoice!edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAgencyBranch = True
    imUpdateAllowed = ilUpdateAllowed
    If igAgyCallSource = CALLDONE Then  'Done
        igAgyCallSource = CALLNONE
        lbcAgency.Clear
        sgAgencyTag = ""
'        sgCommAgfStamp = ""
        mAgencyPop
        If imTerminate Then
            mAgencyBranch = False
            Exit Function
        End If
        gFindMatch sgAgyName, 2, lbcAgency
        sgAgyName = ""
        If gLastFound(lbcAgency) > 1 Then
            imChgMode = True
            lbcAgency.ListIndex = gLastFound(lbcAgency)
            edcDropDown.Text = lbcAgency.List(lbcAgency.ListIndex)
            imChgMode = False
            mAgencyBranch = False
        Else
            imChgMode = True
            lbcAgency.ListIndex = 0
            imChgMode = False
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igAgyCallSource = CALLCANCELLED Then  'Cancelled
        igAgyCallSource = CALLNONE
        sgAgyName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAgyCallSource = CALLTERMINATED Then
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
    If ilIndex >= 0 Then
        slName = lbcAgency.List(ilIndex)
    End If
    'Repopulate if required- if agency changed by another user while in this screen
    'ilRet = gPopAgyBox(SaleHist, lbcAgency, Traffic!lbcAgency)
    ilRet = gPopAgyBox(SaleHist, lbcAgency, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", SaleHist
        On Error GoTo 0
        lbcAgency.AddItem "[Direct]", 0  'Force as first item on list
        lbcAgency.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcAgency
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
    slStr = edcDropDown.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
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
'*      Procedure Name:mBVehPop                        *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mBVehPop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcBVehicle.ListIndex
    If ilIndex >= 0 Then
        slName = lbcBVehicle.List(ilIndex)
    End If
    'ilRet = gPopUserVehicleBox(SaleHist, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH + DORMANTVEH, lbcBVehicle, lbcBVehicleCode)
    ilRet = gPopUserVehicleBox(SaleHist, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + VEHPACKAGE + ACTIVEVEH + DORMANTVEH, lbcBVehicle, tmBVehicleCode(), smBVehicleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBVehPopErr
        gCPErrorMsg ilRet, "mBVehPop (gPopUserVehicleBox: Vehicle)", SaleHist
        On Error GoTo 0
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcBVehicle
            If gLastFound(lbcBVehicle) >= 0 Then
                lbcBVehicle.ListIndex = gLastFound(lbcBVehicle)
            Else
                lbcBVehicle.ListIndex = -1
            End If
        Else
            lbcBVehicle.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mBVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilLoop As Integer

    imPhfChg = False
    lbcAdvertiser.ListIndex = -1
    lbcProduct.ListIndex = -1
    lbcAgency.ListIndex = -1
    lbcSPerson.ListIndex = -1
    lbcBVehicle.ListIndex = -1
    lbcVehicle.ListIndex = -1
    ReDim tgPhfRec(0 To 1) As PHFREC
    tgPhfRec(1).iStatus = -1
    tgPhfRec(1).lRecPos = 0
    ReDim tgPhfDel(0 To 0) As PHFREC
    tgPhfDel(0).iStatus = -1
    tgPhfDel(0).lRecPos = 0
    ReDim smShow(0 To 18, 0 To 1) As String * 50 'Values shown in program area
    ReDim smSave(0 To 15, 0 To 1) As String * 60 'Values saved (program name) in program area
    ReDim imSave(0 To 4, 0 To 1) As Integer 'Values saved (program name) in program area
    ReDim lmSave(0 To 1, 0 To 1) As Long
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
        imSave(ilLoop, 1) = -1
    Next ilLoop
    For ilLoop = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
        lmSave(ilLoop, 1) = 0
    Next ilLoop
    vbcHist.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    vbcHist.Max = LBONE 'LBound(smShow, 2)
    imSettingValue = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateVef                      *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create vehicle if required     *
'*                                                     *
'*******************************************************
Private Sub mCreateVef(slName As String)
    Dim ilRet As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    gFindMatch slName, 0, lbcBVehicle
    If gLastFound(lbcBVehicle) >= 0 Then
        Exit Sub
    End If
    If Trim$(slName) = "" Then
        Exit Sub
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    'Create vehicle record
    tmVef.iCode = 0
    tmVef.sName = slName
    tmVef.iMnfVehGp3Mkt = 0
    tmVef.iMnfVehGp5Rsch = 0
    tmVef.sAddr(0) = ""
    tmVef.sAddr(1) = ""
    tmVef.sAddr(2) = ""
    tmVef.sPhone = ""
    tmVef.sFax = ""
    tmVef.sDialPos = ""
    tmVef.lPvfCode = 0
    tmVef.sUpdateRVF(0) = ""
    tmVef.sUpdateRVF(1) = ""
    tmVef.sUpdateRVF(2) = ""
    tmVef.sUpdateRVF(3) = ""
    tmVef.sUpdateRVF(4) = ""
    tmVef.sUpdateRVF(5) = ""
    tmVef.sUpdateRVF(6) = ""
    tmVef.sUpdateRVF(7) = ""
    'tmVef.sUpdateRVF(8) = ""
    'tmVef.sFormat = ""
    tmVef.iMnfVehGp4Fmt = 0
    tmVef.iMnfVehGp2 = 0
    tmVef.sType = "P"
    tmVef.sCodeStn = ""
    tmVef.iVefCode = 0
    tmVef.iCombineVefCode = 0
    tmVef.iOwnerMnfCode = 0
    tmVef.iProdPct(0) = 0
    tmVef.iProdPct(1) = 0
    tmVef.iProdPct(2) = 0
    tmVef.iProdPct(3) = 0
    tmVef.iProdPct(4) = 0
    tmVef.iProdPct(5) = 0
    tmVef.iProdPct(6) = 0
    tmVef.iProdPct(7) = 0
    'tmVef.iProdPct(8) = 0
    tmVef.sState = "A"
    tmVef.iMnfGroup(0) = 0
    tmVef.iMnfGroup(1) = 0
    tmVef.iMnfGroup(2) = 0
    tmVef.iMnfGroup(3) = 0
    tmVef.iMnfGroup(4) = 0
    tmVef.iMnfGroup(5) = 0
    tmVef.iMnfGroup(6) = 0
    tmVef.iMnfGroup(7) = 0
    'tmVef.iMnfGroup(8) = 0
    tmVef.iSort = 0
    tmVef.iDnfCode = 0
    tmVef.iReallDnfCode = 0
    tmVef.iMnfDemo = 0
    tmVef.iMnfSSCode(0) = 0
    tmVef.iMnfSSCode(1) = 0
    tmVef.iMnfSSCode(2) = 0
    tmVef.iMnfSSCode(3) = 0
    tmVef.iMnfSSCode(4) = 0
    tmVef.iMnfSSCode(5) = 0
    tmVef.iMnfSSCode(6) = 0
    tmVef.iMnfSSCode(7) = 0
    'tmVef.iMnfSSCode(8) = 0
    tmVef.sExportRAB = "N"
    tmVef.lVsfCode = 0
    tmVef.lRateAud = 0
    tmVef.lCPPCPM = 0
    tmVef.lYearAvails = 0
    tmVef.iPctSellout = 0
    tmVef.iMnfVehGp6Sub = 0
    tmVef.iNrfCode = 0
    tmVef.iSSMnfCode = 0
    tmVef.sStdPrice = ""
    tmVef.sStdInvTime = ""
    tmVef.sStdAlter = ""
    tmVef.iStdIndex = 0
    tmVef.sStdAlterName = ""
    tmVef.iRemoteID = tgUrf(0).iRemoteUserID
    tmVef.iAutoCode = tmVef.iCode
    ilRet = btrInsert(hmVef, tmVef, imVefRecLen, INDEXKEY0)
    Do
        tmVef.iRemoteID = tgUrf(0).iRemoteUserID
        tmVef.iAutoCode = tmVef.iCode
        'tmVef.iSourceID = tgUrf(0).iRemoteUserID
        'gPackDate slSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
        'gPackTime slSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
        ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    
    '11/26/17
    gFileChgdUpdate "vef.btr", True
    
    tgMVef(UBound(tgMVef)) = tmVef
    'ReDim Preserve tgMVef(1 To UBound(tgMVef) + 1) As VEF
    ReDim Preserve tgMVef(0 To UBound(tgMVef) + 1) As VEF
    'If UBound(tgMVef) > 2 Then
    If UBound(tgMVef) > 1 Then
        'ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
        ArraySortTyp fnAV(tgMVef(), 0), UBound(tgMVef), 0, LenB(tgMVef(0)), 0, -1, 0
    End If
    'Create Vpf
    ilRet = gVpfFind(SaleHist, tmVef.iCode)
    'lbcBVehicle.Clear
    'lbcBVehicleCode.Tag = ""
    ReDim tmBVehicleCode(0 To 0) As SORTCODE
    smBVehicleCodeTag = ""
    mBVehPop
    'Create Vpf
    ilRet = gVpfFind(SaleHist, tmVef.iCode)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDropDownKeyPress               *
'*                                                     *
'*             Created:5/11/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                      in transaction section         *
'*******************************************************
Private Function mDropDownKeyPress(KeyAscii As Integer, ilNegAllowed As Integer) As Integer
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcDropDown.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If ilNegAllowed Then
        If (KeyAscii = KEYNEG) And ((Len(edcDropDown.Text) = 0) Or (Len(edcDropDown.Text) = edcDropDown.SelLength)) Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) And (KeyAscii <> KEYNEG) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        Else
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
            Beep
            mDropDownKeyPress = False
            Exit Function
        End If
    End If
    slStr = edcDropDown.Text
    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
    If gCompAbsNumberStr(slStr, "999999999.99") > 0 Then
        Beep
        mDropDownKeyPress = False
        Exit Function
    End If
    'If KeyAscii <> KEYBACKSPACE Then
    '    Select Case Chr(KeyAscii)
    '        Case "7"
    '            llRowNo = 1
    '            ilColNo = 1
    '        Case "8"
    '            llRowNo = 1
    '            ilColNo = 2
    '        Case "9"
    '            llRowNo = 1
    '            ilColNo = 3
    '        Case "4"
    '            llRowNo = 2
    '            ilColNo = 1
    '        Case "5"
    '            llRowNo = 2
    '            ilColNo = 2
    '        Case "6"
    '            llRowNo = 2
    '            ilColNo = 3
    '        Case "1"
    '            llRowNo = 3
    '            ilColNo = 1
    '        Case "2"
    '            llRowNo = 3
    '            ilColNo = 2
    '        Case "3"
    '            llRowNo = 3
    '            ilColNo = 3
    '        Case "0"
    '            llRowNo = 4
    '            ilColNo = 1
    '        Case "00"   'Not possible
    '            llRowNo = 4
    '            ilColNo = 2
    '        Case "."
    '            llRowNo = 4
    '            ilColNo = 3
    '        Case "-"
    '            llRowNo = 0
    '    End Select
    '    If llRowNo > 0 Then
    '        flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX
    '        flY = fgPadMinY + (llRowNo - 1) * fgPadDeltaY
    '        imcNumOutLine.Move flX - 15, flY - 15
    '        imcNumOutLine.Visible = True
    '    Else
    '        imcNumOutLine.Visible = False
    '    End If
    'Else
    '    imcNumOutLine.Visible = False
    'End If
    mDropDownKeyPress = True
    Exit Function
End Function
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
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAgfCode As Integer
    Dim ilRet As Integer
    Dim slAgyRate As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    If (lmRowNo < vbcHist.Value) Or (lmRowNo >= vbcHist.Value + vbcHist.LargeChange + 1) Then
        'mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(CTINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcHist.Top + tmCtrls(CTINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True


    Select Case ilBoxNo
        Case CTINDEX
            pbcCT.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcHist, pbcCT, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imSave(1, lmRowNo) = -1 Then
                If lmRowNo > 1 Then
                    imSave(1, lmRowNo) = imSave(1, lmRowNo - 1)
                Else
                    imSave(1, lmRowNo) = 0  'Default to Cash
                End If
            End If
            pbcCT_Paint
            pbcCT.Visible = True
            pbcCT.SetFocus
        Case ADVTINDEX
            mAdvtPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAdvertiser.Height = gListBoxHeight(lbcAdvertiser.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 50
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAdvertiser.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcAdvertiser.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcAdvertiser.Move edcDropDown.Left, edcDropDown.Top - lbcAdvertiser.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(2, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                slStr = Trim$(smSave(2, lmRowNo - 1))
            End If
            If slStr = "" Then
                If imSelectedIndex > 1 Then
                    'Test if Direct Advertiser
                    slNameCode = lbcAdvtAgyCode.List(imSelectedIndex - 2)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If (InStr(slNameCode, "/Direct") > 0) Or (InStr(slNameCode, "/Non-Payee") > 0) Then
                        ilRet = gParseItem(slNameCode, 1, "/", slStr)
                    End If
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcAdvertiser
                If gLastFound(lbcAdvertiser) > 0 Then
                    lbcAdvertiser.ListIndex = gLastFound(lbcAdvertiser)
                Else
                    lbcAdvertiser.ListIndex = 1
                End If
            Else
                lbcAdvertiser.ListIndex = 1
            End If
            If lbcAdvertiser.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcAdvertiser.List(lbcAdvertiser.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODINDEX
            mGetAdvt Trim$(smSave(2, lmRowNo))
            mProdPop imAdfCode
            If imTerminate Then
                Exit Sub
            End If
            lbcProduct.Height = gListBoxHeight(lbcProduct.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 35
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProduct.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcProduct.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcProduct.Move edcDropDown.Left, edcDropDown.Top - lbcProduct.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(3, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If smSave(2, lmRowNo) = smSave(2, lmRowNo - 1) Then
                    slStr = Trim$(smSave(3, lmRowNo - 1))
                End If
            End If
            If slStr = "" Then
                If imAdfCode > 0 Then
                    slStr = tmAdf.sProduct
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcProduct
                If gLastFound(lbcProduct) > 0 Then
                    lbcProduct.ListIndex = gLastFound(lbcProduct)
                Else
                    lbcProduct.ListIndex = 1
                End If
            Else
                lbcProduct.ListIndex = 1
            End If
            If lbcProduct.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcProduct.List(lbcProduct.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AGYINDEX
            mAgencyPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAgency.Height = gListBoxHeight(lbcAgency.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 50
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAgency.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcAgency.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcAgency.Move edcDropDown.Left, edcDropDown.Top - lbcAgency.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(1, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If smSave(2, lmRowNo) = smSave(2, lmRowNo - 1) Then
                    slStr = Trim$(smSave(1, lmRowNo - 1))
                End If
            End If
            If slStr = "" Then
                If imSelectedIndex > 1 Then
                    'Test if Direct Advertiser
                    slNameCode = lbcAdvtAgyCode.List(imSelectedIndex - 2)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If (InStr(slNameCode, "/Direct") > 0) Or (InStr(slNameCode, "/Non-Payee") > 0) Then
                        slStr = lbcAgency.List(1)   'Direct
                    Else
                        ilAgfCode = Val(slCode)
                        For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                            slNameCode = tgAgency(ilLoop).sKey 'Traffic!lbcAgency.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = ilAgfCode Then
                                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                                Exit For
                            End If
                        Next ilLoop
                    End If
                Else
                    mGetAdvt Trim$(smSave(2, lmRowNo))
                    If imAdfCode > 0 Then
                        If tmAdf.iAgfCode > 0 Then
                            For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                                slNameCode = tgAgency(ilLoop).sKey 'Traffic!lbcAgency.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmAdf.iAgfCode Then
                                    ilRet = gParseItem(slNameCode, 1, "\", slStr)
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcAgency
                If gLastFound(lbcAgency) > 0 Then
                    lbcAgency.ListIndex = gLastFound(lbcAgency)
                Else
                    lbcAgency.ListIndex = 2
                End If
            Else
                lbcAgency.ListIndex = 2
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
        Case SPERSONINDEX
            mSPersonPop
            If imTerminate Then
                Exit Sub
            End If
            lbcSPerson.Height = gListBoxHeight(lbcSPerson.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 42
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSPerson.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcSPerson.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcSPerson.Move edcDropDown.Left, edcDropDown.Top - lbcSPerson.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(4, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If smSave(2, lmRowNo) = smSave(2, lmRowNo - 1) Then
                    slStr = Trim$(smSave(4, lmRowNo - 1))
                End If
            End If
            If slStr = "" Then
                mGetAdvt Trim$(smSave(2, lmRowNo))
                If imAdfCode > 0 Then
                    If tmAdf.iSlfCode > 0 Then
                        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
                            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmAdf.iSlfCode Then
                                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcSPerson
                If gLastFound(lbcSPerson) > 0 Then
                    lbcSPerson.ListIndex = gLastFound(lbcSPerson)
                Else
                    lbcSPerson.ListIndex = 1
                End If
            Else
                lbcSPerson.ListIndex = 1
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
        Case INVNOINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(5, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If (smSave(2, lmRowNo) = smSave(2, lmRowNo - 1)) And (smSave(1, lmRowNo) = smSave(1, lmRowNo - 1)) Then
                    slStr = Trim$(smSave(5, lmRowNo - 1))
                End If
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case CNTRNOINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 8
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(6, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If (smSave(2, lmRowNo) = smSave(2, lmRowNo - 1)) And (smSave(1, lmRowNo) = smSave(1, lmRowNo - 1)) Then
                    slStr = Trim$(smSave(6, lmRowNo - 1))
                End If
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case BVEHINDEX
            mBVehPop
            If imTerminate Then
                Exit Sub
            End If
            lbcBVehicle.Height = gListBoxHeight(lbcBVehicle.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcBVehicle.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(7, lmRowNo))
            If (slStr = "") Then
                slStr = Trim$(smSave(8, lmRowNo))
            End If
            If (slStr = "") And (lmRowNo > 1) Then
                If smSave(2, lmRowNo) = smSave(2, lmRowNo - 1) Then
                    slStr = Trim$(smSave(7, lmRowNo - 1))
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcBVehicle
                If gLastFound(lbcBVehicle) >= 0 Then
                    lbcBVehicle.ListIndex = gLastFound(lbcBVehicle)
                Else
                    lbcBVehicle.ListIndex = 0
                End If
            Else
                lbcBVehicle.ListIndex = 0
            End If
            If lbcBVehicle.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcBVehicle.List(lbcBVehicle.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AVEHINDEX
            mVehPop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(8, lmRowNo))
            If (slStr = "") Then
                slStr = Trim$(smSave(7, lmRowNo))
            End If
            If (slStr = "") And (lmRowNo > 1) Then
                If smSave(2, lmRowNo) = smSave(2, lmRowNo - 1) Then
                    slStr = Trim$(smSave(8, lmRowNo - 1))
                End If
            End If
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    lbcVehicle.ListIndex = 0
                End If
            Else
                lbcVehicle.ListIndex = 0
            End If
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
        Case PKLNINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(9, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                If (smSave(5, lmRowNo) = smSave(5, lmRowNo - 1)) Then
                    slStr = Trim$(smSave(9, lmRowNo - 1))
                End If
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case TRANDATEINDEX
            mTranTypePop
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            slStr = Trim$(smSave(10, lmRowNo))
            '6/28/12: TTP 5397
            If (slStr = "") And (lmRowNo > 1) Then
                If (smSave(2, lmRowNo) = smSave(2, lmRowNo - 1)) And (smSave(1, lmRowNo) = smSave(1, lmRowNo - 1)) Then
                    'slStr = gObtainEndStd(Format$(gDateValue(Trim$(smSave(10, lmRowNo - 1))) + 15, "m/d/yy"))
                    slStr = Trim$(smSave(10, lmRowNo - 1))
                End If
            End If
            If slStr = "" Then
                'Set to beginning of the year
                slStr = "1/15/" & Format$(gNow(), "yyyy")
                slStr = gObtainEndStd(slStr)
            End If
            edcDropDown.Text = slStr
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If Trim$(smSave(10, lmRowNo)) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case NTRTYPEINDEX  'NTR Type
            'mNTRTypePop    'tmAdf.iCode
            'If imTerminate Then
            '    Exit Sub
            'End If
            lbcNTRType.Height = gListBoxHeight(lbcNTRType.ListCount, 10)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 20  'tgSpf.iAProd
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcNTRType.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcNTRType.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcNTRType.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            imChgMode = True
            If imSave(2, lmRowNo) < 0 Then
                lbcNTRType.ListIndex = 0
                edcDropDown.Text = lbcNTRType.List(lbcNTRType.ListIndex)
            Else
                lbcNTRType.ListIndex = imSave(2, lmRowNo)
                edcDropDown.Text = lbcNTRType.List(lbcNTRType.ListIndex)
            End If
            imSave(2, lmRowNo) = lbcNTRType.ListIndex
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case NTRTAXINDEX

            If Not imTaxDefined Then
                Exit Sub
            End If
            lbcTax.Height = gListBoxHeight(lbcTax.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + tmCtrls(TRANTYPEINDEX).fBoxW + tmCtrls(GROSSINDEX).fBoxW + tmCtrls(NETINDEX).fBoxW + cmcDropDown.Width
            edcDropDown.MaxLength = 0
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(NTRTAXINDEX).fBoxX, tmCtrls(NTRTAXINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTax.ListIndex = imSave(4, lmRowNo)
            imChgMode = True
            If imSave(4, lmRowNo) < 0 Then
                lbcTax.ListIndex = 0
                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
            Else
                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
            End If
            imChgMode = False
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcTax.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcTax.Width, edcDropDown.Top + edcDropDown.Height
            Else
                lbcTax.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcTax.Width, edcDropDown.Top - lbcVehicle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcTax.Visible = True
            edcDropDown.SetFocus
        Case TRANTYPEINDEX
            'pbcCT.Width = tmCtrls(ilBoxNo).fBoxW
            'gMoveTableCtrl pbcHist, pbcTT, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            'If Trim$(smSave(13, lmRowNo)) = "" Then
            '    If rbcType(0).Value Then
            '        smSave(13, lmRowNo) = "HI"
            '    Else
            '        smSave(13, lmRowNo) = "IN"
            '    End If
            'End If
            'pbcTT_Paint
            'pbcTT.Visible = True
            'pbcTT.SetFocus
            lbcTranType.Height = gListBoxHeight(lbcTranType.ListCount, 10)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 20  'tgSpf.iAProd
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTranType.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If lmRowNo - vbcHist.Value <= vbcHist.LargeChange \ 2 Then
                lbcTranType.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcTranType.Move edcDropDown.Left, edcDropDown.Top - lbcTranType.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(13, lmRowNo))
            If (slStr = "") And (lmRowNo > 1) Then
                slStr = Trim$(smSave(13, lmRowNo - 1))
            End If
            If slStr <> "" Then
                gFindPartialMatch slStr, 0, 2, lbcTranType
                If gLastFound(lbcTranType) >= 0 Then
                    lbcTranType.ListIndex = gLastFound(lbcTranType)
                Else
                    lbcTranType.ListIndex = 0
                End If
            Else
                lbcTranType.ListIndex = 0
            End If
            If lbcTranType.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTranType.List(lbcTranType.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case GROSSINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 12
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(11, lmRowNo))
            If slStr = "" Then
                'Compute from Net if defined
                If Trim$(smSave(12, lmRowNo)) <> "" Then
                    mGetAgy Trim$(smSave(1, lmRowNo))
                    If imAgfCode <= 0 Then
                        slStr = Trim$(smSave(12, lmRowNo))
                    Else
                        slAgyRate = gDivStr(gSubStr("100.00", gIntToStrDec(tmAgf.iComm, 2)), "100.00")
                        slStr = gDivStr(Trim$(smSave(12, lmRowNo)), slAgyRate) '".85")
                    End If
                End If
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case NETINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 12
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(12, lmRowNo))
            If slStr = "" Then
                'Compute from Gross
                If Trim$(smSave(11, lmRowNo)) <> "" Then
                    mGetAgy Trim$(smSave(1, lmRowNo))
                    If imAgfCode <= 0 Then
                        slStr = Trim$(smSave(11, lmRowNo))
                    Else
                        slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                        slStr = gDivStr(gMulStr(Trim$(smSave(11, lmRowNo)), gSubStr("100.00", slAgyRate)), "100.00")
                    End If
                Else
                    If Trim$(smSave(13, lmRowNo)) = "PI" Then
                        slStr = "-"
                        imBypassFocus = True
                    End If
                End If
            End If
            edcDropDown.Text = slStr
            If imBypassFocus = True Then
                edcDropDown.SelStart = 1
            Else
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            End If
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case ACQCOSTINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 12
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(14, lmRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SSPARTINDEX
            mSSPartPop lmRowNo, 0, 0, 0, tgPhfRec(lmRowNo).iStatus, smSave(10, lmRowNo)
            lbcSSPart.Height = gListBoxHeight(lbcSSPart.ListCount, 4)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 41
            gMoveTableCtrl pbcHist, edcDropDown, tmCtrls(ilBoxNo).fBoxX - tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width, tmCtrls(ilBoxNo).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSSPart.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcSSPart.Width, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If imSave(3, lmRowNo) < 0 Then
                'If lbcSSPart.ListCount > 1 Then
                    lbcSSPart.ListIndex = 0
                'End If
                edcDropDown.Text = lbcSSPart.List(lbcSSPart.ListIndex)
            Else
'                If edcDropDown.Text <> lbcSSPart.List(imSave(3, lmRowNo)) Then
'                    edcDropDown.Text = lbcSSPart.List(imSave(3, lmRowNo))
'                Else
'                    edcDropDown_Change
'                End If
                lbcSSPart.ListIndex = imSave(3, lmRowNo)
                edcDropDown.Text = lbcSSPart.List(lbcSSPart.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAdvt                        *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get advertiser                *
'*                                                     *
'*******************************************************
Private Sub mGetAdvt(slAdvtName As String)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    'If Trim$(slAdvtName) = Trim$(tmAdf.sName) Then
    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
        slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
    Else
        slStr = Trim$(tmAdf.sName) & "/Direct"
    End If
    If (Trim$(slAdvtName) = Trim$(slStr)) And (imAdfCode > 0) Then
        Exit Sub
    End If
    gFindMatch slAdvtName, 2, lbcAdvertiser
    If gLastFound(lbcAdvertiser) > 1 Then
        slNameCode = tmAdvertiser(gLastFound(lbcAdvertiser) - 2).sKey    'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvertiser) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imAdfCode = Val(slCode)
        tmAdfSrchKey.iCode = imAdfCode
        If tmAdf.iCode <> imAdfCode Then
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            slStr = tmAdf.sProduct
        End If
    Else
        imAdfCode = 0
        tmAdf.sProduct = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAgy                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Agency                     *
'*                                                     *
'*******************************************************
Private Sub mGetAgy(slAgyName As String)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    gFindMatch slAgyName, 2, lbcAgency
    If gLastFound(lbcAgency) > 1 Then
        slNameCode = tgAgency(gLastFound(lbcAgency) - 2).sKey    'Traffic!lbcAgency.List(gLastFound(lbcAgency) - 2)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imAgfCode = Val(slCode)
        tmAgfSrchKey.iCode = imAgfCode
        If tmAgf.iCode <> imAgfCode Then
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
    Else
        imAgfCode = 0
        tmAgf.iComm = 0
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
    Dim ilLoop As Integer
    imTerminate = False
    imFirstActivate = True
    imcKey.Picture = IconTraf!imcKey.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    'SaleHist.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    mInitBox
    gCenterStdAlone SaleHist
    'SaleHist.Show
    Screen.MousePointer = vbHourglass
    ReDim tgPhfRec(0 To 1) As PHFREC
    tgPhfRec(1).iStatus = -1
    tgPhfRec(1).lRecPos = 0
    ReDim tgPhfDel(0 To 0) As PHFREC
    tgPhfDel(0).iStatus = -1
    tgPhfDel(0).lRecPos = 0
    ReDim smShow(0 To 18, 0 To 1) As String * 50 'Values shown in program area
    ReDim smSave(0 To 15, 0 To 1) As String * 60 'Values saved (program name) in program area
    ReDim imSave(0 To 4, 0 To 1) As Integer 'Values saved (program name) in program area
    ReDim lmSave(0 To 1, 0 To 1) As Long 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
        imSave(ilLoop, 1) = -1
    Next ilLoop
    For ilLoop = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
        lmSave(ilLoop, 1) = 0
    Next ilLoop
'    mInitDDE
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imWarningShown = False
    imPriceChgd = False
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imBoxNo = -1 'Initialize current Box to N/A
    lmRowNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imPhfChg = False
    imDragType = -1
    inGenMsgRet = False
    imBypassFocus = False
    imSettingValue = False
    imAdfCode = 0
    tmAdf.sName = "|"
    tmAgf.sName = "|"
    imAgfCode = 0
    imCalType = 0   'Standard
    imSelectedIndex = -1
    smNowDate = Format$(Now, "m/d/yy")
    lmSalesHistStartDate = gDateValue(gObtainYearStartDate(0, smNowDate))
    lmSalesHistEndDate = gDateValue(gObtainYearEndDate(0, smNowDate))
    rbcType(0).Value = True
    imBypassSetting = False
    hmRvf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rvf.Btr)", SaleHist
    On Error GoTo 0
    hmPhf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Phf.Btr)", SaleHist
    On Error GoTo 0
    imRvfPhfRecLen = Len(tmRvfPhf)
    hmPrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", SaleHist
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", SaleHist
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", SaleHist
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmVef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", SaleHist
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmSlf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.Btr)", SaleHist
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)
    hmSof = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sof.Btr)", SaleHist
    On Error GoTo 0
    imSofRecLen = Len(tmSof)
    hmSbf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", SaleHist
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)

    hmPif = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmPif, "", sgDBPath & "Pif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pif.Btr)", SaleHist
    On Error GoTo 0

    lbcTranType.Clear
    mTranTypePop
    If imTerminate Then
        Exit Sub
    End If
    If lbcTranType.ListCount <= 0 Then
        MsgBox "Transaction Types missing, please create from List Item Tran Type"
        imTerminate = True
        Exit Sub
    End If
    lbcNTRType.Clear 'Force list box to be populated
    mNTRTypePop
    If imTerminate Then
        Exit Sub
    End If
    If (Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR Then
        imTaxDefined = False
    Else
        imTaxDefined = True
        ilRet = gPopTaxRateBox(True, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
    End If
    If Not imTaxDefined Then
        ReDim tmTaxSortCode(0 To 0) As SORTCODE
    End If
    lbcTax.AddItem "[None]", 0
    'Sales Source
    ReDim tmSSPart(0 To 0) As SSPART
    smSMnfStamp = ""
    ilRet = gObtainMnfForType("S", smSMnfStamp, tmSMnf())
    smHMnfStamp = ""
    ilRet = gObtainMnfForType("H1", smHMnfStamp, tmHMnf())
    ilRet = gObtainVef()
    cbcSelect.Clear 'Force list box to be populated
    lbcAdvertiser.Clear 'Force list box to be populated
    lbcAgency.Clear 'Force list box to be populated
    mPopulate   'This will populate advt; agy and cbcSelect
    mSPersonPop
    If imTerminate Then
        Exit Sub
    End If
    mBVehPop
    If imTerminate Then
        Exit Sub
    End If
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    slDate = Format$(gNow(), "m/d/yy")
    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
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
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long
    Dim ilLoop As Integer

    flTextHeight = pbcHist.TextHeight("1") - 35
    'Position panel and picture areas with panel
    'plcSelect.Move 5715, 105
    plcHist.Move 135, 630, pbcHist.Width + fgPanelAdj + vbcHist.Width, pbcHist.Height + fgPanelAdj
    pbcHist.Move plcHist.Left + fgBevelX, plcHist.Top + fgBevelY
    vbcHist.Move pbcHist.Left + pbcHist.Width, pbcHist.Top
    pbcKey.Move plcHist.Left, plcHist.Top
    'Cash/Trade
    gSetCtrl tmCtrls(CTINDEX), 30, 375, 180, fgBoxGridH
    'Advertiser
    gSetCtrl tmCtrls(ADVTINDEX), 225, tmCtrls(CTINDEX).fBoxY, 900, fgBoxGridH
    'Product
    gSetCtrl tmCtrls(PRODINDEX), 1140, tmCtrls(CTINDEX).fBoxY, 750, fgBoxGridH
    'Agency
    gSetCtrl tmCtrls(AGYINDEX), 1905, tmCtrls(CTINDEX).fBoxY, 750, fgBoxGridH
    'Salesperson
    gSetCtrl tmCtrls(SPERSONINDEX), 2670, tmCtrls(CTINDEX).fBoxY, 420, fgBoxGridH
    'Invoice #
    gSetCtrl tmCtrls(INVNOINDEX), 3105, tmCtrls(CTINDEX).fBoxY, 405, fgBoxGridH
    'Contract #
    gSetCtrl tmCtrls(CNTRNOINDEX), 3525, tmCtrls(CTINDEX).fBoxY, 495, fgBoxGridH
    'Bill Vehicle
    gSetCtrl tmCtrls(BVEHINDEX), 4035, tmCtrls(CTINDEX).fBoxY, 495, fgBoxGridH
    'Air Vehicle
    gSetCtrl tmCtrls(AVEHINDEX), 4545, tmCtrls(CTINDEX).fBoxY, 480, fgBoxGridH
    'Package Line
    gSetCtrl tmCtrls(PKLNINDEX), 5040, tmCtrls(CTINDEX).fBoxY, 270, fgBoxGridH
    'Transaction Date
    gSetCtrl tmCtrls(TRANDATEINDEX), 5325, tmCtrls(CTINDEX).fBoxY, 570, fgBoxGridH
    'NTR Type
    gSetCtrl tmCtrls(NTRTYPEINDEX), 5910, tmCtrls(CTINDEX).fBoxY, 300, fgBoxGridH
    'NTR Tax
    gSetCtrl tmCtrls(NTRTAXINDEX), 6225, tmCtrls(CTINDEX).fBoxY, 180, fgBoxGridH
    'Transaction Type
    gSetCtrl tmCtrls(TRANTYPEINDEX), 6420, tmCtrls(CTINDEX).fBoxY, 180, fgBoxGridH
    'Gross
    gSetCtrl tmCtrls(GROSSINDEX), 6615, tmCtrls(CTINDEX).fBoxY, 735, fgBoxGridH
    'Net
    gSetCtrl tmCtrls(NETINDEX), 7365, tmCtrls(CTINDEX).fBoxY, 735, fgBoxGridH
    'Acquisition Cost
    gSetCtrl tmCtrls(ACQCOSTINDEX), 8115, tmCtrls(CTINDEX).fBoxY, 495, fgBoxGridH
    'SS/Participant
    gSetCtrl tmCtrls(SSPARTINDEX), 8625, tmCtrls(CTINDEX).fBoxY, 255, fgBoxGridH
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop


    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
        Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    pbcHist.Picture = LoadPicture("")
    pbcHist.Width = llMax
    plcHist.Width = llMax + vbcHist.Width + 2 * fgBevelX + 15
    lacFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    '1/19/10:  Make Report invisible
    If cmcYear.Visible Then
        'cmcDone.Left = (SaleHist.Width - 5 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcDone.Left = (SaleHist.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    Else
        'cmcDone.Left = (SaleHist.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcDone.Left = (SaleHist.Width - 3 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    End If
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcSave.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    'cmcReport.Left = cmcSave.Left + cmcSave.Width + ilSpaceBetweenButtons
    'cmcYear.Left = cmcReport.Left + cmcReport.Width + ilSpaceBetweenButtons
    cmcYear.Left = cmcSave.Left + cmcSave.Width + ilSpaceBetweenButtons
    cmcDone.Top = SaleHist.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    'cmcReport.Top = cmcDone.Top
    cmcYear.Top = cmcDone.Top
    imcTrash.Top = cmcDone.Top - imcTrash.Height / 2
    imcTrash.Left = SaleHist.Width - (3 * imcTrash.Width) / 2
    llAdjTop = imcTrash.Top - plcType.Top - plcType.Height - 120 - tmCtrls(1).fBoxY
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcHist.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcHist.Height = llAdjTop + 2 * fgBevelY
    pbcHist.Left = plcHist.Left + fgBevelX
    pbcHist.Top = plcHist.Top + fgBevelY
    pbcHist.Height = plcHist.Height - 2 * fgBevelY
    vbcHist.Left = pbcHist.Left + pbcHist.Width + 15
    vbcHist.Top = pbcHist.Top
    vbcHist.Height = pbcHist.Height
    ''cbcSelect.Left = plcHist.Left + plcHist.Width - cbcSelect.Width
    'edcCntrNo.Left = plcHist.Left + plcHist.Width - edcCntrNo.Width
    'lacCntrNo.Left = edcCntrNo.Left - lacCntrNo.Width - 120
    'cbcSelect.Left = lacCntrNo.Left - cbcSelect.Width - 120
    frcSelection.Left = plcHist.Left + plcHist.Width - frcSelection.Width
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(llRowNo As Long)
    Dim ilLoop As Integer

    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, llRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
        imSave(ilLoop, llRowNo) = -1
    Next ilLoop
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, llRowNo) = ""
    Next ilLoop
    tgPhfRec(llRowNo).iStatus = 0
    tgPhfRec(llRowNo).lRecPos = 0
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
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llRowNo As Long
    Dim ilBoxNo As Integer
    For llRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        For ilBoxNo = CTINDEX To SSPARTINDEX Step 1
            Select Case ilBoxNo
                Case CTINDEX
                    If imSave(1, llRowNo) = 0 Then
                        slStr = "C"
                    ElseIf imSave(1, llRowNo) = 1 Then
                        slStr = "T"
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case ADVTINDEX
                    slStr = Trim$(smSave(2, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case PRODINDEX
                    slStr = Trim$(smSave(3, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case AGYINDEX
                    slStr = Trim$(smSave(1, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case SPERSONINDEX
                    slStr = Trim$(smSave(4, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case INVNOINDEX
                    slStr = Trim$(smSave(5, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case CNTRNOINDEX
                    slStr = Trim$(smSave(6, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case BVEHINDEX
                    slStr = Trim$(smSave(7, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case AVEHINDEX
                    slStr = Trim$(smSave(8, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case PKLNINDEX
                    slStr = Trim$(smSave(9, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case TRANDATEINDEX
                    slStr = Trim$(smSave(10, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case NTRTYPEINDEX
                    If imSave(2, llRowNo) <= 0 Then
                        slStr = ""
                    Else
                        slStr = lbcNTRType.List(imSave(2, llRowNo))
                    End If
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case NTRTAXINDEX
                    'If (lmSave(1, llRowNo) > 0) Or (Not imTaxDefined) Then
                    If (Not imTaxDefined) Then
                        slStr = ""
                    Else
                        If imSave(4, llRowNo) = -1 Then
                            slStr = ""
                        ElseIf imSave(4, llRowNo) = 0 Then
                            slStr = "N"
                        Else
                            slStr = "Y"
                        End If
                    End If
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case TRANTYPEINDEX
                    slStr = Trim$(smSave(13, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case GROSSINDEX
                    slStr = Trim$(smSave(11, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case NETINDEX
                    slStr = Trim$(smSave(12, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case ACQCOSTINDEX
                    slStr = Trim$(smSave(14, llRowNo))
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
                Case SSPARTINDEX
                    mSSPartPop llRowNo, 0, 0, 0, tgPhfRec(llRowNo).iStatus, smSave(10, llRowNo)
                    If imSave(3, llRowNo) <= 0 Then
                        slStr = ""
                    Else
                        slStr = lbcSSPart.List(imSave(3, llRowNo))
                    End If
                    gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, llRowNo) = tmCtrls(ilBoxNo).sShow
            End Select
        Next ilBoxNo
    Next llRowNo
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slTotalTaxAmt                 ilTrfVef                      ilVef                     *
'*  ilTaxable                                                                             *
'******************************************************************************************

'
'   mMoveCtrlToRec
'   Where:
'
    Dim llLoop As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRowNo As Long
    Dim slStr As String
    Dim ilTax1 As Integer       '1-17-02 tax 1 applicable?
    Dim ilTax2 As Integer       '1-17-02 tax 2 applicable?
    Dim slTax1Pct As String        '1-17-02 tax 1 pct value
    Dim slTax2Pct As String        '1-17-02 tax 2 pct value
    Dim slBothTaxPct As String   '1-17-02 total tax1 & tax2 values
    Dim slNet As String         '1-17-02 net input (net + taxes input)
    Dim slAgyNet As String      '1-17-02 calc agy net amount
    Dim slTax1Amt As String         '1-18-02
    Dim slTax2Amt As String         '1-18-02

    Dim slDollar As String
    Dim ilTrfAgyAdvt As Integer
    Dim llTax1Rate As Long
    Dim llTax2Rate As Long
    Dim ilSbfTrfCode As Integer
    Dim slTemp As String
    Dim llTax1 As Long
    Dim llTax2 As Long
    Dim slGrossNet As String

    'For llRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
    For llRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        'Set Cash/Trade
        If imSave(1, llRowNo) = 1 Then
            tgPhfRec(llRowNo).tPhf.sCashTrade = "T"
        Else
            tgPhfRec(llRowNo).tPhf.sCashTrade = "C"
        End If
        'Set Agency
        If Trim$(smSave(1, llRowNo)) = "[Direct]" Then
            tgPhfRec(llRowNo).tPhf.iAgfCode = 0
        Else
            gFindMatch Trim$(smSave(1, llRowNo)), 2, lbcAgency
            If gLastFound(lbcAgency) > 1 Then
                slNameCode = tgAgency(gLastFound(lbcAgency) - 2).sKey    'Traffic!lbcAgency.List(gLastFound(lbcAgency) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgPhfRec(llRowNo).tPhf.iAgfCode = Val(slCode)
            End If
        End If
        'Set Advertiser
        gFindMatch Trim$(smSave(2, llRowNo)), 2, lbcAdvertiser
        If gLastFound(lbcAdvertiser) > 1 Then
            slNameCode = tmAdvertiser(gLastFound(lbcAdvertiser) - 2).sKey    'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvertiser) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iAdfCode = Val(slCode)
            'Product
            tgPhfRec(llRowNo).tPhf.lPrfCode = mAddProd(Trim$(smSave(3, llRowNo)), tgPhfRec(llRowNo).tPhf.iAdfCode)
            tgPhfRec(llRowNo).sProduct = Trim$(smSave(3, llRowNo))
        Else
            tgPhfRec(llRowNo).tPhf.iAdfCode = 0
            'Product
            tgPhfRec(llRowNo).tPhf.lPrfCode = 0
            tgPhfRec(llRowNo).sProduct = ""
        End If
        'Set Salesperson Name
        gFindMatch Trim$(smSave(4, llRowNo)), 2, lbcSPerson
        If gLastFound(lbcSPerson) > 1 Then
            slNameCode = tgSalesperson(gLastFound(lbcSPerson) - 2).sKey  'Traffic!lbcSalesperson.List(gLastFound(lbcSPerson) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iSlfCode = Val(slCode)
        Else
            tgPhfRec(llRowNo).tPhf.iSlfCode = 0
        End If
        'Set Invoice #
        If Trim$(smSave(5, llRowNo)) <> "" Then
            tgPhfRec(llRowNo).tPhf.lInvNo = Val(Trim$(smSave(5, llRowNo)))
        Else
            tgPhfRec(llRowNo).tPhf.lInvNo = 0
        End If
        'Set Contract #
        tgPhfRec(llRowNo).tPhf.lCntrNo = Val(Trim$(smSave(6, llRowNo)))
        If (tgPhfRec(llRowNo).iStatus = 0) Then  'New selected
            'Reference #
            tgPhfRec(llRowNo).tPhf.lRefInvNo = 0
        End If
        'Set Bill Vehicle Name
        gFindMatch Trim$(smSave(7, llRowNo)), 0, lbcBVehicle
        If gLastFound(lbcBVehicle) >= 0 Then
            slNameCode = tmBVehicleCode(gLastFound(lbcBVehicle)).sKey 'lbcBVehicleCode.List(gLastFound(lbcBVehicle))
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iBillVefCode = Val(slCode)
        Else
            tgPhfRec(llRowNo).tPhf.iBillVefCode = 0
        End If
        'Set Air Vehicle Name
        gFindMatch Trim$(smSave(8, llRowNo)), 0, lbcVehicle
        If gLastFound(lbcVehicle) >= 0 Then
            slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iAirVefCode = Val(slCode)
        Else
            tgPhfRec(llRowNo).tPhf.iAirVefCode = 0
        End If
        'Check #
        If (tgPhfRec(llRowNo).iStatus = 0) Then  'New selected
            '6/7/15: Check number changed to string
            'tgPhfRec(llRowNo).tPhf.lCheckNo = 0
            tgPhfRec(llRowNo).tPhf.sCheckNo = ""
        End If
'        'Transaction Type
'        If rbcType(0).Value Then
'            tgPhfRec(llRowNo).tPhf.sTranType = "HI"
'        Else
'            'If InStr(Trim$(smSave(12, llRowNo)), "-") > 0 Then
'            'If (InStr(smSave(12, llRowNo), "-") > 0) And (gStrDecToLong(smSave(11, llRowNo), 2) = 0) Then
'            If (gStrDecToLong(smSave(12, llRowNo), 2) <> 0) And (gStrDecToLong(smSave(11, llRowNo), 2) = 0) Then
'                tgPhfRec(llRowNo).tPhf.sTranType = "PI"
'            Else
'                tgPhfRec(llRowNo).tPhf.sTranType = "IN"
'            End If
'        End If
        If (tgPhfRec(llRowNo).iStatus = 0) Then  'New selected
            tgPhfRec(llRowNo).tPhf.sTranType = smSave(13, llRowNo)
            tgPhfRec(llRowNo).tPhf.sAction = ""
        End If
        'Set Invoice #
        tgPhfRec(llRowNo).tPhf.iPkLineNo = Val(Trim$(smSave(9, llRowNo)))
        'Transaction Date
        If (tgPhfRec(llRowNo).iStatus = 0) Then  'New selected
            gPackDate Trim$(smSave(10, llRowNo)), tgPhfRec(llRowNo).tPhf.iTranDate(0), tgPhfRec(llRowNo).tPhf.iTranDate(1)
            tgPhfRec(llRowNo).tPhf.iAgePeriod = Month(Trim$(smSave(10, llRowNo)))
            tgPhfRec(llRowNo).tPhf.iAgingYear = Year(gDateValue(Trim$(smSave(10, llRowNo))))
            gPackDate "", tgPhfRec(llRowNo).tPhf.iPurgeDate(0), tgPhfRec(llRowNo).tPhf.iPurgeDate(1)
        Else
            gUnpackDate tgPhfRec(llRowNo).tPhf.iTranDate(0), tgPhfRec(llRowNo).tPhf.iTranDate(1), slStr
            If gDateValue(slStr) <> gDateValue(Trim$(smSave(10, llRowNo))) Then
                gPackDate Trim$(smSave(10, llRowNo)), tgPhfRec(llRowNo).tPhf.iTranDate(0), tgPhfRec(llRowNo).tPhf.iTranDate(1)
                tgPhfRec(llRowNo).tPhf.iAgePeriod = Month(Trim$(smSave(10, llRowNo)))
                tgPhfRec(llRowNo).tPhf.iAgingYear = Year(gDateValue(Trim$(smSave(10, llRowNo))))
            End If
        End If
        'NTR
        tgPhfRec(llRowNo).tPhf.lSbfCode = lmSave(1, llRowNo)
        tgPhfRec(llRowNo).tPhf.iMnfItem = 0
        If imSave(2, llRowNo) > 0 Then
            slNameCode = tmNTRTypeCode(imSave(2, llRowNo) - 1).sKey  'Traffic!lbcUserVehicle.List(llLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iMnfItem = Val(slCode)
        End If
        'NTR Tax
        tgPhfRec(llRowNo).tPhf.iBacklogTrfCode = 0
        If (tgPhfRec(llRowNo).tPhf.iMnfItem > 0) And (imSave(4, llRowNo) > 0) And (tgPhfRec(llRowNo).tPhf.lSbfCode <= 0) Then
            tgPhfRec(llRowNo).tPhf.iBacklogTrfCode = lbcTax.ItemData(imSave(4, llRowNo))
        End If
        'Set Gross
        If (tgPhfRec(llRowNo).tPhf.sTranType = "PI") Or (tgPhfRec(llRowNo).tPhf.sTranType = "PO") Or (Left$(tgPhfRec(llRowNo).tPhf.sTranType, 1) = "W") Then
            gStrToPDN "", 2, 6, tgPhfRec(llRowNo).tPhf.sGross
        Else
            gStrToPDN Trim$(smSave(11, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sGross
        End If
        'Participant-Set without Participant defined (reports will create the participants)
        tgPhfRec(llRowNo).tPhf.iMnfGroup = 0
        mSSPartPop llRowNo, 0, 0, 0, tgPhfRec(llRowNo).iStatus, smSave(10, llRowNo)
        If imSave(3, llRowNo) = 0 Then
            If tgPhfRec(llRowNo).iStatus = 0 Then
                'Set for Auto-Split
                tgPhfRec(llRowNo).tPhf.iMnfGroup = -1
            End If
        ElseIf imSave(3, llRowNo) > 0 Then
            'slNameCode = tmSSPart(imSave(3, llRowNo) - 1).sKey  'Traffic!lbcUserVehicle.List(llLoop)
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPhfRec(llRowNo).tPhf.iMnfGroup = tmSSPart(imSave(3, llRowNo) - 1).iMnfGroup   '(slCode)
        End If
        tgPhfRec(llRowNo).tPhf.lPcfCode = 0
        'Set Net & taxes if applicable
        '1-17-02 determine if taxes are applicable for this agency, then backcompute to get net and taxes
        '12/17/06-Change to tax by agency or vehicle
        'If tgSpf.iBTax(0) <> 0 Or tgSpf.iBTax(1) <> 0 Then
        llTax1 = tgPhfRec(llRowNo).tPhf.lTax1
        llTax2 = tgPhfRec(llRowNo).tPhf.lTax2
        tgPhfRec(llRowNo).tPhf.lTax1 = 0
        tgPhfRec(llRowNo).tPhf.lTax2 = 0
        If tgPhfRec(llRowNo).tPhf.sCashTrade = "C" Then
            If (tgPhfRec(llRowNo).tPhf.sTranType <> "PO") Then
                If (tgPhfRec(llRowNo).tPhf.sTranType = "PI") Or (Left$(tgPhfRec(llRowNo).tPhf.sTranType, 1) = "W") Or (tgPhfRec(llRowNo).iStatus <> 0) Then
                    'Compute tax as proportional change
                    gPDNToStr tgPhfRec(llRowNo).tPhf.sNet, 2, slStr
                    slStr = gAddStr(slStr, gLongToStrDec(llTax1 + llTax2, 2))
                    If gStrDecToLong(slStr, 2) <> gStrDecToLong(smSave(12, llRowNo), 2) Then
                        If llTax1 <> 0 Then
                            slTemp = gLongToStrDec(llTax1, 2)
                            slTax1Amt = gDivStr(gMulStr(smSave(12, llRowNo), slTemp), slStr)
                            tgPhfRec(llRowNo).tPhf.lTax1 = gStrDecToLong(slTax1Amt, 2)
                        End If
                        If llTax2 <> 0 Then
                            slTemp = gLongToStrDec(llTax2, 2)
                            slTax2Amt = gDivStr(gMulStr(smSave(12, llRowNo), slTemp), slStr)
                            tgPhfRec(llRowNo).tPhf.lTax2 = gStrDecToLong(slTax2Amt, 2)
                        End If
                        slStr = gSubStr(smSave(12, llRowNo), gLongToStrDec(tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2, 2))
                        gStrToPDN Trim$(slStr), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                    Else
                        tgPhfRec(llRowNo).tPhf.lTax1 = llTax1
                        tgPhfRec(llRowNo).tPhf.lTax2 = llTax2
                        slStr = gSubStr(smSave(12, llRowNo), gLongToStrDec(tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2, 2))
                        gStrToPDN Trim$(slStr), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                    End If
                Else
                    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                        'get the agency or direct advertiser
                        ilTax1 = False
                        ilTax2 = False
                        slTax1Pct = ".00"
                        slTax2Pct = ".00"
                        slBothTaxPct = "1.00"
                        ilTrfAgyAdvt = gGetTrfIndexForAgyAdvt(tgPhfRec(llRowNo).tPhf.iAdfCode, tgPhfRec(llRowNo).tPhf.iAgfCode)

                        If tgPhfRec(llRowNo).tPhf.iMnfItem > 0 Then
                            '12/24/06:  Will add Tax list box selection
                            ilSbfTrfCode = 0
                            If tgPhfRec(llRowNo).tPhf.lSbfCode > 0 Then
                                tmSbfSrchKey1.lCode = tgPhfRec(llRowNo).tPhf.lSbfCode
                                ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                If (ilRet = BTRV_ERR_NONE) Then
                                    ilSbfTrfCode = tmSbf.iTrfCode
                                End If
                            Else
                                ilSbfTrfCode = tgPhfRec(llRowNo).tPhf.iBacklogTrfCode
                            End If
                            gGetNTRTaxRates ilSbfTrfCode, llTax1Rate, llTax2Rate, slGrossNet
                        Else
                            If ilTrfAgyAdvt <> -1 Then
                                gGetAirTimeTaxValues ilTrfAgyAdvt, tgPhfRec(llRowNo).tPhf.iAirVefCode, llTax1Rate, llTax2Rate, slGrossNet
                            Else
                                llTax1Rate = 0
                                llTax2Rate = 0
                            End If
                        End If
                        If llTax1Rate > 0 Or llTax2Rate > 0 Then
                            If llTax1Rate > 0 Then
                                ilTax1 = True
                            End If
                            slTax1Pct = gLongToStrDec(llTax1Rate, 4)
                            slBothTaxPct = gAddStr(slBothTaxPct, slTax1Pct)
                            If llTax2Rate > 0 Then
                                ilTax2 = True
                            End If
                            slTax2Pct = gLongToStrDec(llTax2Rate, 4)
                            slBothTaxPct = gAddStr(slBothTaxPct, slTax2Pct)
                        End If
                        If ilTax1 Or ilTax2 Then
                            If ((Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA) Or ((Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA) Then
                                'If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                                If slGrossNet <> "N" Then
                                    slNet = smSave(12, llRowNo)
                                    'If tgPhfRec(llRowNo).tPhf.iAgfCode > 0 Then
                                    '    If tgPhfRec(llRowNo).tPhf.iAgfCode <> tmAgf.iCode Then  'prevent extra read if already in mem.
                                    '        tmAgfSrchKey.iCode = tgPhfRec(llRowNo).tPhf.iAgfCode
                                    '        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                                    '    End If
                                    '    slAgyNet = gDivStr(slNet, gSubStr(slBothTaxPct, gIntToStrDec(tmAgf.iComm, 4)))
                                    'Else
                                    '    slAgyNet = gDivStr(slNet, slBothTaxPct)
                                    'End If
                                    'slAgyNet = gSubStr(slNet, gLongToStrDec(llTax1Rate + llTax2Rate, 2))
                                    slDollar = smSave(11, llRowNo)
                                Else
                                    slNet = smSave(12, llRowNo)
                                    slDollar = gSubStr(slNet, smSave(15, llRowNo))
                                End If
                                'gStrToPDN slAgyNet, 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                                'slTotalTaxAmt = gSubStr(slNet, slAgyNet)
                                If (ilTax1) And (ilTax2) Then       'both taxes apply
                                    slTax1Amt = gDivStr(gMulStr(slDollar, slTax1Pct), "100.")
                                    tgPhfRec(llRowNo).tPhf.lTax1 = gStrDecToLong(slTax1Amt, 2)
                                    slTax2Amt = gDivStr(gMulStr(slDollar, slTax2Pct), "100.")
                                    tgPhfRec(llRowNo).tPhf.lTax2 = gStrDecToLong(slTax2Amt, 2)
                                ElseIf ilTax1 Then                  'only sales tax1
                                    slTax1Amt = gDivStr(gMulStr(slDollar, slTax1Pct), "100.")
                                    tgPhfRec(llRowNo).tPhf.lTax1 = gStrDecToLong(slTax1Amt, 2)
                                ElseIf ilTax2 Then                  'only sales tax2
                                    slTax2Amt = gDivStr(gMulStr(slDollar, slTax2Pct), "100.")
                                    tgPhfRec(llRowNo).tPhf.lTax2 = gStrDecToLong(slTax2Amt, 2)
                                End If
                                slAgyNet = gSubStr(slNet, gLongToStrDec(tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2, 2))
                                gStrToPDN slAgyNet, 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                            Else
                                gStrToPDN Trim$(smSave(12, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                            End If
                        Else
                            gStrToPDN Trim$(smSave(12, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                        End If
                    Else
                        gStrToPDN Trim$(smSave(12, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
                    End If
                End If
            Else
                gStrToPDN Trim$(smSave(12, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
            End If
        Else
            gStrToPDN Trim$(smSave(12, llRowNo)), 2, 6, tgPhfRec(llRowNo).tPhf.sNet
        End If
        smSave(15, llRowNo) = gLongToStrDec(tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2, 2)
        tgPhfRec(llRowNo).tPhf.lAcquisitionCost = gStrDecToLong(smSave(14, llRowNo), 2)
        'Invoice date match transaction date
        tgPhfRec(llRowNo).tPhf.iInvDate(0) = tgPhfRec(llRowNo).tPhf.iTranDate(0)
        tgPhfRec(llRowNo).tPhf.iInvDate(1) = tgPhfRec(llRowNo).tPhf.iTranDate(1)
    Next llRowNo
    For llRowNo = LBONE To UBound(tgPhfRec) - 1 Step 1
        'If PI get Ageing Information from IN
        If tgPhfRec(llRowNo).tPhf.sTranType = "PI" Then
            For llLoop = LBONE To UBound(tgPhfRec) - 1 Step 1
                If (tgPhfRec(llLoop).tPhf.sTranType = "IN") And (tgPhfRec(llRowNo).tPhf.lInvNo = tgPhfRec(llLoop).tPhf.lInvNo) Then
                    tgPhfRec(llRowNo).tPhf.iAgePeriod = tgPhfRec(llLoop).tPhf.iAgePeriod
                    tgPhfRec(llRowNo).tPhf.iAgingYear = tgPhfRec(llLoop).tPhf.iAgingYear
                    tgPhfRec(llRowNo).tPhf.iInvDate(0) = tgPhfRec(llLoop).tPhf.iTranDate(0)
                    tgPhfRec(llRowNo).tPhf.iInvDate(1) = tgPhfRec(llLoop).tPhf.iTranDate(1)
                    Exit For
                End If
            Next llLoop
        End If
    Next llRowNo
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRowNo As Long
    Dim llUpper As Long
    Dim llTaxes As Long           '1-17-02
    Dim ilTax As Integer
    Dim ilTrfIndex As Integer

    llUpper = UBound(tgPhfRec)
    ReDim smShow(0 To 18, 0 To llUpper) As String * 50 'Values shown in program area
    ReDim smSave(0 To 15, 0 To llUpper) As String * 60 'Values saved (program name) in program area
    ReDim imSave(0 To 4, 0 To llUpper) As Integer 'Values saved (program name) in program area
    ReDim lmSave(0 To 1, 0 To llUpper) As Long 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, llUpper) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, llUpper) = ""
    Next ilLoop
    For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
        imSave(ilLoop, llUpper) = -1
    Next ilLoop
    For ilLoop = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
        lmSave(ilLoop, llUpper) = 0
    Next ilLoop
    'Init value in the case that no records are associated with the salesperson
    If llUpper = LBONE Then
        llRowNo = lmRowNo
        lmRowNo = 1
        mInitNew lmRowNo
        lmRowNo = llRowNo
    End If
    For llRowNo = LBONE To UBound(tgPhfRec) - 1 Step 1
        'Get Cash/Trade
        If tgPhfRec(llRowNo).tPhf.sCashTrade <> "T" Then
            imSave(1, llRowNo) = 0
        Else
            imSave(1, llRowNo) = 1
        End If
        'Get Agency Name
        If tgPhfRec(llRowNo).tPhf.iAgfCode > 0 Then
            For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                slNameCode = tgAgency(ilLoop).sKey 'Traffic!lbcAgency.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tgPhfRec(llRowNo).tPhf.iAgfCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smSave(1, llRowNo))
                    Exit For
                End If
            Next ilLoop
        Else
            smSave(1, llRowNo) = "[Direct]"
        End If
        'Get Advertiser Name
        smSave(2, llRowNo) = "[None]"
        For ilLoop = 0 To UBound(tmAdvertiser) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
            slNameCode = tmAdvertiser(ilLoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tgPhfRec(llRowNo).tPhf.iAdfCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smSave(2, llRowNo))
                Exit For
            End If
        Next ilLoop
        'Get product
        smSave(3, llRowNo) = Trim$(tgPhfRec(llRowNo).sProduct)
        If Trim$(smSave(3, llRowNo)) = "" Then
            smSave(3, llRowNo) = "[None]"
        End If
        'Get Salesperson Name
        smSave(4, llRowNo) = "[None]"
        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tgPhfRec(llRowNo).tPhf.iSlfCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smSave(4, llRowNo))
                Exit For
            End If
        Next ilLoop
        'Get Invoice Number
        smSave(5, llRowNo) = Trim$(str$(tgPhfRec(llRowNo).tPhf.lInvNo))
        'Get Contract Number
        smSave(6, llRowNo) = Trim$(str$(tgPhfRec(llRowNo).tPhf.lCntrNo))
        'Get Bill Vehicle Name
        smSave(7, llRowNo) = ""
        For ilLoop = 0 To UBound(tmBVehicleCode) - 1 Step 1  'lbcBVehicleCode.ListCount - 1 Step 1
            slNameCode = tmBVehicleCode(ilLoop).sKey    'lbcBVehicleCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tgPhfRec(llRowNo).tPhf.iBillVefCode Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", smSave(7, llRowNo))
                Exit For
            End If
        Next ilLoop
        'Get Air Vehicle Name
        smSave(8, llRowNo) = ""
        For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tgPhfRec(llRowNo).tPhf.iAirVefCode Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", smSave(8, llRowNo))
                Exit For
            End If
        Next ilLoop
        'Get Package Line # (9)
        If tgPhfRec(llRowNo).tPhf.iPkLineNo > 0 Then
            smSave(9, llRowNo) = Trim$(str$(tgPhfRec(llRowNo).tPhf.iPkLineNo))
        Else
            smSave(9, llRowNo) = ""
        End If
        'Get Transaction Date
        gUnpackDate tgPhfRec(llRowNo).tPhf.iTranDate(0), tgPhfRec(llRowNo).tPhf.iTranDate(1), smSave(10, llRowNo)
        lmSave(1, llRowNo) = tgPhfRec(llRowNo).tPhf.lSbfCode
        'NTR Type
        imSave(2, llRowNo) = -1
        If tgPhfRec(llRowNo).tPhf.iMnfItem > 0 Then
            For ilLoop = 0 To UBound(tmNTRTypeCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                slNameCode = tmNTRTypeCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tgPhfRec(llRowNo).tPhf.iMnfItem Then
                    imSave(2, llRowNo) = ilLoop + 1
                    Exit For
                End If
            Next ilLoop
        End If
        'NTR Tax
        'If (Not imTaxDefined) Or (tgPhfRec(llRowNo).tPhf.lSbfCode > 0) Or (tgPhfRec(llRowNo).tPhf.iMnfItem <= 0) Then
        imSave(4, llRowNo) = -1
        If (tgPhfRec(llRowNo).tPhf.iMnfItem <= 0) And ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Then
            'Air Time
            If (tgPhfRec(llRowNo).tPhf.lTax1 > 0) Or (tgPhfRec(llRowNo).tPhf.lTax2 > 0) Then
                ilTrfIndex = -1
                If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                    ilTrfIndex = gGetTrfIndexForAgyAdvt(tgPhfRec(llRowNo).tPhf.iAdfCode, tgPhfRec(llRowNo).tPhf.iAgfCode)
                ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
                    ilTrfIndex = gGetTrfIndexForVeh(tgPhfRec(llRowNo).tPhf.iAirVefCode)
                End If
                If ilTrfIndex <> -1 Then
                    For ilTax = 0 To UBound(tmTaxSortCode) - 1 Step 1
                        slNameCode = tmTaxSortCode(ilTax).sKey
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If tgTrf(ilTrfIndex).iCode = Val(slCode) Then
                            imSave(4, llRowNo) = ilTax + 1
                            Exit For
                        End If
                    Next ilTax
                End If
            End If
        ElseIf (tgPhfRec(llRowNo).tPhf.iMnfItem > 0) And ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
            'NTR
            If tgPhfRec(llRowNo).tPhf.iBacklogTrfCode > 0 Then
                For ilTax = 0 To UBound(tmTaxSortCode) - 1 Step 1
                    slNameCode = tmTaxSortCode(ilTax).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tgPhfRec(llRowNo).tPhf.iBacklogTrfCode = Val(slCode) Then
                        imSave(4, llRowNo) = ilTax + 1
                        Exit For
                    End If
                Next ilTax
            Else
                If tgPhfRec(llRowNo).tPhf.lSbfCode > 0 Then
                    tmSbfSrchKey1.lCode = tgPhfRec(llRowNo).tPhf.lSbfCode
                    ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) Then
                        For ilTax = 0 To UBound(tmTaxSortCode) - 1 Step 1
                            slNameCode = tmTaxSortCode(ilTax).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If tmSbf.iTrfCode = Val(slCode) Then
                                imSave(4, llRowNo) = ilTax + 1
                                Exit For
                            End If
                        Next ilTax
                    End If
                End If
            End If
        End If
        smSave(15, llRowNo) = gLongToStrDec(tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2, 2)
        smSave(13, llRowNo) = tgPhfRec(llRowNo).tPhf.sTranType
        'Get Gross
        gPDNToStr tgPhfRec(llRowNo).tPhf.sGross, 2, smSave(11, llRowNo)
        'Get Net
        '1-17-02 combine tax with net
        'gPDNToStr tgPhfRec(llRowNo).tPhf.sNet, 2, smSave(12, llRowNo)
        llTaxes = tgPhfRec(llRowNo).tPhf.lTax1 + tgPhfRec(llRowNo).tPhf.lTax2
        gPDNToStr tgPhfRec(llRowNo).tPhf.sNet, 2, slStr
        If ((Trim$(slStr) = "-0.00") And (Trim$(smSave(11, llRowNo)) = "0.00")) Or ((Trim$(smSave(11, llRowNo)) = "-0.00") And (Trim$(slStr) = "0.00")) Then
            slStr = "0.00"
            smSave(11, llRowNo) = "0.00"
        End If
        smSave(12, llRowNo) = gAddStr(slStr, gLongToStrDec(llTaxes, 2))

        'Acquisition Cost
        smSave(14, llRowNo) = gLongToStrDec(tgPhfRec(llRowNo).tPhf.lAcquisitionCost, 2)

        'Participant
        imSave(3, llRowNo) = -1
        If tgPhfRec(llRowNo).tPhf.iMnfGroup > 0 Then
            mSSPartPop llRowNo, 0, 0, 0, 1, smSave(10, llRowNo)
            For ilLoop = 0 To UBound(tmSSPart) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmSSPart(ilLoop).iMnfGroup = tgPhfRec(llRowNo).tPhf.iMnfGroup Then
                    imSave(3, llRowNo) = ilLoop + 1
                    Exit For
                End If
            Next ilLoop
        End If
    Next llRowNo
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim slNameCode As String  'Name and code
    Dim slName As String
    Dim slStr As String
    Dim ilLoop As Integer

    mAdvtPop
    For ilLoop = 0 To UBound(tmAdvertiser) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
        slStr = Trim$(tmAdvertiser(ilLoop).sKey)   'Traffic!lbcAdvertiser.List(ilLoop)
        If (InStr(slStr, "/Direct") > 0) Or (InStr(slStr, "/Non-Payee") > 0) Then
            lbcAdvtAgyCode.AddItem slStr
        End If
    Next ilLoop
    mAgencyPop
    For ilLoop = 0 To UBound(tgAgency) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        slStr = Trim$(tgAgency(ilLoop).sKey)   'Traffic!lbcAgency.List(ilLoop)
        lbcAdvtAgyCode.AddItem slStr
    Next ilLoop
    For ilLoop = 0 To lbcAdvtAgyCode.ListCount - 1 Step 1
        slNameCode = lbcAdvtAgyCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        cbcSelect.AddItem slName
    Next ilLoop
    cbcSelect.AddItem "[All HI]", 0
    cbcSelect.AddItem "[New HI]", 0
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
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcProduct, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName) And (edcDropDown.Text <> "[New]")) Or (edcDropDown.Text <> "[None]") Then
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mProdBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igAdvtProdCallSource = CALLSOURCEPOSTITEM
    sgAdvtProdName = lbcAdvertiser.List(lbcAdvertiser.ListIndex)
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AdvtProd.Exe " & slStr, 1)
    'SaleHist.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AdvtProd.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'SaleHist.Enabled = True
    'Invoice!edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProdBranch = True
    imUpdateAllowed = ilUpdateAllowed
    If igAdvtProdCallSource = CALLDONE Then  'Done
        igAdvtProdCallSource = CALLNONE
'        gSetMenuState True
        lbcProduct.Clear
        sgProdCodeTag = ""
        mProdPop imAdfCode
        If imTerminate Then
            mProdBranch = False
            Exit Function
        End If
        gFindMatch sgAdvtProdName, 1, lbcProduct
        If gLastFound(lbcProduct) > 0 Then
            imChgMode = True
            lbcProduct.ListIndex = gLastFound(lbcProduct)
            edcDropDown.Text = lbcProduct.List(lbcProduct.ListIndex)
            imChgMode = False
            mProdBranch = False
        Else
            imChgMode = True
            lbcProduct.ListIndex = -1
            edcDropDown.Text = sgAdvtProdName
            imChgMode = False
            edcDropDown.SetFocus
            sgAdvtProdName = ""
            Exit Function
        End If
        sgAdvtProdName = ""
    End If
    If igAdvtProdCallSource = CALLCANCELLED Then  'Cancelled
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtProdCallSource = CALLTERMINATED Then
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
Private Sub mProdPop(ilAdfCode As Integer)
'
'   mProdPop
'   Where:
'       imAdfCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If ilAdfCode <= 0 Then
        lbcProduct.Clear
        'lbcProdCode.Clear
        'lbcProdCode.Tag = ""
        ReDim tgProdCode(0 To 0) As SORTCODE
        sgProdCodeTag = ""
        lbcProduct.AddItem "[None]", 0  'Force as first item on list
        lbcProduct.AddItem "[New]", 0  'Force as first item on list
        Exit Sub
    End If
    ilIndex = lbcProduct.ListIndex
    If ilIndex > 1 Then
        slName = lbcProduct.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(SaleHist, ilAdfCode, lbcProduct, lbcProdCode)
    ilRet = gPopAdvtProdBox(SaleHist, ilAdfCode, lbcProduct, tgProdCode(), sgProdCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gPopAdvtProdBox)", SaleHist
        On Error GoTo 0
        lbcProduct.AddItem "[None]", 0  'Force as first item on list
        lbcProduct.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcProduct
            If gLastFound(lbcProduct) > 1 Then
                lbcProduct.ListIndex = gLastFound(lbcProduct)
            Else
                lbcProduct.ListIndex = -1
            End If
        Else
            lbcProduct.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRvfPhfRec                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRvfPhfRec(ilAdfCode As Integer, ilAgfCode As Integer, slCntrNo As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlCharTypeBuff                                                                        *
'******************************************************************************************

'
'   iRet = mReadRvfPhfRec (ilAdfCode, ilAgfCode)
'   Where:
'       ilAdfCode(I)- Adf Code
'       ilAgfCode(I)- Agf Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llUpper As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim hlFile As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim ilAddRec As Integer
    Dim llDate As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE   'Type field record

    If rbcType(0).Value Then
        hlFile = hmPhf
        ilSPass = 1
        ilEPass = 1
    Else
        hlFile = hmRvf
        ilSPass = 1
        ilEPass = 1
        If (ilAdfCode = 0) And (ilAgfCode = 0) Then
            'ilEPass = 2
        End If
    End If
    llNoRec = btrRecords(hlFile) / 2
    If llNoRec = 0 Then
        ReDim tgPhfRec(0 To 1) As PHFREC
    Else
        ReDim tgPhfRec(0 To llNoRec) As PHFREC
    End If
    tgPhfRec(1).iStatus = -1
    tgPhfRec(1).lRecPos = 0
    ReDim tgPhfDel(0 To 0) As PHFREC
    tgPhfDel(0).iStatus = -1
    tgPhfDel(0).lRecPos = 0

    imPhfChg = False

    llUpper = 1 'UBound(tgPhfRec)
    For ilPass = ilSPass To ilEPass Step 1
        btrExtClear hlFile   'Clear any previous extend operation
        ilExtLen = Len(tgPhfRec(1).tPhf)  'Extract operation record size
        If (ilAdfCode = 0) And (ilAgfCode > 0) Then
            tmRvfPhfSrchKey0.iAgfCode = ilAgfCode
            ilRet = btrGetEqual(hlFile, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
                ilRet = BTRV_ERR_END_OF_FILE
            Else
                If ilRet <> BTRV_ERR_NONE Then
                    ReDim tgPhfRec(0 To 1) As PHFREC
                    mReadRvfPhfRec = False
                    Exit Function
                End If
            End If
        ElseIf (ilAdfCode > 0) And (ilAgfCode = 0) Then
            tmRvfPhfSrchKey1.iAdfCode = ilAdfCode
            ilRet = btrGetEqual(hlFile, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
                ilRet = BTRV_ERR_END_OF_FILE
            Else
                If ilRet <> BTRV_ERR_NONE Then
                    ReDim tgPhfRec(0 To 1) As PHFREC
                    mReadRvfPhfRec = False
                    Exit Function
                End If
            End If
        ElseIf (ilAdfCode = 0) And (ilAgfCode = 0) And slCntrNo <> "" Then
            tmRvfPhfSrchKey4.lCntrNo = Val(slCntrNo)
            tmRvfPhfSrchKey4.iTranDate(0) = 0
            tmRvfPhfSrchKey4.iTranDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hlFile, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
                ilRet = BTRV_ERR_END_OF_FILE
            Else
                If ilRet <> BTRV_ERR_NONE Then
                    ReDim tgPhfRec(0 To 1) As PHFREC
                    mReadRvfPhfRec = False
                    Exit Function
                End If
            End If
        Else
            ilRet = btrGetFirst(hlFile, tmRvfPhf, imRvfPhfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)
            If (ilAdfCode = 0) And (ilAgfCode > 0) Then
                ilOffSet = gFieldOffset("Phf", "PhfAgfCode")
                tlIntTypeBuff.iType = ilAgfCode
                ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                On Error GoTo mReadRvfPhfRecErr
                gBtrvErrorMsg ilRet, "mReadRvfPhfRec (btrExtAddLogicConst):" & "Phf.Btr", SaleHist
                On Error GoTo 0
            ElseIf (ilAdfCode > 0) And (ilAgfCode = 0) Then
                ilOffSet = gFieldOffset("Phf", "PhfAdfCode")
                tlIntTypeBuff.iType = ilAdfCode
                ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                On Error GoTo mReadRvfPhfRecErr
                gBtrvErrorMsg ilRet, "mReadRvfPhfRec (btrExtAddLogicConst):" & "Phf.Btr", SaleHist
                On Error GoTo 0
            ElseIf (ilAdfCode = 0) And (ilAgfCode = 0) And (slCntrNo <> "") Then
                ilOffSet = gFieldOffset("Phf", "PhfCntrNo")
                tlLongTypeBuff.lCode = Val(slCntrNo)
                ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)
                On Error GoTo mReadRvfPhfRecErr
                gBtrvErrorMsg ilRet, "mReadRvfPhfRec (btrExtAddLogicConst):" & "Phf.Btr", SaleHist
                On Error GoTo 0
            End If
            ilRet = btrExtAddField(hlFile, 0, ilExtLen) 'Extract the whole record
            On Error GoTo mReadRvfPhfRecErr
            gBtrvErrorMsg ilRet, "mReadRvfPhfRec (btrExtAddField):" & "Phf.Btr", SaleHist
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
            ilRet = btrExtGetNext(hlFile, tgPhfRec(llUpper).tPhf, ilExtLen, tgPhfRec(llUpper).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mReadRvfPhfRecErr
                gBtrvErrorMsg ilRet, "mReadRvfPhfRec (btrExtGetNextExt):" & "Phf.Btr", SaleHist
                On Error GoTo 0
                'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
                ilExtLen = Len(tgPhfRec(1).tPhf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFile, tgPhfRec(llUpper).tPhf, ilExtLen, tgPhfRec(llUpper).lRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
'                    slStr = ""
'                    ilRecOk = True
'                    If tgPhfRec(llUpper).tPhf.iAdfCode <> tmAdf.iCode Then
'                        tmAdfSrchKey.iCode = tgPhfRec(llUpper).tPhf.iAdfCode 'ilCode
'                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Else
'                        ilRet = BTRV_ERR_NONE
'                    End If
'                    If ilRet = BTRV_ERR_NONE Then
'                        If tgPhfRec(llUpper).tPhf.lPrfCode > 0 Then
'                            If tgPhfRec(llUpper).tPhf.lPrfCode <> tmPrf.lCode Then
'                                tmPrfSrchKey0.lCode = tgPhfRec(llUpper).tPhf.lPrfCode 'ilCode
'                                ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                            Else
'                                ilRet = BTRV_ERR_NONE
'                            End If
'                            If ilRet <> BTRV_ERR_NONE Then
'                                ilRecOk = False
'                            End If
'                        Else
'                            tmPrf.sName = ""
'                        End If
'                    Else
'                        ilRecOk = False
'                    End If
'                    If ilRecOk Then
'                        slCntrNo = Trim$(Str$(tgPhfRec(llUpper).tPhf.lCntrNo))
'                        Do While Len(slCntrNo) < 10
'                            slCntrNo = "0" & slCntrNo
'                        Loop
'                        gUnpackDateForSort tgPhfRec(llUpper).tPhf.iTranDate(0), tgPhfRec(llUpper).tPhf.iTranDate(1), slDate
'                        tgPhfRec(llUpper).sKey = tmAdf.sName & tmPrf.sName & slCntrNo & slDate
                        'Bypass PO since Advertiser, vehicles are not set.  The Client can use Transfer to move PO's
                        ilAddRec = True
                        If (Trim$(tgPhfRec(llUpper).tPhf.sTranType) = "PO") And (rbcType(0).Value) Then
                            ilAddRec = False
                        End If
                        If (ilAdfCode = 0) And (ilAgfCode = 0) And (slCntrNo = "") Then
                            If rbcType(0).Value Then
                                'Only Invoice (HI)
                                If (Trim$(tgPhfRec(llUpper).tPhf.sTranType) <> "HI") Then
                                    ilAddRec = False
                                End If
                            Else
                                'No exceptions
                            End If
                        End If
                        '1/6/09: Only check dates if [All HI] selected
                        If (ilAddRec) And (rbcType(0).Value) And (imSelectedIndex = 1) Then
                            gUnpackDateLong tgPhfRec(llUpper).tPhf.iTranDate(0), tgPhfRec(llUpper).tPhf.iTranDate(1), llDate
                            If (llDate < lmSalesHistStartDate) Or (llDate > lmSalesHistEndDate) Then
                                ilAddRec = False
                            End If
                        End If
                        If ilAddRec Then
                            '3/2/05
                            'If we want to include Merch and Promotion, then smSave(1,--) must allow for M and P
                            'Also, pbcCT and mMoveCtrlToRec needs to be changed
                            If (Trim$(tgPhfRec(llUpper).tPhf.sCashTrade) = "C") Or (Trim$(tgPhfRec(llUpper).tPhf.sCashTrade) = "T") Then
                                tgPhfRec(llUpper).iStatus = 1
        '                        tgPhfRec(llUpper).sProduct = tmPrf.sName
                                llUpper = llUpper + 1
                                If llUpper >= UBound(tgPhfRec) Then
                                    ReDim Preserve tgPhfRec(0 To llUpper + 500) As PHFREC
                                End If
                                tgPhfRec(llUpper).iStatus = -1
                                tgPhfRec(llUpper).lRecPos = 0
                            End If
                        End If
'                    End If
                    ilRet = btrExtGetNext(hlFile, tgPhfRec(llUpper).tPhf, ilExtLen, tgPhfRec(llUpper).lRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlFile, tgPhfRec(llUpper).tPhf, ilExtLen, tgPhfRec(llUpper).lRecPos)
                    Loop
                Loop
            End If
        End If
    Next ilPass
    ReDim Preserve tgPhfRec(0 To llUpper) As PHFREC
'    If llUpper > 1 Then
'        ArraySortTyp fnAV(tgPhfRec(), 1), UBound(tgPhfRec) - 1, 0, LenB(tgPhfRec(1)), 0, LenB(tgPhfRec(1).sKey), 0
'    End If
    mBuildKey
    'mInitBudgetCtrls
    mReadRvfPhfRec = True
    Exit Function
mReadRvfPhfRecErr:
    On Error GoTo 0
    mReadRvfPhfRec = False
    Exit Function
End Function
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
    Dim llRowNo As Long
    Dim slMsg As String
    Dim llPhf As Long
    Dim hlFile As Integer
    Dim tlPhf As RVF
    Dim tlPhf1 As MOVEREC
    Dim tlPhf2 As MOVEREC

    imPriceChgd = False
    ReDim tmPhfSplit(0 To 0) As PHFREC
    imSSPartSplit = False

    If rbcType(0).Value Then
        hlFile = hmPhf
    Else
        hlFile = hmRvf
    End If
    mSetShow imBoxNo
    For llRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(llRowNo) = NO Then
            mSaveRec = False
            lmRowNo = llRowNo
            Exit Function
        End If
    Next llRowNo
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    If rbcType(1).Value Then
        'Remove history of matching invoice numbers
        'For llRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        For llRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
            mTestInvNo Trim$(smSave(5, llRowNo)), 1, llRowNo
        Next llRowNo
    End If
    ilRet = btrBeginTrans(hlFile, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
        Exit Function
    End If
    For llPhf = LBONE To UBound(tgPhfRec) - 1 Step 1
        Do  'Loop until record updated or added
            If (tgPhfRec(llPhf).iStatus = 0) Then  'New selected
                'User
                gPackDate smNowDate, tgPhfRec(llPhf).tPhf.iDateEntrd(0), tgPhfRec(llPhf).tPhf.iDateEntrd(1)
                tgPhfRec(llPhf).tPhf.lCode = 0
                tgPhfRec(llPhf).tPhf.iUrfCode = tgUrf(0).iCode
                tgPhfRec(llPhf).tPhf.lCefCode = 0
                tgPhfRec(llPhf).tPhf.lSbfCode = 0
                tgPhfRec(llPhf).tPhf.lPcfCode = 0
                tgPhfRec(llPhf).tPhf.sInCollect = "N"
                tgPhfRec(llPhf).tPhf.iRemoteID = 0
                'tgPhfRec(llPhf).tPhf.lAcquisitionCost = 0
                'All new transaction saved to History entered as Air Time (Revenue)
                'All new transactions saved to Receivables entered as Installment (Billed)
                If rbcType(0).Value Then
                    '6/28/12: Test if using installment
                    'tgPhfRec(llPhf).tPhf.sType = "A"
                    If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT) Then
                        tgPhfRec(llPhf).tPhf.sType = ""
                    Else
                        tgPhfRec(llPhf).tPhf.sType = "A"
                    End If
                Else
                    If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT) Then
                        tgPhfRec(llPhf).tPhf.sType = ""
                    Else
                        tgPhfRec(llPhf).tPhf.sType = "I"
                    End If
                End If
                '1/17/09:  Added buyer
                'tgPhfRec(llPhf).tPhf.sUnused = ""
                'ilRet = btrInsert(hlFile, tgPhfRec(llPhf).tPhf, imRvfPhfRecLen, INDEXKEY0)
                ilRet = mAdjSSPart(tgPhfRec(llPhf).tPhf)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hlFile)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
                    Exit Function
                End If
                slMsg = "mSaveRec (btrInsert: Sales History)"
                'ilRet = btrGetPosition(hlFile, tgPhfRec(llPhf).lRecPos)
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hlFile)
                '    Screen.MousePointer = vbDefault
                '    ilRet = MsgBox("Update Not Completed, Try Later", vbOkOnly + vbExclamation, "Backlog")
                '    Exit Function
                'End If
                tgPhfRec(llPhf).iStatus = 1
            ElseIf (tgPhfRec(llPhf).iStatus = 1) Then  'Old record-Update
                'slMsg = "mSaveRec (btrGetDirect: Sales History)"
                'ilRet = btrGetDirect(hlFile, tlPhf, imRvfPhfRecLen, tgPhfRec(llPhf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                slMsg = "mSaveRec (btrGetEqual: Sales History)"
                tmRvfSrchKey2.lCode = tgPhfRec(llPhf).tPhf.lCode
                ilRet = btrGetEqual(hlFile, tlPhf, imRvfPhfRecLen, tmRvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hlFile)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
                    Exit Function
                End If
                tgPhfRec(llPhf).tPhf.sType = tlPhf.sType
                LSet tlPhf1 = tlPhf
                LSet tlPhf2 = tgPhfRec(llPhf).tPhf
                If StrComp(tlPhf1.sChar, tlPhf2.sChar, 0) <> 0 Then
                    'tmRec = tlPhf
                    'ilRet = gGetByKeyForUpdate("RVF", hlFile, tmRec)
                    'tlPhf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hlFile)
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("Update Not Completed, Try Later", vbOkOnly + vbExclamation, "Backlog")
                    '    Exit Function
                    'End If
                    tgPhfRec(llPhf).tPhf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hlFile, tgPhfRec(llPhf).tPhf, imRvfPhfRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Sales History)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hlFile)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
            Exit Function
        End If
    Next llPhf
    For llPhf = LBound(tgPhfDel) To UBound(tgPhfDel) - 1 Step 1
        If tgPhfDel(llPhf).iStatus = 1 Then
            Do
                'slMsg = "mSaveRec (btrGetDirect: Sales History)"
                'ilRet = btrGetDirect(hlFile, tlPhf, imRvfPhfRecLen, tgPhfDel(llPhf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                slMsg = "mSaveRec (btrGetEqual: Sales History)"
                tmRvfSrchKey2.lCode = tgPhfDel(llPhf).tPhf.lCode
                ilRet = btrGetEqual(hlFile, tlPhf, imRvfPhfRecLen, tmRvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hlFile)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
                    Exit Function
                End If
                'tmRec = tlPhf
                'ilRet = gGetByKeyForUpdate("RVF", hlFile, tmRec)
                'tlPhf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                 '   ilRet = btrAbortTrans(hlFile)
                '    Screen.MousePointer = vbDefault
                '    ilRet = MsgBox("Update Not Completed, Try Later", vbOkOnly + vbExclamation, "Backlog")
                '    Exit Function
                'End If
                ilRet = btrDelete(hlFile)
                slMsg = "mSaveRec (btrDelete: Sales History)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hlFile)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Backlog")
                Exit Function
            End If
        End If
    Next llPhf
    ilRet = btrEndTrans(hlFile)
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
    Dim llLoop As Long
    Dim ilNew As Integer
    If imPhfChg And (UBound(tgPhfRec) > LBONE) Or (UBound(tgPhfDel) > LBound(tgPhfDel)) Then
        If ilAsk Then
            ilNew = True
            For llLoop = LBONE To UBound(tgPhfRec) - 1 Step 1
                If tgPhfRec(llLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next llLoop
            For llLoop = LBound(tgPhfDel) To UBound(tgPhfDel) - 1 Step 1
                If tgPhfDel(llLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next llLoop
            If Not ilNew Then
                slMess = "Save Changes"
            Else
                slMess = "Add Changes"
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
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/12/93       By:D. LeVine      *
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
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
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
    If (imBypassSetting) Or (Not imUpdateAllowed) Or (Not igPasswordOk) Then
        Exit Sub
    End If
    ilAltered = imPhfChg
    If (Not ilAltered) And (UBound(tgPhfDel) > LBound(tgPhfDel)) Then
        ilAltered = True
    End If
    If ilAltered Then
        pbcHist.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        cbcSelect.Enabled = False
        edcCntrNo.Enabled = False
        rbcType(0).Enabled = False
        rbcType(1).Enabled = False
    Else
        If (imSelectedIndex < 0) And (Trim(edcCntrNo.Text) = "") Then
            pbcHist.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cbcSelect.Enabled = True
            edcCntrNo.Enabled = True
        Else
            pbcHist.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            cbcSelect.Enabled = True
            edcCntrNo.Enabled = True
        End If
        rbcType(0).Enabled = True
        rbcType(1).Enabled = True
    End If
    'Update button set if all mandatory fields have data and any field altered
    'If (mTestFields() = YES) And (ilAltered) And (UBound(tgPhfRec) > 1) Then
    If (mTestFields() = YES) And (ilAltered) And (UBound(tgPhfRec) > 1) Or (ilAltered And (UBound(tgPhfDel) > LBound(tgPhfDel))) Then
        cmcSave.Enabled = True
        cmcYear.Enabled = False
    Else
        cmcSave.Enabled = False
        If (rbcType(0).Value) And (imSelectedIndex = 1) Then
            cmcYear.Enabled = True
        Else
            cmcYear.Enabled = False
        End If
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (lmRowNo < vbcHist.Value) Or (lmRowNo >= vbcHist.Value + vbcHist.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo
        Case CTINDEX
            pbcCT.Visible = True
            pbcCT.SetFocus
        Case ADVTINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AGYINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SPERSONINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case INVNOINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case CNTRNOINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case BVEHINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AVEHINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PKLNINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case TRANDATEINDEX
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TRANTYPEINDEX
            pbcTT.Visible = True
            pbcTT.SetFocus
        Case GROSSINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case NETINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case ACQCOSTINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SSPARTINDEX  'Sales Source/Participant
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetMinMax                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set scroll bar min/max         *
'*                                                     *
'*******************************************************
Private Sub mSetMinMax()
    imSettingValue = True
    vbcHist.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcHist.LargeChange + 1 Then ' + 1 Then
        vbcHist.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcHist.Max = UBound(smShow, 2) - vbcHist.LargeChange
    End If
    imSettingValue = True
    If vbcHist.Value = vbcHist.Min Then
        vbcHist_Change
    Else
        vbcHist.Value = vbcHist.Min
    End If
    imSettingValue = False
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
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slAgyRate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilSbfTrfCode As Integer
    Dim ilTrfAgyAdvt As Integer
    Dim ilAdfCode As Integer
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim llTax1Rate As Long
    Dim llTax2Rate As Long
    Dim slTax1Rate As String
    Dim slTax2Rate As String
    Dim slTaxRate As String
    Dim slGrossNet As String

    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (lmRowNo < vbcHist.Value) Or (lmRowNo >= vbcHist.Value + vbcHist.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo
        Case CTINDEX
            pbcCT.Visible = False
            If imSave(1, lmRowNo) = 0 Then
                slStr = "C"
            ElseIf imSave(1, lmRowNo) = 1 Then
                slStr = "T"
            Else
                slStr = ""
            End If
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
        Case ADVTINDEX
            lbcAdvertiser.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(2, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(2, lmRowNo) = slStr
                If StrComp(slStr, "[None]", vbTextCompare) = 0 Then
                    smSave(3, lmRowNo) = ""
                End If
            End If
        Case PRODINDEX
            lbcProduct.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(3, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(3, lmRowNo) = slStr
            End If
        Case AGYINDEX
            lbcAgency.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(1, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(1, lmRowNo) = slStr
            End If
        Case SPERSONINDEX
            lbcSPerson.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(4, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(4, lmRowNo) = slStr
            End If
        Case INVNOINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(5, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(5, lmRowNo) = slStr
                mTestInvNo slStr, 0, lmRowNo
            End If
        Case CNTRNOINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(6, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(6, lmRowNo) = slStr
            End If
        Case BVEHINDEX
            lbcBVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(7, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(7, lmRowNo) = slStr
                mCreateVef slStr
            End If
        Case AVEHINDEX
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(8, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(8, lmRowNo) = slStr
            End If
        Case PKLNINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(9, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(9, lmRowNo) = slStr
            End If
        Case TRANDATEINDEX
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(10, lmRowNo)) <> slStr Then
                    imPhfChg = True
                    smSave(10, lmRowNo) = slStr
                End If
            End If
        Case NTRTYPEINDEX  'Vehicle
            lbcNTRType.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcNTRType.ListIndex <> imSave(2, lmRowNo) Then
                imPhfChg = True
            End If
            If lbcNTRType.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcNTRType.List(lbcNTRType.ListIndex)
            End If
            imSave(2, lmRowNo) = lbcNTRType.ListIndex
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
        Case NTRTAXINDEX
            lbcTax.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTax.ListIndex < 0 Then
                slStr = ""
            ElseIf lbcTax.ListIndex = 0 Then
                slStr = "N"
            Else
                slStr = "Y"
            End If
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If imSave(4, lmRowNo) <> lbcTax.ListIndex Then
                imSave(4, lmRowNo) = lbcTax.ListIndex
                imPhfChg = True
            End If
        Case TRANTYPEINDEX
            'pbcTT.Visible = False
            'If rbcType(0).Value Then
            '    If Trim$(smSave(13, lmRowNo)) = "HI" Then
            '        slStr = "H"
            '    Else
            '        slStr = ""
            '    End If
            'Else
            '    If Trim$(smSave(13, lmRowNo)) = "PI" Then
            '        slStr = "PI"
            '    ElseIf Trim$(smSave(13, lmRowNo)) = "IN" Then
            '        slStr = "IN"
            '    Else
            '        slStr = ""
            '    End If
            'End If
            'gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            'smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            lbcTranType.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (lbcTranType.ListIndex < 0) Then
                slStr = ""
            Else
                slStr = Left$(lbcTranType.List(lbcTranType.ListIndex), 2)
            End If
            If smSave(13, lmRowNo) <> slStr Then
                imPhfChg = True
            End If
            smSave(13, lmRowNo) = slStr
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
        Case GROSSINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If Trim$(smSave(11, lmRowNo)) <> slStr Then
                If (Trim$(smSave(13, lmRowNo)) = "IN") Or (Trim$(smSave(13, lmRowNo)) = "HI") Then
                    If InStr(1, slStr, "-", vbTextCompare) > 0 Then
                        inGenMsgRet = True
                        sgGenMsg = "Invoices should be Positive" & ": Make Positive, Leave Negative or Cancel"
                        sgCMCTitle(0) = "Make +"    '"Make Positive"
                        sgCMCTitle(1) = "Leave -"   '"Leave Negative"
                        sgCMCTitle(2) = "Cancel"
                        sgCMCTitle(3) = ""
                        igDefCMC = 0
                        igEditBox = 0
                        GenMsg.Show vbModal
                        DoEvents
                        inGenMsgRet = False
                        If igAnsCMC = 0 Then
                            slStr = Abs(slStr)
                        ElseIf igAnsCMC = 2 Then
                            Exit Sub
                        End If
                    End If
                End If
                imPriceChgd = True
                imPhfChg = True
                smSave(11, lmRowNo) = slStr
            End If
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            'Required if Sign question asked
            pbcHist_Paint
        Case NETINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If Trim$(smSave(12, lmRowNo)) <> slStr Then
                If Trim$(smSave(13, lmRowNo)) = "PI" Then
                    If InStr(1, slStr, "-", vbTextCompare) <= 0 Then
                        sgGenMsg = "Payments should be Negative" & ": Make Negative, Leave Positive or Cancel"
                        sgCMCTitle(0) = "Make -"    '"Make Negative"
                        sgCMCTitle(1) = "Leave +"   '"Leave Positive"
                        sgCMCTitle(2) = "Cancel"
                        sgCMCTitle(3) = ""
                        igDefCMC = 0
                        igEditBox = 0
                        inGenMsgRet = True
                        GenMsg.Show vbModal
                        DoEvents
                        inGenMsgRet = False
                        If igAnsCMC = 0 Then
                            slStr = "-" & slStr
                        ElseIf igAnsCMC = 2 Then
                            Exit Sub
                        End If
                    End If
                End If
                If (Trim$(smSave(13, lmRowNo)) = "IN") Or (Trim$(smSave(13, lmRowNo)) = "HI") Then
                    If InStr(1, slStr, "-", vbTextCompare) > 0 Then
                        inGenMsgRet = True
                        sgGenMsg = "Invoices should be Positive" & ": Make Positive, Leave Negative or Cancel"
                        sgCMCTitle(0) = "Make +"    '"Make Positive"
                        sgCMCTitle(1) = "Leave -"   '"Leave Negative"
                        sgCMCTitle(2) = "Cancel"
                        sgCMCTitle(3) = ""
                        igDefCMC = 0
                        igEditBox = 0
                        GenMsg.Show vbModal
                        DoEvents
                        inGenMsgRet = False
                        If igAnsCMC = 0 Then
                            slStr = Abs(slStr)
                        ElseIf igAnsCMC = 2 Then
                            Exit Sub
                        End If
                    End If
                End If
                gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
                imPriceChgd = True
                imPhfChg = True
                smSave(12, lmRowNo) = slStr
                If (Trim$(smSave(13, lmRowNo)) <> "PI") And (Trim$(smSave(13, lmRowNo)) <> "PO") And (Left$(smSave(13, lmRowNo), 1) <> "W") Then
                    'If (Trim$(smSave(12, lmRowNo)) <> "") And (Trim$(smSave(11, lmRowNo)) = "") Then
                    If (Trim$(smSave(12, lmRowNo)) <> "") Then
                        slAgyRate = "1.00"
                        mGetAgy smSave(1, lmRowNo)
                        If imSave(1, lmRowNo) = 1 Then  'Trade
                            If imAgfCode <= 0 Then
                                slStr = smSave(12, lmRowNo)
                            Else
                                slAgyRate = gDivStr(gSubStr("100.00", gIntToStrDec(tmAgf.iComm, 2)), "100.00")
                                slStr = gDivStr(smSave(12, lmRowNo), slAgyRate) '".85")
                            End If
                        Else
                            ilAdfCode = 0
                            If imAgfCode <= 0 Then
                                gFindMatch Trim$(smSave(2, lmRowNo)), 2, lbcAdvertiser
                                If gLastFound(lbcAdvertiser) > 1 Then
                                    slNameCode = tmAdvertiser(gLastFound(lbcAdvertiser) - 2).sKey    'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvertiser) - 1)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    ilAdfCode = Val(slCode)
                                End If
                            End If

                            ilTrfAgyAdvt = gGetTrfIndexForAgyAdvt(ilAdfCode, imAgfCode)
                            If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) And (imSave(2, lmRowNo) <= 0) And (ilTrfAgyAdvt <> -1) Then
                                'Air time: Backcompute gross with net including tax
                                ilVefCode = 0
                                gFindMatch Trim$(smSave(8, lmRowNo)), 0, lbcVehicle
                                If gLastFound(lbcVehicle) >= 0 Then
                                    slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    ilVefCode = Val(slCode)
                                End If
                                gGetAirTimeTaxValues ilTrfAgyAdvt, ilVefCode, llTax1Rate, llTax2Rate, slGrossNet
                                If imAgfCode <= 0 Then
                                    slAgyRate = "1.0"
                                Else
                                    slAgyRate = gDivStr(gSubStr("100.00", gIntToStrDec(tmAgf.iComm, 2)), "100.00")
                                End If
                                slTax1Rate = gLongToStrDec(llTax1Rate, 4)
                                slTax2Rate = gLongToStrDec(llTax2Rate, 4)
                                slTaxRate = gDivStr(gAddStr(slTax1Rate, slTax2Rate), "100.00")
                                slStr = gDivStr(smSave(12, lmRowNo), gAddStr(slAgyRate, slTaxRate)) '".85")
                            ElseIf ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) And (imSave(2, lmRowNo) > 0) Then
                                'NTR: Backcompute gross with net including tax
                                '12/24/06:  Need to add tax selective
                                ilSbfTrfCode = 0
                                If lmSave(1, lmRowNo) > 0 Then
                                    tmSbfSrchKey1.lCode = lmSave(1, lmRowNo)
                                    ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) Then
                                        ilSbfTrfCode = tmSbf.iTrfCode
                                    End If
                                Else
                                    If (imSave(4, lmRowNo) > 0) Then
                                        ilSbfTrfCode = lbcTax.ItemData(imSave(4, lmRowNo))
                                    End If
                                End If
                                gGetNTRTaxRates ilSbfTrfCode, llTax1Rate, llTax2Rate, slGrossNet

                                If imAgfCode <= 0 Then
                                    slAgyRate = "1.0"
                                Else
                                    slAgyRate = gDivStr(gSubStr("100.00", gIntToStrDec(tmAgf.iComm, 2)), "100.00")
                                End If
                                slTax1Rate = gLongToStrDec(llTax1Rate, 4)
                                slTax2Rate = gLongToStrDec(llTax2Rate, 4)
                                slTaxRate = gDivStr(gAddStr(slTax1Rate, slTax2Rate), "100.00")
                                slStr = gDivStr(smSave(12, lmRowNo), gAddStr(slAgyRate, slTaxRate)) '".85")
                            Else
                                If imAgfCode <= 0 Then
                                    slStr = smSave(12, lmRowNo)
                                Else
                                    slAgyRate = gDivStr(gSubStr("100.00", gIntToStrDec(tmAgf.iComm, 2)), "100.00")
                                    slStr = gDivStr(smSave(12, lmRowNo), slAgyRate) '".85")
                                End If
                            End If
                        End If
                        slStr = gRoundStr(slStr, ".01", 2)
                        If Trim$(smSave(11, lmRowNo)) = "" Then
                            gSetShow pbcHist, slStr, tmCtrls(GROSSINDEX)
                            smShow(GROSSINDEX, lmRowNo) = tmCtrls(GROSSINDEX).sShow
                            smSave(11, lmRowNo) = slStr
                        End If
                        'Tax = Net with Tax - Net w/o Tax
                        'Net w/o Tax = Gross *(1 - AgyComm) = Gross * slAgyRate
                        If imSave(1, lmRowNo) = 1 Then  'Trade
                            smSave(15, lmRowNo) = "0.00"
                        Else
                            smSave(15, lmRowNo) = gSubStr(smSave(12, lmRowNo), gMulStr(slStr, slAgyRate))
                        End If
                    Else
                        smSave(15, lmRowNo) = "0.00"
                    End If
                Else
                    slStr = ""
                    gSetShow pbcHist, slStr, tmCtrls(GROSSINDEX)
                    smShow(GROSSINDEX, lmRowNo) = tmCtrls(GROSSINDEX).sShow
                    smSave(11, lmRowNo) = slStr
                End If
                'Required if Sign question asked
                pbcHist_Paint
            End If
        Case ACQCOSTINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(14, lmRowNo)) <> slStr Then
                imPhfChg = True
                smSave(14, lmRowNo) = slStr
            End If
        Case SSPARTINDEX  'Sales Source/Participant
            lbcSSPart.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If imSave(3, lmRowNo) <> lbcSSPart.ListIndex Then
                imPhfChg = True
            End If
            If lbcSSPart.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcSSPart.List(lbcSSPart.ListIndex)
            End If
            imSave(3, lmRowNo) = lbcSSPart.ListIndex
            gSetShow pbcHist, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, lmRowNo) = tmCtrls(ilBoxNo).sShow
    End Select
    mSetCommands
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
'       ilInfo(I)- True= Info box; False=Adjustment
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
    If ((ilRet = 0) And (Not imDoubleClickName) And (edcDropDown.Text <> "[New]")) Then
        mSPersonBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(SALESPEOPLELIST)) Then
    '    imDoubleClickName = False
    '    mSPersonBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igSlfCallSource = CALLSOURCEPOSTITEM
    If edcDropDown.Text = "[New]" Then
        sgSlfName = ""
    Else
        sgSlfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSlfCallSource)) & "\" & sgSlfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "SPerson.Exe " & slStr, 1)
    'SaleHist.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    SPerson.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgSlfName)
    igSlfCallSource = Val(sgSlfName)
    ilParse = gParseItem(slStr, 2, "\", sgSlfName)
    'SaleHist.Enabled = True
    'Invoice!edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSPersonBranch = True
    imUpdateAllowed = ilUpdateAllowed
    ilCallReturn = igSlfCallSource
    slName = sgSlfName
    'igVsfCallSource = CALLNONE
    igSlfCallSource = CALLNONE
    'sgVsfName = ""
    sgSlfName = ""
    mSPersonBranch = True
    If ilCallReturn = CALLDONE Then  'Done
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
            'mInfoSetChg SALESPINDEX
        Else
            imChgMode = True
            lbcSPerson.ListIndex = 0
            edcDropDown.Text = lbcSPerson.List(0)
            imChgMode = False
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
    ilIndex = lbcSPerson.ListIndex
    If ilIndex > 1 Then
        slName = lbcSPerson.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSPersonComboBox(SaleHist, cbcSelect, Traffic!lbcSalesperson, Traffic!cbcSelectCombo, igSlfFirstNameFirst)
    ilType = 0  'All
    'ilRet = gPopSalespersonBox(SaleHist, ilType, False, True, lbcSPerson, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(SaleHist, ilType, False, True, lbcSPerson, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", SaleHist
        On Error GoTo 0
        lbcSPerson.AddItem "[None]", 0  'Force as first item on list
        lbcSPerson.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcSPerson
            If gLastFound(lbcSPerson) >= 0 Then
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
'*      Procedure Name:mSSPartPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales Source/Participant*
'*                                                     *
'*******************************************************
Private Sub mSSPartPop(llRowNo As Long, ilInAirVefCode As Integer, ilInSlfCode As Integer, ilInAdfCode As Integer, ilFlag As Integer, slTranDate As String)
    '  ilFlag = 0 = Add AutoSplit
    '
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim ilSS As Integer
    Dim ilP As Integer
    Dim ilMnfSSCode As Integer
    Dim ilRet As Integer
    Dim ilAirVefCode As Integer
    Dim ilSlfCode As Integer
    Dim ilAdfCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    ReDim tmSSPart(0 To 0) As SSPART

    lbcSSPart.Clear
    ilMnfSSCode = 0
    If llRowNo > 0 Then
        gFindMatch Trim$(smSave(8, llRowNo)), 0, lbcVehicle
        If gLastFound(lbcVehicle) >= 0 Then
            slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAirVefCode = Val(slCode)
        Else
            Exit Sub
        End If
        gFindMatch Trim$(smSave(4, llRowNo)), 2, lbcSPerson
        If gLastFound(lbcSPerson) > 1 Then
            slNameCode = tgSalesperson(gLastFound(lbcSPerson) - 2).sKey  'Traffic!lbcSalesperson.List(gLastFound(lbcSPerson) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSlfCode = Val(slCode)
        Else
            Exit Sub
        End If
        gFindMatch Trim$(smSave(2, llRowNo)), 2, lbcAdvertiser
        If gLastFound(lbcAdvertiser) > 1 Then
            slNameCode = tmAdvertiser(gLastFound(lbcAdvertiser) - 2).sKey    'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvertiser) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAdfCode = Val(slCode)
        Else
            Exit Sub
        End If
    Else
        ilAirVefCode = ilInAirVefCode
        ilSlfCode = ilInSlfCode
        ilAdfCode = ilInAdfCode
    End If
    tmSlfSrchKey.iCode = ilSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        tmSofSrchKey.iCode = tmSlf.iSofCode
        ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilMnfSSCode = tmSof.iMnfSSCode
        End If
    End If
    tmAdfSrchKey.iCode = ilAdfCode
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        tmAdf.sRepInvGen = "I"
    End If
    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
    '    If ilAirVefCode = tgMVef(ilVef).iCode Then
        ilVef = gBinarySearchVef(ilAirVefCode)
        If ilVef <> -1 Then
            ilRet = gObtainPIF_ForVefDate(hmPif, tgMVef(ilVef).iCode, slTranDate, tmPif())
            'For ilLoop = LBound(tgMVef(ilVef).iMnfSSCode) To UBound(tgMVef(ilVef).iMnfSSCode) - 1 Step 1
            For ilLoop = LBound(tmPif) To UBound(tmPif) - 1 Step 1
                'If (tgMVef(ilVef).iMnfSSCode(ilLoop) > 0) And ((ilMnfSSCode = 0) Or (ilMnfSSCode = tgMVef(ilVef).iMnfSSCode(ilLoop))) Then
                If ((ilMnfSSCode = 0) Or (ilMnfSSCode = tmPif(ilLoop).iMnfSSCode)) Then
                    For ilSS = LBound(tmSMnf) To UBound(tmSMnf) - 1 Step 1
                        'If tgMVef(ilVef).iMnfSSCode(ilLoop) = tmSMnf(ilSS).iCode Then
                        If tmPif(ilLoop).iMnfSSCode = tmSMnf(ilSS).iCode Then
                            If Trim$(tmSMnf(ilSS).sUnitType) = "A" Then
                                For ilP = LBound(tmHMnf) To UBound(tmHMnf) - 1 Step 1
                                    'If tgMVef(ilVef).iMnfGroup(ilLoop) = tmHMnf(ilP).iCode Then
                                    If tmPif(ilLoop).iMnfGroup = tmHMnf(ilP).iCode Then
                                        tmSSPart(UBound(tmSSPart)).sKey = Trim$(tmSMnf(ilSS).sName) & "/" & Trim$(tmHMnf(ilP).sName)
                                        tmSSPart(UBound(tmSSPart)).iMnfSSCode = tmSMnf(ilSS).iCode
                                        tmSSPart(UBound(tmSSPart)).iMnfGroup = tmHMnf(ilP).iCode
                                        tmSSPart(UBound(tmSSPart)).iVefIndex = ilVef
                                        tmSSPart(UBound(tmSSPart)).iSSPartLp = ilLoop   'Not Used
                                        'tmSSPart(UBound(tmSSPart)).iProdPct = tgMVef(ilVef).iProdPct(ilLoop)
                                        tmSSPart(UBound(tmSSPart)).iProdPct = tmPif(ilLoop).iProdPct
                                        If tmAdf.sRepInvGen = "E" Then
                                            'tmSSPart(UBound(tmSSPart)).sUpdateRVF = tgMVef(ilVef).sExtUpdateRvf(ilLoop)
                                            tmSSPart(UBound(tmSSPart)).sUpdateRVF = tmPif(ilLoop).sExtUpdateRvf
                                        Else
                                            'tmSSPart(UBound(tmSSPart)).sUpdateRVF = tgMVef(ilVef).sUpdateRVF(ilLoop)
                                            tmSSPart(UBound(tmSSPart)).sUpdateRVF = tmPif(ilLoop).sUpdateRVF
                                        End If
                                        ReDim Preserve tmSSPart(0 To UBound(tmSSPart) + 1) As SSPART
                                        Exit For
                                    End If
                                Next ilP
                            End If
                            Exit For
                        End If
                    Next ilSS
                End If
            Next ilLoop
        End If
    'Next ilVef
    If UBound(tmSSPart) - 1 > 0 Then
        ArraySortTyp fnAV(tmSSPart(), 0), UBound(tmSSPart), 0, LenB(tmSSPart(0)), 0, LenB(tmSSPart(0).sKey), 0
    End If
    For ilLoop = LBound(tmSSPart) To UBound(tmSSPart) - 1 Step 1
        lbcSSPart.AddItem Trim$(tmSSPart(ilLoop).sKey)
    Next ilLoop
    If ilFlag = 0 Then
        lbcSSPart.AddItem "[Auto Split]", 0  'Force as first item on list
    Else
        lbcSSPart.AddItem "[None]", 0  'Force as first item on list
    End If
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload SaleHist
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
    Dim llRowNo As Long
    'For llRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
    For llRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        'If smSave(1, llRowNo) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If

        If (Trim$(smSave(2, llRowNo)) = "") Or (Trim$(smSave(2, llRowNo)) = "[None]") Then
            If Trim$(smSave(13, llRowNo)) <> "PO" Then
                mTestFields = NO
                Exit Function
            End If
        End If
        If Trim$(smSave(4, llRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(13, llRowNo)) <> "PO" Then
            If Trim$(smSave(5, llRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Trim$(smSave(7, llRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Trim$(smSave(8, llRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
        End If
        If Trim$(smSave(10, llRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(11, llRowNo)) = "" Then
            If (Trim$(smSave(13, llRowNo)) = "HI") Or (Trim$(smSave(13, llRowNo)) = "IN") Or (Trim$(smSave(13, llRowNo)) = "AN") Then
                mTestFields = NO
                Exit Function
            End If
        End If
        If Trim$(smSave(12, llRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If imSave(1, llRowNo) < 0 Then
            mTestFields = NO
            Exit Function
        End If
    Next llRowNo
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestInvNo                      *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test or delete Sales History   *
'*                      with matching invoice #        *
'*                                                     *
'*******************************************************
Private Sub mTestInvNo(slInvNo As String, ilTestOrDelete As Integer, llRowNo As Long)
'
'   ilTestOrDelete(I)- 0=Test only; 1=Delete only
'   llRowNo(I)- Row number
'
    Dim ilAdfCode As Integer
    Dim ilAgfCode As Integer
    Dim llInvNo As Long
    Dim ilRes As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If rbcType(0).Value Then    'sales history
        Exit Sub
    End If
    If Trim$(slInvNo) = "" Then
        Exit Sub
    End If
    'disallow deletions 8/11/00
    If ilTestOrDelete = 1 Then
        Exit Sub
    End If
    llInvNo = Val(slInvNo)
    ilAdfCode = 0
    gFindMatch Trim$(smSave(2, llRowNo)), 2, lbcAdvertiser
    If gLastFound(lbcAdvertiser) > 1 Then
        slNameCode = tmAdvertiser(gLastFound(lbcAdvertiser) - 2).sKey    'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvertiser) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilAdfCode = Val(slCode)
    End If
    ilAgfCode = 0
    If Trim$(smSave(1, llRowNo)) <> "[Direct]" Then
        gFindMatch Trim$(smSave(1, llRowNo)), 2, lbcAgency
        If gLastFound(lbcAgency) > 1 Then
            slNameCode = tgAgency(gLastFound(lbcAgency) - 2).sKey    'Traffic!lbcAgency.List(gLastFound(lbcAgency) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAgfCode = Val(slCode)
        End If
    End If
    If (ilAdfCode <> 0) Then
        tmRvfPhfSrchKey1.iAdfCode = ilAdfCode
        ilRet = btrGetEqual(hmPhf, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmRvfPhf.iAdfCode = ilAdfCode)
            If (tmRvfPhf.lInvNo = llInvNo) And (tmRvfPhf.sTranType = "HI") Then
                If ilTestOrDelete <> 1 Then
                    ilRes = MsgBox("Duplicate History for this Invoice will be Removed", vbOKOnly + vbExclamation, "Warning")
                    Exit Sub
                Else
                    ilRet = btrDelete(hmPhf)
                End If
            End If
            ilRet = btrGetNext(hmPhf, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    ElseIf ilAgfCode <> 0 Then
        tmRvfPhfSrchKey0.iAgfCode = ilAgfCode
        ilRet = btrGetEqual(hmPhf, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmRvfPhf.iAgfCode = ilAgfCode)
            If (tmRvfPhf.lInvNo = llInvNo) And (tmRvfPhf.sTranType = "HI") Then
                If ilTestOrDelete <> 1 Then
                    ilRes = MsgBox("Duplicate History for this Invoice will be Removed", vbOKOnly + vbExclamation, "Warning")
                    Exit Sub
                Else
                    ilRet = btrDelete(hmPhf)
                End If
            End If
            ilRet = btrGetNext(hmPhf, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    Else
        'This case should not happen
    End If
End Sub
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
    Dim ilLoop As Integer
    Dim ilFound As Integer
    If (Trim$(smSave(2, llRowNo)) = "") Or (Trim$(smSave(2, llRowNo)) = "[None]") Then
        If Trim$(smSave(13, llRowNo)) <> "PO" Then
            Beep
            ilRes = MsgBox("Advertiser must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ADVTINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smSave(2, llRowNo)) <> "" Then
        mGetAdvt Trim$(smSave(2, llRowNo))
    End If
    If (Trim$(smSave(1, llRowNo)) = "") And (tmAdf.sBillAgyDir <> "D") Then
        Beep
        ilRes = MsgBox("Agency must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = AGYINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(4, llRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Salesperson must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = SPERSONINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(13, llRowNo)) <> "PO" Then
        If Trim$(smSave(5, llRowNo)) = "" Then
            Beep
            ilRes = MsgBox("Invoice # must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = INVNOINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If Trim$(smSave(7, llRowNo)) = "" Then
            Beep
            ilRes = MsgBox("Billing Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = BVEHINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If Trim$(smSave(8, llRowNo)) = "" Then
            Beep
            ilRes = MsgBox("Airing Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = AVEHINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smSave(10, llRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Transaction Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = TRANDATEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(12, llRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Net must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = NETINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(11, llRowNo)) = "" Then
        If (Trim$(smSave(13, llRowNo)) = "HI") Or (Trim$(smSave(13, llRowNo)) = "IN") Or (Trim$(smSave(13, llRowNo)) = "AN") Then
            Beep
            ilRes = MsgBox("Gross must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = GROSSINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    'Check Signs
    'If (InStr(smSave(12, llRowNo), "-") > 0) And (gStrDecToLong(smSave(11, llRowNo), 2) = 0) Then
    If (gStrDecToLong(smSave(12, llRowNo), 2) <> 0) And (gStrDecToLong(smSave(11, llRowNo), 2) = 0) And (Trim$(smSave(13, llRowNo)) = "PI") Then
        'PI- Test if IN defined to get ageing info from
        ilFound = False
        For ilLoop = 1 To UBound(smSave, 2) - 1 Step 1
            If ilLoop <> llRowNo Then
                'If (InStr(smSave(12, ilLoop), "-") > 0) And (gStrDecToLong(smSave(11, ilLoop), 2) = 0) Then
                If (gStrDecToLong(smSave(12, ilLoop), 2) <> 0) And (gStrDecToLong(smSave(11, ilLoop), 2) = 0) Then
                Else
                    If smSave(5, llRowNo) = smSave(5, ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            ilRes = MsgBox("No Invoice found for Payment", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = INVNOINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    'Check Signs
    If (gStrDecToLong(smSave(12, llRowNo), 2) <> 0) And (gStrDecToLong(smSave(11, llRowNo), 2) <> 0) Then
        If Left$(Trim$(smSave(11, llRowNo)), 1) = "-" Then
            If Left$(Trim$(smSave(12, llRowNo)), 1) <> "-" Then
                Beep
                ilRes = MsgBox("Gross & Net must have the same Sign", vbOKOnly + vbExclamation, "Incomplete")
                imBoxNo = GROSSINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        End If
        If Left$(Trim$(smSave(12, llRowNo)), 1) = "-" Then
            If Left$(Trim$(smSave(11, llRowNo)), 1) <> "-" Then
                Beep
                ilRes = MsgBox("Gross & Net must have the same Sign", vbOKOnly + vbExclamation, "Incomplete")
                imBoxNo = GROSSINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If
    If imSave(1, llRowNo) < 0 Then
        Beep
        ilRes = MsgBox("Cash/Trade must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = CTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
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
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehicle.ListIndex
    If ilIndex >= 0 Then
        slName = lbcVehicle.List(ilIndex)
    End If
    'ilRet = gPopUserVehicleBox(SaleHist, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(SaleHist, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SaleHist
        On Error GoTo 0
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
            Else
                lbcVehicle.ListIndex = -1
            End If
        Else
            lbcVehicle.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcCT_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcCT_KeyPress(KeyAscii As Integer)
    If imBoxNo = CTINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imSave(1, lmRowNo) <> 0 Then
                imPhfChg = True
            End If
            imSave(1, lmRowNo) = 0
            pbcCT_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imSave(1, lmRowNo) <> 1 Then
                imPhfChg = True
            End If
            imSave(1, lmRowNo) = 1
            pbcCT_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imBoxNo = CTINDEX Then
            If imSave(1, lmRowNo) = 0 Then
                imPhfChg = True
                imSave(1, lmRowNo) = 1
            Else
                imPhfChg = True
                imSave(1, lmRowNo) = 0
            End If
        End If
        pbcCT_Paint
    End If
End Sub
Private Sub pbcCT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = CTINDEX Then
        If imSave(1, lmRowNo) = 0 Then
            imPhfChg = True
            imSave(1, lmRowNo) = 1
        Else
            imPhfChg = True
            imSave(1, lmRowNo) = 0
        End If
    End If
    pbcCT_Paint
End Sub
Private Sub pbcCT_Paint()
    pbcCT.Cls
    pbcCT.CurrentX = fgBoxInsetX
    pbcCT.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = CTINDEX Then
        If imSave(1, lmRowNo) = 0 Then
            pbcCT.Print "Cash"
        ElseIf imSave(1, lmRowNo) = 1 Then
            pbcCT.Print "Trade"
        End If
    End If
End Sub
Private Sub pbcHist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcHist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llMaxRow As Long
    Dim llCompRow As Long
    Dim llRow As Long
    Dim llRowNo As Long
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slTax As String
    Dim slStr As String

    If Button = 2 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    llCompRow = vbcHist.LargeChange + 1
    If UBound(tgPhfRec) > llCompRow Then
        llMaxRow = llCompRow
    Else
        llMaxRow = UBound(tgPhfRec) + 1
    End If
    For llRow = 1 To llMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((llRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((llRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    llRowNo = llRow + vbcHist.Value - 1
                    If llRowNo > UBound(smSave, 2) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    'PO don't require Advertiser
                    'If (ilBox > ADVTINDEX) And (Trim$(smSave(2, llRowNo)) = "") Then
                    '    Beep
                    '    mSetFocus imBoxNo
                    '    Exit Sub
                    'End If
                    If (tgSpf.sUsingNTR <> "Y") And ((ilBox = NTRTYPEINDEX) Or (ilBox = NTRTAXINDEX)) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = NTRTAXINDEX) Then
                        If (Not imTaxDefined) Then
                            Beep
                            Exit Sub
                        End If
                        If (imSave(2, llRowNo) > 0) And (lmSave(1, llRowNo) <= 0) Then
                            slNameCode = tmNTRTypeCode(imSave(2, llRowNo) - 1).sKey 'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 6, "\", slTax)
                            If ilRet = CP_MSG_NONE Then
                                If slTax <> "Y" Then
                                    Beep
                                    Exit Sub
                                End If
                            End If
                        Else
                            Beep
                            Exit Sub
                        End If
                    End If
                    If (Not mUsingAcqCost()) And (ilBox = ACQCOSTINDEX) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If ilBox = TRANTYPEINDEX Then
                        If (tgPhfRec(llRowNo).iStatus <> 0) Or ((rbcType(0).Value) And (imSelectedIndex <= 1) And (Trim$(edcCntrNo.Text) = "")) Then  'New selected
                            If Trim$(smSave(13, llRowNo)) = "" And (rbcType(0).Value) Then
                                smSave(13, llRowNo) = "HI"
                                slStr = "HI"
                                gSetShow pbcHist, slStr, tmCtrls(TRANTYPEINDEX)
                                smShow(TRANTYPEINDEX, llRowNo) = tmCtrls(TRANTYPEINDEX).sShow
                            End If
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    End If
                    If (ilBox = GROSSINDEX) And ((Trim$(smSave(13, llRowNo)) = "PI") Or (Left$(smSave(13, llRowNo), 1) = "W") Or (Trim$(smSave(13, llRowNo)) = "PO")) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    'The SetShow must be prior to mSSPart call in case focus is on ss/part field and lbcSSPart changed with pop call
                    mSetShow imBoxNo
                    If ilBox = SSPARTINDEX Then
                        mSSPartPop llRowNo, 0, 0, 0, tgPhfRec(llRowNo).iStatus, smSave(10, llRowNo)
                        If lbcSSPart.ListCount <= 1 Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    End If

                    lmRowNo = llRow + vbcHist.Value - 1
                    If (lmRowNo = UBound(smSave, 2)) And (Trim$(smSave(2, lmRowNo)) = "") Then
                        mInitNew lmRowNo
                    End If
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next llRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcHist_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    mPaintHistTitle
    ilStartRow = vbcHist.Value '+ 1  'Top location
    ilEndRow = vbcHist.Value + vbcHist.LargeChange ' + 1
    If ilEndRow > UBound(smSave, 2) Then
        If Trim$(smShow(1, UBound(smShow, 2))) <> "" Then
            ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smSave, 2) - 1
        End If
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcHist.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcHist.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15)  '- 30 '+ fgBoxInsetY
            slStr = Trim$(smShow(ilBox, ilRow))
            pbcHist.Print slStr
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    If imBoxNo = ADVTINDEX Then
        If mAdvtBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = AGYINDEX Then
        If mAgencyBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smSave, 2) = 1) Then
                imTabDirection = 0  'Set-Left to right
                lmRowNo = 1
                mInitNew lmRowNo
            Else
                If UBound(smSave, 2) <= vbcHist.LargeChange Then 'was <=
                    vbcHist.Max = LBONE 'LBound(smSave, 2)
                Else
                    vbcHist.Max = UBound(smSave, 2) - vbcHist.LargeChange '- 1
                End If
                lmRowNo = 1
                If lmRowNo >= UBound(smSave, 2) Then
                    mInitNew lmRowNo
                End If
                imSettingValue = True
                vbcHist.Value = vbcHist.Min
                imSettingValue = False
            End If
            ilBox = CTINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case CTINDEX, 0
            mSetShow imBoxNo
            If (imBoxNo < 1) And (lmRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = NETINDEX
            If lmRowNo <= 1 Then
                imBoxNo = -1
                lmRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            lmRowNo = lmRowNo - 1
            If lmRowNo < vbcHist.Value Then
                imSettingValue = True
                vbcHist.Value = vbcHist.Value - 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case NETINDEX
            If (tgSpf.sUsingNTR = "Y") Then
                ilBox = NTRTYPEINDEX
            Else
                ilBox = TRANDATEINDEX
            End If
        Case GROSSINDEX
            If (tgSpf.sUsingNTR = "Y") Then
                ilBox = NTRTYPEINDEX
            Else
                ilBox = TRANDATEINDEX
            End If
        Case SSPARTINDEX
            If mUsingAcqCost() Then
                ilBox = ACQCOSTINDEX
            Else
                ilBox = NETINDEX
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slTax As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If inGenMsgRet = True Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    If imBoxNo = ADVTINDEX Then
        If mAdvtBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = AGYINDEX Then
        If mAgencyBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SPERSONINDEX Then
        If mSPersonBranch() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            lmRowNo = UBound(smSave, 2) - 1
            imSettingValue = True
            If lmRowNo <= vbcHist.LargeChange + 1 Then
                vbcHist.Value = 1
            Else
                vbcHist.Value = lmRowNo - vbcHist.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = NETINDEX
        Case NETINDEX, ACQCOSTINDEX, SSPARTINDEX
            mSetShow imBoxNo
            If imBoxNo = NETINDEX Then
                If mUsingAcqCost() Then
                    ilBox = imBoxNo + 1
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
                mSSPartPop lmRowNo, 0, 0, 0, tgPhfRec(lmRowNo).iStatus, smSave(10, lmRowNo)
                If lbcSSPart.ListCount > 1 Then
                    ilBox = imBoxNo + 1
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            ElseIf imBoxNo = ACQCOSTINDEX Then
                mSSPartPop lmRowNo, 0, 0, 0, tgPhfRec(lmRowNo).iStatus, smSave(10, lmRowNo)
                If lbcSSPart.ListCount > 1 Then
                    ilBox = imBoxNo + 1
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
            If mTestSaveFields(lmRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If

            If lmRowNo >= UBound(smSave, 2) Then
                imPhfChg = True
                ReDim Preserve smShow(0 To 18, 0 To lmRowNo + 1) As String * 50 'Values shown in program area
                ReDim Preserve smSave(0 To 15, 0 To lmRowNo + 1) As String * 60 'Values saved (program name) in program area
                ReDim Preserve imSave(0 To 4, 0 To lmRowNo + 1) As Integer 'Values saved (program name) in program area
                ReDim Preserve lmSave(0 To 1, 0 To lmRowNo + 1) As Long 'Values saved (program name) in program area
                For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                    smShow(ilLoop, lmRowNo + 1) = ""
                Next ilLoop
                For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                    smSave(ilLoop, lmRowNo + 1) = ""
                Next ilLoop
                For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                    imSave(ilLoop, lmRowNo + 1) = -1
                Next ilLoop
                For ilLoop = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                    lmSave(ilLoop, lmRowNo + 1) = 0
                Next ilLoop
                ReDim Preserve tgPhfRec(0 To UBound(tgPhfRec) + 1) As PHFREC
                tgPhfRec(UBound(tgPhfRec)).iStatus = 0
                tgPhfRec(UBound(tgPhfRec)).lRecPos = 0
            End If
            If lmRowNo >= UBound(smSave, 2) - 1 Then
                lmRowNo = lmRowNo + 1
                mInitNew lmRowNo
                If UBound(smSave, 2) <= vbcHist.LargeChange Then 'was <=
                    vbcHist.Max = LBONE 'LBound(smSave, 2) '- 1
                Else
                    vbcHist.Max = UBound(smSave, 2) - vbcHist.LargeChange '- 1
                End If
            Else
                lmRowNo = lmRowNo + 1
            End If
            If lmRowNo > vbcHist.Value + vbcHist.LargeChange Then
                imSettingValue = True
                vbcHist.Value = vbcHist.Value + 1
                imSettingValue = False
            End If
            If lmRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                mSetCommands
                'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (lmRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcHist.Top + tmCtrls(CTINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = CTINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = CTINDEX
        Case TRANDATEINDEX
            If (tgSpf.sUsingNTR <> "Y") Then
                If Trim$(smSave(13, lmRowNo)) = "" And (rbcType(0).Value) Then
                    smSave(13, lmRowNo) = "HI"
                    slStr = "HI"
                    gSetShow pbcHist, slStr, tmCtrls(TRANTYPEINDEX)
                    smShow(TRANTYPEINDEX, lmRowNo) = tmCtrls(TRANTYPEINDEX).sShow
                End If
                If (tgPhfRec(lmRowNo).iStatus <> 0) Or ((rbcType(0).Value) And (imSelectedIndex <= 1) And (Trim$(edcCntrNo.Text) = "")) Then  'New selected
                    If (Trim$(smSave(13, lmRowNo)) = "PI") Or (Trim$(smSave(13, lmRowNo)) = "PO") Or (Left$(smSave(13, lmRowNo), 1) = "W") Then
                        ilBox = NETINDEX
                    Else
                        ilBox = GROSSINDEX
                    End If
                Else
                    ilBox = TRANTYPEINDEX
                End If
            Else
                ilBox = imBoxNo + 1
            End If
        Case TRANTYPEINDEX
            mSetShow imBoxNo
            If (smSave(13, lmRowNo) = "IN") Or (smSave(13, lmRowNo) = "AN") Or (smSave(13, lmRowNo) = "HI") Then
                ilBox = imBoxNo + 1
            Else
                ilBox = NETINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case NTRTYPEINDEX
            If Trim$(smSave(13, lmRowNo)) = "" And (rbcType(0).Value) Then
                smSave(13, lmRowNo) = "HI"
                slStr = "HI"
                gSetShow pbcHist, slStr, tmCtrls(TRANTYPEINDEX)
                smShow(TRANTYPEINDEX, lmRowNo) = tmCtrls(TRANTYPEINDEX).sShow
            End If
            If (tgPhfRec(lmRowNo).iStatus <> 0) Then   'New selected
                If (Trim$(smSave(13, lmRowNo)) = "PI") Or (Left$(smSave(13, lmRowNo), 1) = "W") Or (Trim$(smSave(13, lmRowNo)) = "PO") Then
                    ilBox = NETINDEX
                Else
                    ilBox = GROSSINDEX
                End If
            Else
                If Not imTaxDefined Then
                    ilBox = TRANTYPEINDEX
                Else
                    ilBox = NTRTAXINDEX
                    If (lbcNTRType.ListIndex > 0) And (lmSave(1, lmRowNo) <= 0) Then
                        slNameCode = tmNTRTypeCode(lbcNTRType.ListIndex - 1).sKey 'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 6, "\", slTax)
                        If ilRet = CP_MSG_NONE Then
                            If slTax <> "Y" Then
                                ilBox = TRANTYPEINDEX
                            End If
                        End If
                    Else
                        ilBox = TRANTYPEINDEX
                    End If
                End If
                If (rbcType(0).Value) And (ilBox = TRANTYPEINDEX) And (imSelectedIndex <= 1) And (Trim$(edcCntrNo.Text) = "") Then
                    If (Trim$(smSave(13, lmRowNo)) = "PI") Or (Left$(smSave(13, lmRowNo), 1) = "W") Or (Trim$(smSave(13, lmRowNo)) = "PO") Then
                        ilBox = NETINDEX
                    Else
                        ilBox = GROSSINDEX
                    End If
                End If
            End If
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcTT_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcTT_KeyPress(KeyAscii As Integer)
    If imBoxNo = TRANTYPEINDEX Then
        If (KeyAscii = Asc("I")) Or (KeyAscii = Asc("i")) Then
            If Trim$(smSave(13, lmRowNo)) <> "IN" Then
                imPhfChg = True
            End If
            smSave(13, lmRowNo) = "IN"
            pbcTT_Paint
        ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
            If Trim$(smSave(13, lmRowNo)) <> "PI" Then
                imPhfChg = True
            End If
            smSave(13, lmRowNo) = "PI"
            pbcTT_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imBoxNo = TRANTYPEINDEX Then
            If Trim$(smSave(13, lmRowNo)) = "IN" Then
                imPhfChg = True
                smSave(13, lmRowNo) = "PI"
            Else
                imPhfChg = True
                smSave(13, lmRowNo) = "IN"
            End If
        End If
        pbcTT_Paint
    End If
End Sub

Private Sub pbcTT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = TRANTYPEINDEX Then
        If Trim$(smSave(13, lmRowNo)) = "IN" Then
            imPhfChg = True
            smSave(13, lmRowNo) = "PI"
        Else
            imPhfChg = True
            smSave(13, lmRowNo) = "IN"
        End If
    End If
    pbcTT_Paint

End Sub

Private Sub pbcTT_Paint()
    pbcTT.Cls
    pbcTT.CurrentX = fgBoxInsetX
    pbcTT.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = TRANTYPEINDEX Then
        If Trim$(smSave(13, lmRowNo)) = "IN" Then
            pbcTT.Print "IN"
        ElseIf Trim$(smSave(13, lmRowNo)) = "PI" Then
            pbcTT.Print "PI"
        End If
    End If
End Sub

Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If (imSelectedIndex >= 0) And Value Then
        Screen.MousePointer = vbHourglass
        If imSelectedIndex > 1 Then
            slNameCode = lbcAdvtAgyCode.List(imSelectedIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If (InStr(slNameCode, "/Direct") > 0) Or (InStr(slNameCode, "/Non-Payee") > 0) Then
                If Not mReadRvfPhfRec(Val(slCode), 0, "") Then
                End If
            Else
                If Not mReadRvfPhfRec(0, Val(slCode), "") Then
                End If
            End If
        ElseIf imSelectedIndex = 1 Then
            If Not mReadRvfPhfRec(0, 0, "") Then
            End If
        Else
            ReDim tgPhfRec(0 To 1) As PHFREC
            tgPhfRec(1).iStatus = -1
            tgPhfRec(1).lRecPos = 0
            ReDim tgPhfDel(0 To 0) As PHFREC
            tgPhfDel(0).iStatus = -1
            tgPhfDel(0).lRecPos = 0
        End If
        pbcHist.Cls
        If imSelectedIndex >= 1 Then
            mMoveRecToCtrl
            mInitShow
        Else
            mClearCtrlFields
        End If
        'pbcHist_Paint
        mSetMinMax
        mSetCommands
        Screen.MousePointer = vbDefault
    ElseIf (imSelectedIndex = -1) And (Trim$(edcCntrNo.Text) <> "") And Value Then
        Screen.MousePointer = vbHourglass
        If Not mReadRvfPhfRec(0, 0, Trim$(edcCntrNo.Text)) Then
        End If
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
        Screen.MousePointer = vbDefault
    End If
    If (Value) And (cbcSelect.ListCount >= 2) Then
        If Index = 0 Then
            cbcSelect.List(0) = "[New HI]"
            cbcSelect.List(1) = "[All HI]"
        Else
            cbcSelect.List(0) = "[New]"
            cbcSelect.List(1) = "[All]"
        End If
    End If
    If Value Then
        mTranTypePop
        If lbcTranType.ListCount <= 0 Then
            MsgBox "Transaction Types missing, please exit Backlog and create from List Item Tran Type"
            Exit Sub
        End If
    End If
End Sub

Private Sub rbcType_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case ADVTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAdvertiser, edcDropDown, imChgMode, imLbcArrowSetting
        Case PRODINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProduct, edcDropDown, imChgMode, imLbcArrowSetting
        Case AGYINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAgency, edcDropDown, imChgMode, imLbcArrowSetting
        Case SPERSONINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSPerson, edcDropDown, imChgMode, imLbcArrowSetting
        Case BVEHINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcBVehicle, edcDropDown, imChgMode, imLbcArrowSetting
        Case AVEHINDEX
        Case NTRTYPEINDEX  'Vehicle
            imLbcArrowSetting = False
            gProcessLbcClick lbcNTRType, edcDropDown, imChgMode, imLbcArrowSetting
        Case SSPARTINDEX  'Vehicle
            imLbcArrowSetting = False
            gProcessLbcClick lbcSSPart, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub

Private Sub tmcCntrNo_Timer()
    Dim ilRet As Integer
    
    tmcCntrNo.Enabled = False
    If imSelectedIndex <> -1 Then
        '1/3/18: Client was failing with this call sometimes.
        'pbcSTab.SetFocus
        Exit Sub
    End If
    If edcCntrNo.Text <> "" Then
        Screen.MousePointer = vbHourglass
        ilRet = mReadRvfPhfRec(0, 0, Trim$(edcCntrNo.Text))
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
        Screen.MousePointer = vbDefault
        'pbcSTab.SetFocus
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub tmcDrag_Timer()
    Dim llCompRow As Long
    Dim llMaxRow As Long
    Dim llRow As Long
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            llCompRow = vbcHist.LargeChange + 1
            If UBound(smSave, 2) > llCompRow Then
                llMaxRow = llCompRow
            Else
                llMaxRow = UBound(smSave, 2)
            End If
            For llRow = 1 To llMaxRow Step 1
                If (fmDragY >= ((llRow - 1) * (fgBoxGridH + 15) + tmCtrls(CTINDEX).fBoxY)) And (fmDragY <= ((llRow - 1) * (fgBoxGridH + 15) + tmCtrls(CTINDEX).fBoxY + tmCtrls(CTINDEX).fBoxH)) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    lmRowNo = -1
                    lmRowNo = llRow + vbcHist.Value - 1
                    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame.Move 0, tmCtrls(CTINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcHist.Top + tmCtrls(CTINDEX).fBoxY + (lmRowNo - vbcHist.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame.Drag vbBeginDrag
                    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next llRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    cmcCancel_Click
End Sub

Private Sub vbcHist_Change()
    If imSettingValue Then
        pbcHist.Cls
        pbcHist_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        imBoxNo = -1
        lmRowNo = -1
        pbcHist.Cls
        pbcHist_Paint
        'If (igWinStatus(INVOICESJOB) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        '    mEnableBox imBoxNo
        'End If
    End If
End Sub
Private Sub vbcHist_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    lmRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Backlog"
End Sub

Private Sub mBuildKey()

    Dim slStr As String
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim slCntrNo As String
    Dim slDate As String
    Dim llLoop As Long

    For llUpper = LBONE To UBound(tgPhfRec) - 1 Step 1
        slStr = ""
        If tgPhfRec(llUpper).tPhf.iAdfCode <> tmAdf.iCode Then
            tmAdfSrchKey.iCode = tgPhfRec(llUpper).tPhf.iAdfCode 'ilCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet = BTRV_ERR_NONE Then
            If tgPhfRec(llUpper).tPhf.lPrfCode > 0 Then
                If tgPhfRec(llUpper).tPhf.lPrfCode <> tmPrf.lCode Then
                    tmPrfSrchKey0.lCode = tgPhfRec(llUpper).tPhf.lPrfCode 'ilCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    tmPrf.sName = "Product Missing"
                End If
            Else
                tmPrf.sName = ""
            End If
        Else
            tmAdf.sName = "Advertiser Missing"
            tmPrf.sName = ""
        End If
        slCntrNo = Trim$(str$(tgPhfRec(llUpper).tPhf.lCntrNo))
        Do While Len(slCntrNo) < 10
            slCntrNo = "0" & slCntrNo
        Loop
        gUnpackDateForSort tgPhfRec(llUpper).tPhf.iTranDate(0), tgPhfRec(llUpper).tPhf.iTranDate(1), slDate
        tgPhfRec(llUpper).sKey = tmAdf.sName & tmPrf.sName & slCntrNo & slDate
        tgPhfRec(llUpper).iStatus = 1
        tgPhfRec(llUpper).sProduct = tmPrf.sName
    Next llUpper
    llUpper = UBound(tgPhfRec)
    If llUpper > 1 Then
        'ArraySortTyp fnAV(tgPhfRec(), 1), UBound(tgPhfRec) - 1, 0, LenB(tgPhfRec(1)), 0, LenB(tgPhfRec(1).sKey), 0
        For llLoop = LBound(tgPhfRec) To UBound(tgPhfRec) - 1 Step 1
            tgPhfRec(llLoop) = tgPhfRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tgPhfRec(0 To UBound(tgPhfRec) - 1) As PHFREC
        ArraySortTyp fnAV(tgPhfRec(), 0), UBound(tgPhfRec), 0, LenB(tgPhfRec(0)), 0, LenB(tgPhfRec(0).sKey), 0
        ReDim Preserve tgPhfRec(0 To UBound(tgPhfRec) + 1) As PHFREC
        For llLoop = UBound(tgPhfRec) - 1 To LBound(tgPhfRec) Step -1
            tgPhfRec(llLoop + 1) = tgPhfRec(llLoop)
        Next llLoop
    End If

End Sub

Private Function mUsingAcqCost() As Integer

    mUsingAcqCost = False
    '6/7/15: replaced acquisition from site override with Barter in system options
    If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
        mUsingAcqCost = True
        Exit Function
    End If
    If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
        mUsingAcqCost = True
        Exit Function
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintHistTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcHist.ForeColor
    slFontName = pbcHist.FontName
    flFontSize = pbcHist.FontSize
    ilFillStyle = pbcHist.FillStyle
    llFillColor = pbcHist.FillColor
    pbcHist.ForeColor = BLUE
    pbcHist.FontBold = False
    pbcHist.FontSize = 7
    pbcHist.FontName = "Arial"
    pbcHist.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmCtrls(CTINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcHist.Line (tmCtrls(CTINDEX).fBoxX - 15, 15)-Step(tmCtrls(CTINDEX).fBoxW + 15, tmCtrls(CTINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(CTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcHist.Print "C"
    pbcHist.CurrentX = tmCtrls(CTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "T"
    pbcHist.Line (tmCtrls(ADVTINDEX).fBoxX - 15, 15)-Step(tmCtrls(ADVTINDEX).fBoxW + 15, tmCtrls(ADVTINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(ADVTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcHist.Print "Advertiser"
    pbcHist.Line (tmCtrls(PRODINDEX).fBoxX - 15, 15)-Step(tmCtrls(PRODINDEX).fBoxW + 15, tmCtrls(PRODINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(PRODINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcHist.Print "Product"
    pbcHist.Line (tmCtrls(AGYINDEX).fBoxX - 15, 15)-Step(tmCtrls(AGYINDEX).fBoxW + 15, tmCtrls(AGYINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(AGYINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Agency"
    pbcHist.Line (tmCtrls(SPERSONINDEX).fBoxX - 15, 15)-Step(tmCtrls(SPERSONINDEX).fBoxW + 15, tmCtrls(SPERSONINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(SPERSONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Sales-"
    pbcHist.CurrentX = tmCtrls(SPERSONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "person"
    pbcHist.Line (tmCtrls(INVNOINDEX).fBoxX - 15, 15)-Step(tmCtrls(INVNOINDEX).fBoxW + 15, tmCtrls(INVNOINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(INVNOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcHist.Print "Inv"
    pbcHist.CurrentX = tmCtrls(INVNOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "#"
    pbcHist.Line (tmCtrls(CNTRNOINDEX).fBoxX - 15, 15)-Step(tmCtrls(CNTRNOINDEX).fBoxW + 15, tmCtrls(CNTRNOINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(CNTRNOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcHist.Print "Contr"
    pbcHist.CurrentX = tmCtrls(CNTRNOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "#"
    pbcHist.Line (tmCtrls(BVEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(BVEHINDEX).fBoxW + 15, tmCtrls(BVEHINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(BVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Bill"
    pbcHist.CurrentX = tmCtrls(BVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Vehicle"
    pbcHist.Line (tmCtrls(AVEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(AVEHINDEX).fBoxW + 15, tmCtrls(AVEHINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(AVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Aired"
    pbcHist.CurrentX = tmCtrls(AVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Vehicle"
    pbcHist.Line (tmCtrls(PKLNINDEX).fBoxX - 15, 15)-Step(tmCtrls(PKLNINDEX).fBoxW + 15, tmCtrls(PKLNINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(PKLNINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Pkg"
    pbcHist.CurrentX = tmCtrls(PKLNINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Ln"
    pbcHist.Line (tmCtrls(TRANDATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(TRANDATEINDEX).fBoxW + 15, tmCtrls(TRANDATEINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(TRANDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Tran"
    pbcHist.CurrentX = tmCtrls(TRANDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Date"
    pbcHist.Line (tmCtrls(NTRTYPEINDEX).fBoxX - 15, 15)-Step(tmCtrls(NTRTYPEINDEX).fBoxW + 15, tmCtrls(NTRTYPEINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(NTRTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "NTR"
    pbcHist.CurrentX = tmCtrls(NTRTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Type"
    pbcHist.Line (tmCtrls(NTRTAXINDEX).fBoxX - 15, 15)-Step(tmCtrls(NTRTAXINDEX).fBoxW + 15, tmCtrls(NTRTAXINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(NTRTAXINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "T"
    pbcHist.CurrentX = tmCtrls(NTRTAXINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "X"
    pbcHist.Line (tmCtrls(TRANTYPEINDEX).fBoxX - 15, 15)-Step(tmCtrls(TRANTYPEINDEX).fBoxW + 15, tmCtrls(TRANTYPEINDEX).fBoxY - 30), BLUE, B
    pbcHist.Line (tmCtrls(TRANTYPEINDEX).fBoxX, 30)-Step(tmCtrls(TRANTYPEINDEX).fBoxW - 15, tmCtrls(TRANTYPEINDEX).fBoxY - 45), LIGHTBLUE, BF
    pbcHist.CurrentX = tmCtrls(TRANTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "T"
    pbcHist.CurrentX = tmCtrls(TRANTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "T"
    pbcHist.Line (tmCtrls(GROSSINDEX).fBoxX - 15, 15)-Step(tmCtrls(GROSSINDEX).fBoxW + 15, tmCtrls(GROSSINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(GROSSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Gross"
    pbcHist.Line (tmCtrls(NETINDEX).fBoxX - 15, 15)-Step(tmCtrls(NETINDEX).fBoxW + 15, tmCtrls(NETINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(NETINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Net"
    pbcHist.Line (tmCtrls(ACQCOSTINDEX).fBoxX - 15, 15)-Step(tmCtrls(ACQCOSTINDEX).fBoxW + 15, tmCtrls(ACQCOSTINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(ACQCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "Acq"
    pbcHist.CurrentX = tmCtrls(ACQCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Cost"
    pbcHist.Line (tmCtrls(SSPARTINDEX).fBoxX - 15, 15)-Step(tmCtrls(SSPARTINDEX).fBoxW + 15, tmCtrls(SSPARTINDEX).fBoxY - 30), BLUE, B
    pbcHist.CurrentX = tmCtrls(SSPARTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = 15
    pbcHist.Print "SS/"
    pbcHist.CurrentX = tmCtrls(SSPARTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcHist.CurrentY = ilHalfY + 15
    pbcHist.Print "Part"

    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            pbcHist.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            If ilLoop = TRANTYPEINDEX Then
                pbcHist.Line (tmCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmCtrls(ilLoop).fBoxW - 15, tmCtrls(ilLoop).fBoxH - 30), LIGHTBLUE, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + tmCtrls(1).fBoxH < pbcHist.Height
    vbcHist.LargeChange = ilLineCount - 1
    pbcHist.FontSize = flFontSize
    pbcHist.FontName = slFontName
    pbcHist.FontSize = flFontSize
    pbcHist.ForeColor = llColor
    pbcHist.FontBold = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTranTypePop                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Transaction Types     *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mTranTypePop()
'
'   mAgyDPPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer

    ilIndex = lbcTranType.ListIndex
    If ilIndex >= 0 Then
        slName = lbcTranType.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(Collect, lbcTranType, lbcTranTypeCode, "YW")
    If (rbcType(0).Value) And (imSelectedIndex <= 1) And (Trim$(edcCntrNo.Text) = "") Then
        lbcTranType.Clear
        lbcTranType.AddItem "HI Invoice"
    Else
        smTranTypeCodeTag = ""
        ilRet = gPopMnfPlusFieldsBox(SaleHist, lbcTranType, tmTranTypeCode(), smTranTypeCodeTag, "Y")
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove PO since gross and net will be zero
        If rbcType(0).Value Then
            For ilLoop = 0 To lbcTranType.ListCount - 1 Step 1
                If Left$(lbcTranType.List(ilLoop), 2) = "PO" Then
                    lbcTranType.RemoveItem (ilLoop)
                    Exit For
                End If
            Next ilLoop
        End If
        On Error GoTo mTranTypePopErr
        gCPErrorMsg ilRet, "mTranTypePop (gPopMnfPlusFieldsBox)", SaleHist
        On Error GoTo 0
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcTranType
            If gLastFound(lbcTranType) >= 0 Then
                lbcTranType.ListIndex = gLastFound(lbcTranType)
            Else
                lbcTranType.ListIndex = -1
            End If
        Else
            lbcTranType.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mTranTypePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
