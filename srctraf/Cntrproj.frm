VERSION 5.00
Begin VB.Form CntrProj 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   105
   ClientTop       =   1410
   ClientWidth     =   9330
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
   ScaleWidth      =   9330
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   180
      Left            =   45
      Picture         =   "CntrProj.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   105
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
      Left            =   5820
      TabIndex        =   1
      Top             =   120
      Width           =   3450
   End
   Begin VB.CommandButton cmcBlock 
      Appearance      =   0  'Flat
      Caption         =   "&Block"
      Height          =   285
      Left            =   6735
      TabIndex        =   43
      Top             =   5385
      Width           =   1050
   End
   Begin VB.TextBox edcComment 
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
      Height          =   780
      HelpContextID   =   8
      Left            =   5145
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   1350
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.ListBox lbcNR 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3630
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lbcProd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1515
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3810
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcPropNo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1545
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox lbcSOffice 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2610
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcPot 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6795
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3420
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7875
      Top             =   5280
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8220
      Top             =   5295
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
      Height          =   225
      Left            =   5640
      ScaleHeight     =   225
      ScaleWidth      =   3000
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3000
      Begin VB.OptionButton rbcType 
         Caption         =   "Quarter"
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Month"
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
         Left            =   1020
         TabIndex        =   39
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Week"
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
         Left            =   1935
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.PictureBox plcShow 
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   3225
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3225
      Begin VB.OptionButton rbcShow 
         Caption         =   "Corporate"
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
         Left            =   795
         TabIndex        =   35
         Top             =   0
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton rbcShow 
         Caption         =   "Standard"
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
         Left            =   2025
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
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
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
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
      Left            =   3195
      Picture         =   "CntrProj.frx":030A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2460
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2355
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3165
      Visible         =   0   'False
      Width           =   2685
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
      Left            =   4605
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   -15
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
      Left            =   4065
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   525
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
      Left            =   3660
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   525
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
      Height          =   1695
      Left            =   240
      Picture         =   "CntrProj.frx":0404
      ScaleHeight     =   1665
      ScaleWidth      =   3570
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.CheckBox ckcDiff 
      Caption         =   "Differences"
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
      Left            =   3765
      TabIndex        =   22
      Top             =   390
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox plcDetailSum 
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
      Left            =   900
      ScaleHeight     =   195
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   2085
      Begin VB.OptionButton rbcDetailSum 
         Caption         =   "Summary"
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
         Left            =   825
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   1140
      End
      Begin VB.OptionButton rbcDetailSum 
         Caption         =   "Detail"
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
         TabIndex        =   20
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo"
      Height          =   285
      Left            =   4545
      TabIndex        =   26
      Top             =   5385
      Width           =   1050
   End
   Begin VB.CommandButton cmcRollover 
      Appearance      =   0  'Flat
      Caption         =   "Roll&over"
      Height          =   285
      Left            =   5640
      TabIndex        =   27
      Top             =   5385
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
      Left            =   -15
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3555
      TabIndex        =   25
      Top             =   5385
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
      Height          =   105
      Left            =   -30
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   17
      Top             =   1050
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
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   585
      Width           =   60
   End
   Begin VB.VScrollBar vbcProj 
      Height          =   4155
      LargeChange     =   14
      Left            =   9000
      Min             =   1
      TabIndex        =   18
      Top             =   705
      Value           =   1
      Width           =   240
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2460
      TabIndex        =   24
      Top             =   5385
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   60
      ScaleHeight     =   270
      ScaleWidth      =   1260
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1260
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1380
      TabIndex        =   23
      Top             =   5385
      Width           =   1050
   End
   Begin VB.PictureBox pbcProj 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   240
      Picture         =   "CntrProj.frx":13ABA
      ScaleHeight     =   4185
      ScaleWidth      =   8775
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   8775
      Begin VB.PictureBox pbcLnWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   7890
         Picture         =   "CntrProj.frx":3BC20
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   45
         Top             =   150
         Width           =   270
      End
      Begin VB.PictureBox pbcLnWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4875
         Picture         =   "CntrProj.frx":3BED2
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   44
         Top             =   120
         Width           =   270
      End
      Begin VB.PictureBox pbcState 
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
         Height          =   795
         Left            =   1260
         ScaleHeight     =   795
         ScaleWidth      =   1395
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3375
         Width           =   1395
      End
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
         TabIndex        =   42
         Top             =   1365
         Visible         =   0   'False
         Width           =   8745
      End
   End
   Begin VB.PictureBox plcProj 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4335
      Left            =   180
      ScaleHeight     =   4275
      ScaleWidth      =   9045
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   9105
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8295
      TabIndex        =   28
      Top             =   4860
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   5175
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   90
      Picture         =   "CntrProj.frx":3C184
      Top             =   345
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8730
      Picture         =   "CntrProj.frx":3C48E
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "CntrProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Cntrproj.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CntrProj.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Projection input screen code
Option Explicit
Option Compare Text
Dim tmCtrls(0 To 13)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer
Dim smShow() As String  'Values shown (1=Prop No; 2=SalesOffice; 3=Advt; 4=Product;
                        ' 5=Demo; 6=Vehicle; 7=New/Return; 8=Potential; 9=Comment; 10-12=Dollars; 13=Total)
Dim smSave() As String  'Values saved (1=Proposal #; 2=SalesOffice; 3=Advertsier; 4=Product; 5=Demo Name; 6=Vehicle; 7=Potential; 8=Comment)
Dim lmSave() As Long    'Value saved (1-4=Dollars, 5=ChfCode)
Dim imSave() As Integer  'Values saved (1=Index into tgPjf1Rec; 2= New(0) or Return(1); 3=Changed Flag (True, False))
Dim tmSOfficeCode() As SORTCODE
Dim smSOfficeCodeTag As String
Dim tmAdvertiser() As SORTCODE
Dim smAdvertiserTag As String
'Current Week Advt/Prod
Dim tmCAPCtrls(0 To 4)  As FIELDAREA
Dim lmCAPSave(0 To 4) As Long
'Prior Week Advt/Prod
Dim tmPAPCtrls(0 To 4)  As FIELDAREA  'Original
Dim lmPAPSave(0 To 4) As Long
'Original Week Advt/Prod
Dim lmOAPSave(0 To 4) As Long
'Current Week Totals
Dim tmCTCtrls(0 To 4)  As FIELDAREA
Dim lmCTSave(0 To 4) As Long
'Prior Week Totals
Dim tmPTCtrls(0 To 4)  As FIELDAREA
Dim lmPTSave(0 To 4) As Long
'Original Week Totals
Dim lmOTSave(0 To 4) As Long
Dim tmSaleOffice() As PJSALEOFFICE
Dim tmUserVeh() As USERVEH
Dim imPjfChg As Integer  'True=Vehicle or salesperon value changed; False=No changes
Dim tmPjf As PJF        'Pjf record image
Dim tmPjfSrchKey As PJFKEY0    'Pjf key record image
Dim hmPjf As Integer    'Projection file handle
Dim imPjfRecLen As Integer        'PJF record length
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey As LONGKEY0    'Chf key record image
Dim hmCHF As Integer    'Contract file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmClf As CLF        'Clf record image
Dim hmClf As Integer    'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmCff As CFF        'Cff record image
Dim hmCff As Integer    'Contract flights file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmSlf As SLF        'Slf record image
Dim tmSlfSrchKey As INTKEY0    'Slf key record image
Dim hmSlf As Integer    'Salesperson file handle
Dim imSlfRecLen As Integer        'SLF record length
Dim tmAdf As ADF        'Adf record image
Dim tmAdfSrchKey As INTKEY0    'Adf key record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmPrf As PRF        'Prf record image
Dim tmPrfSrchKey As LONGKEY0    'Prf key record image
Dim hmPrf As Integer    'Product file handle
Dim imPrfRecLen As Integer        'PRF record length
Dim tmCxf As CXF        'Cxf record image
Dim tmCxfSrchKey As LONGKEY0    'Cxf key record image
Dim hmCxf As Integer    'Comment file handle
Dim imCxfRecLen As Integer        'CXF record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
'Dim tmRec As LPOPREC
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPropPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imSettingValue As Integer
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imIgnoreRightMove As Integer
Dim imIgnoreSetting As Integer  'Remove corporate show option
Dim imSlspSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim lmNowDate As Long
Dim imCurYear As Integer
Dim imCurMonth As Integer
Dim smMonDate As String
Dim smRolloverDate As String
Dim imState As Integer  '0=Totals; 1=Differences
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer
'Period (column) Information
Dim imPdYear As Integer
Dim imPdStartWk As Integer 'start week number
Dim imPdStartFltNo As Integer
Dim imPrjStartYear As Integer
Dim imPrjStartWk As Integer
Dim imPrjNoYears As Integer
Dim tmPdGroups(0 To 3) As PJPDGROUPS
Dim tmWKCtrls(0 To 4)  As FIELDAREA
Dim tmNWCtrls(0 To 3)  As FIELDAREA
Dim imHotSpot(0 To 4, 0 To 4) As Integer
Dim imInHotSpot As Integer
Dim imShowIndex As Integer
Dim imTypeIndex As Integer
Dim imFirstActivate As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1  'used for lower bounds of zero based arrays

Const SOFFICEINDEX = 1
Const ADVTINDEX = 2
Const PRODINDEX = 3
Const PROPNOINDEX = 4
Const DEMOINDEX = 5
Const VEHINDEX = 6
Const NRINDEX = 7
Const POTINDEX = 8
Const COMMENTINDEX = 9
Const PD1INDEX = 10
Const PD2INDEX = 11
Const PD3INDEX = 12
Const TOTALINDEX = 13
Const WK1INDEX = 1
Const WK2INDEX = 2
Const WK3INDEX = 3
Const WKTINDEX = 4
Const GTDOLLAR1INDEX = 1
Const GTDOLLAR2INDEX = 2
Const GTDOLLAR3INDEX = 3
Const GTTOTALINDEX = 4
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
            ilLoopCount = ilLoopCount + 1
            ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
            If ilRet = 0 Then
                ilIndex = cbcSelect.ListIndex   'cbcSelect.ListCount - cbcSelect.ListIndex
                slNameCode = tgSalesperson(ilIndex).sKey   'Traffic!lbcSalesperson.List(ilIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                mInitProjCtrls
                If Not mReadPjfRec(Val(slCode)) Then    ', smMonDate) Then
                    GoTo cbcSelectErr
                End If
                smRolloverDate = mGetPriorRollDate(Val(slCode), smMonDate)
                If Not mReadPPjfRec(Val(slCode), smRolloverDate) Then
                    GoTo cbcSelectErr
                End If
                imPropPopReqd = True
            Else
                'If ilRet = 1 Then
                '    cbcSelect.ListIndex = 0
                'End If
                'ilRet = 1   'Clear fields as no match name found
                pbcProj.Cls
                mSetCommands
                imChgMode = False
                Screen.MousePointer = vbDefault    'Default
                Exit Sub
            End If
            pbcProj.Cls
            'If ilRet = 0 Then
                imSlspSelectedIndex = cbcSelect.ListIndex
                mMoveRecToCtrl True
                'mInitShow
            'Else
            '    imSlspSelectedIndex = 0
            '    mClearCtrlFields
            'End If
            pbcProj_Paint
        Loop While (imSlspSelectedIndex <> cbcSelect.ListIndex) And ((imSlspSelectedIndex <> 0) Or (cbcSelect.ListIndex >= 0))
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
    imRowNo = -1
    If imFirstFocus Then
        imFirstFocus = False
        If ((tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0)) And (cbcSelect.ListCount = 1) Then
            cbcSelect.ListIndex = 0
            imSlspSelectedIndex = 0
            If imUpdateAllowed And pbcSTab.Enabled Then
                pbcSTab.SetFocus
                Exit Sub
            ElseIf Not imUpdateAllowed Then
                cmcDone.SetFocus
                Exit Sub
            End If
        End If
    End If
    imComboBoxIndex = imSlspSelectedIndex
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
Private Sub ckcDiff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcBlock_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    BlkDate.Show vbModal
End Sub
Private Sub cmcBlock_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcBlock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
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
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case PROPNOINDEX
            lbcPropNo.Visible = Not lbcPropNo.Visible
        Case SOFFICEINDEX
            lbcSOffice.Visible = Not lbcSOffice.Visible
        Case ADVTINDEX
            lbcAdvt.Visible = Not lbcAdvt.Visible
        Case PRODINDEX
            lbcProd.Visible = Not lbcProd.Visible
        Case DEMOINDEX
            lbcDemo.Visible = Not lbcDemo.Visible
        Case VEHINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case NRINDEX
            lbcNR.Visible = Not lbcNR.Visible
        Case POTINDEX
            lbcPot.Visible = Not lbcPot.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String        'General string
    Dim ilRptType As Integer
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = PROPOSALPROJECTION
    ilRptType = 0   '0=Proposal; 1=Order
    'Screen.MousePointer = vbHourglass  'Wait
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    ''Traffic!edcLinkSrceHelpMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CntrProj^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
        Else
            slStr = "CntrProj^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CntrProj^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
    '    Else
    '        slStr = "CntrProj^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'CntrProj.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'CntrProj.Enabled = True
    ''Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcRollover_Click()
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    Rollover.Show vbModal
    If imSlspSelectedIndex < 0 Then
        pbcProj.Cls
    Else
        ilRet = mReadPjfRec(tmSlf.iCode)    ', smMonDate)
        smRolloverDate = mGetPriorRollDate(tmSlf.iCode, smMonDate)
        ilRet = mReadPPjfRec(tmSlf.iCode, smRolloverDate)
        mInitProjCtrls
        pbcProj.Cls
        mMoveRecToCtrl True
        pbcProj_Paint
    End If
    mSetCommands
End Sub
Private Sub cmcRollover_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcRollover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSave_Click()
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imPjfChg = False
    mSetCommands
    pbcSTab.SetFocus
End Sub
Private Sub cmcSave_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcUndo_Click()
    Dim ilRet As Integer
    ilRet = mReadPjfRec(tmSlf.iCode)    ', smMonDate)
    mInitProjCtrls
    pbcProj.Cls
    mMoveRecToCtrl True
    pbcProj_Paint
    mSetCommands
End Sub
Private Sub cmcUndo_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer

    Select Case imBoxNo
        Case PROPNOINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcPropNo, imBSMode, imComboBoxIndex
        Case SOFFICEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSOffice, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSOffice.ListIndex = 0
            End If
        Case ADVTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAdvt, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAdvt.ListIndex = 0
            End If
        Case PRODINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcProd.ListIndex = 0
            End If
        Case DEMOINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcDemo, imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo.ListIndex = 0
            End If
        Case VEHINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case NRINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcNR, imBSMode, imComboBoxIndex
        Case POTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcPot, imBSMode, slStr)
            If ilRet = 1 Then
                lbcPot.ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case PROPNOINDEX
            imComboBoxIndex = lbcPropNo.ListIndex
        Case SOFFICEINDEX
        Case ADVTINDEX
        Case ADVTINDEX
        Case PRODINDEX
        Case DEMOINDEX
        Case VEHINDEX
            imComboBoxIndex = lbcVehicle.ListIndex
        Case NRINDEX
            imComboBoxIndex = lbcNR.ListIndex
        Case POTINDEX
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
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    'If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KEYDOWN) Then
    '    imDirProcess = KeyCode 'mDirection 0
    '    pbcTab.SetFocus
    '    Exit Sub
    'End If
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case PROPNOINDEX
                gProcessArrowKey Shift, KeyCode, lbcPropNo, imLbcArrowSetting
            Case SOFFICEINDEX
                gProcessArrowKey Shift, KeyCode, lbcSOffice, imLbcArrowSetting
            Case ADVTINDEX
                gProcessArrowKey Shift, KeyCode, lbcAdvt, imLbcArrowSetting
            Case PRODINDEX
                gProcessArrowKey Shift, KeyCode, lbcProd, imLbcArrowSetting
            Case DEMOINDEX
                gProcessArrowKey Shift, KeyCode, lbcDemo, imLbcArrowSetting
            Case VEHINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case NRINDEX
                gProcessArrowKey Shift, KeyCode, lbcNR, imLbcArrowSetting
            Case POTINDEX
                gProcessArrowKey Shift, KeyCode, lbcPot, imLbcArrowSetting
        End Select
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case SOFFICEINDEX, ADVTINDEX, PRODINDEX, DEMOINDEX, NRINDEX, POTINDEX
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
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(PROPOSALSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
        pbcProj.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    Else
        If imSlspSelectedIndex < 0 Then
            pbcProj.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cmcSave.Enabled = False
            cmcUndo.Enabled = False
        Else
            pbcProj.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
        End If
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    '12/12/14: Disallow salesperson from rollover operation.
    mSetCommands
    Me.KeyPreview = True
    CntrProj.Refresh
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
    'If Screen.Width * 15 = 640 Then
    '    fmAdjFactorW = 1#
    '    fmAdjFactorH = 1#
    'Else
    '    fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
    '    Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    '    fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
    '    Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    'End If
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
    
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    Erase tmAdvertiser
    Erase tmSOfficeCode
    Erase tmSaleOffice
    Erase tmUserVeh
    Erase tgPjf1Rec
    Erase tgPjf2Rec
    Erase tgPjfDel
    Erase tgOPjf1Rec
    Erase tgOPjf2Rec
    Erase tgPPjf1Rec
    Erase tgPPjf2Rec
    Erase tgClfCntrProj
    Erase tgCffCntrProj
    Erase smShow
    Erase smSave
    Erase lmSave
    Erase imSave
'    btrExtClear hmDsf   'Clear any previous extend operation
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    btrExtClear hmCff   'Clear any previous extend operation
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    btrExtClear hmPjf   'Clear any previous extend operation
    ilRet = btrClose(hmPjf)
    btrDestroy hmPjf
    btrExtClear hmSlf   'Clear any previous extend operation
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmPrf   'Clear any previous extend operation
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    btrExtClear hmCxf   'Clear any previous extend operation
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf

    Set CntrProj = Nothing
    End
End Sub
Private Sub imcHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub imcKey_Click()
    pbcKey.ZOrder vbBringToFront
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilPjf As Integer
    If (imRowNo < vbcProj.Value) Or (imRowNo > vbcProj.Value + vbcProj.LargeChange) Then
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    ilPjf = imSave(1, ilRowNo)
    If ilUpperBound <= LBONE Then
        mClearCtrlFields
        mSetCommands
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        Exit Sub
    End If
    If ilPjf > 0 Then
        If tgPjf1Rec(ilPjf).iStatus = 1 Then
            tgPjfDel(UBound(tgPjfDel)).tPjf = tgPjf1Rec(ilPjf).tPjf
            tgPjfDel(UBound(tgPjfDel)).iStatus = tgPjf1Rec(ilPjf).iStatus
            tgPjfDel(UBound(tgPjfDel)).lRecPos = tgPjf1Rec(ilPjf).lRecPos
            ReDim Preserve tgPjfDel(0 To UBound(tgPjfDel) + 1) As PJF2REC
            ilPjf = tgPjf1Rec(ilPjf).i2RecIndex
            If ilPjf > 0 Then
                If tgPjf2Rec(ilPjf).iStatus = 1 Then
                    tgPjfDel(UBound(tgPjfDel)).tPjf = tgPjf2Rec(ilPjf).tPjf
                    tgPjfDel(UBound(tgPjfDel)).iStatus = tgPjf2Rec(ilPjf).iStatus
                    tgPjfDel(UBound(tgPjfDel)).lRecPos = tgPjf2Rec(ilPjf).lRecPos
                    ReDim Preserve tgPjfDel(0 To UBound(tgPjfDel) + 1) As PJF2REC
                End If
            End If
        End If
        ilPjf = imSave(1, ilRowNo)
        If ilRowNo <= UBound(tgPjf1Rec) - 1 Then
            'Remove record from tgRjf1Rec- Leave tgPjf2Rec
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                tgPjf1Rec(ilLoop) = tgPjf1Rec(ilLoop + 1)
                tgPjf1Rec(ilLoop).iSaveIndex = tgPjf1Rec(ilLoop).iSaveIndex - 1
            Next ilLoop
            ReDim Preserve tgPjf1Rec(0 To UBound(tgPjf1Rec) - 1) As PJF1REC
        End If
    End If
    For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
        For ilIndex = 1 To UBound(smSave, 1) Step 1
            smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
        Next ilIndex
        For ilIndex = 1 To UBound(imSave, 1) Step 1
            If ilIndex = 1 Then
                imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1) - 1
            Else
                imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1)
            End If
        Next ilIndex
        For ilIndex = 1 To UBound(lmSave, 1) Step 1
            lmSave(ilIndex, ilLoop) = lmSave(ilIndex, ilLoop + 1)
        Next ilIndex
        For ilIndex = 1 To UBound(smShow, 1) Step 1
            smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
        Next ilIndex
    Next ilLoop
    ilUpperBound = UBound(smSave, 2)
    If ilUpperBound > 1 Then
        ReDim Preserve smShow(0 To TOTALINDEX, 0 To ilUpperBound - 1) As String 'Values shown in program area
        ReDim Preserve smSave(0 To 8, 0 To ilUpperBound - 1) As String 'Values saved (program name) in program area
        ReDim Preserve lmSave(0 To 5, 0 To ilUpperBound - 1) As Long 'Values saved (program name) in program area
        ReDim Preserve imSave(0 To 3, 0 To ilUpperBound - 1) As Integer 'Values saved (program name) in program area
    End If
    imSettingValue = True
    If UBound(smSave, 2) <= vbcProj.LargeChange Then 'was <=
        vbcProj.Max = LBONE 'LBound(smSave, 2) '- 1
    Else
        vbcProj.Max = UBound(smSave, 2) - vbcProj.LargeChange '- 1
    End If
    imSettingValue = False
    imPjfChg = True
    mSetCommands
    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcProj.Cls
    mGetShowPrices
    pbcProj_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
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
Private Sub lbcAdvt_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcAdvt_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcAdvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcAdvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcDemo_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcDemo, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcDemo_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcDemo, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcNR_Click()
    gProcessLbcClick lbcNR, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcNR_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPot_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcPot, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcPot_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcPot_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcPot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcPot, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcProd_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
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
Private Sub lbcPropNo_Click()
    gProcessLbcClick lbcPropNo, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcPropNo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSOffice_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSOffice_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSOffice_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSOffice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSOffice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddProd                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Product record             *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Function mAddProd(ilAdfCode As Integer, slProduct As String) As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    tmAdfSrchKey.iCode = ilAdfCode
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mAddProd = 0
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    imPrfRecLen = Len(tmPrf)
    Do  'Loop until record updated or added
        tmPrf.lCode = 0
        tmPrf.iAdfCode = tmAdf.iCode
        tmPrf.sName = slProduct
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
        tmPrf.lAutoCode = tmPrf.lCode
        ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
        sgProdCodeTag = ""
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAddProd = 0
    Else
        mAddProd = tmPrf.lCode
        'If tgSpf.sRemoteUsers = "Y" Then
            Do
                'tmPrfSrchKey.lCode = tmPrf.lCode
                'ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                gPackTime slSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        'End If
    End If
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
    ilRet = gOptionalLookAhead(edcDropDown, lbcAdvt, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        imDoubleClickName = False
        mAdvtBranch = False
        Exit Function
    End If
    If igWinStatus(ADVERTISERSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mAdvtBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(ADVERTISERSLIST)) Then
    '    imDoubleClickName = False
    '    mAdvtBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    igAdvtCallSource = CALLSOURCECONTRACT
    If edcDropDown.Text = "[New]" Then
        sgAdvtName = ""
    Else
        sgAdvtName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CntrProj^Test\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
        Else
            slStr = "CntrProj^Prod\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CntrProj^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    '    Else
    '        slStr = "CntrProj^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Advt.Exe " & slStr, 1)
    'CntrProj.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Advt.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtName)
    igAdvtCallSource = Val(sgAdvtName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtName)
    'CntrProj.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAdvtBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    If igAdvtCallSource = CALLDONE Then  'Done
        igAdvtCallSource = CALLNONE
'        gSetMenuState True
        'slFilter = cbcFilter.Text
        'ilFilter = imFilterSelectedIndex
'        slStamp = FileDateTime(sgDBPath & "Adf.Btr")
'        If StrComp(slStamp, Traffic!lbcAdvertiser.Tag, 1) <> 0 Then
        lbcAdvt.Clear
        smAdvertiserTag = ""
        mAdvtPop
        If imTerminate Then
            mAdvtBranch = False
            Exit Function
        End If
'        End If
        gFindMatch sgAdvtName, 1, lbcAdvt
        If gLastFound(lbcAdvt) > 0 Then
            imChgMode = True
            lbcAdvt.ListIndex = gLastFound(lbcAdvt)
            edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            imChgMode = False
            mAdvtBranch = False
            mSetChg ADVTINDEX
        Else
            imChgMode = True
            lbcAdvt.ListIndex = 1
            edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            imChgMode = False
            mSetChg ADVTINDEX
            edcDropDown.SetFocus
            sgAdvtName = ""
            Exit Function
        End If
        sgAdvtName = ""
    End If
    If igAdvtCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAdvtCallSource = CALLNONE
        sgAdvtName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtCallSource = CALLTERMINATED Then
'        gSetMenuState True
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
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcAdvt.ListIndex
    If ilIndex > 0 Then
        slName = lbcAdvt.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(CntrProj, lbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(CntrProj, lbcAdvt, tmAdvertiser(), smAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", CntrProj
        On Error GoTo 0
        lbcAdvt.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcAdvt
            If gLastFound(lbcAdvt) > 0 Then
                lbcAdvt.ListIndex = gLastFound(lbcAdvt)
            Else
                lbcAdvt.ListIndex = -1
            End If
        Else
            lbcAdvt.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAanyZeros                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test for zeros in records      *
'*                                                     *
'*******************************************************
Private Function mAnyZeros() As Integer
    Dim ilPjf As Integer
    Dim ilPjf2 As Integer
    Dim ilRowNo As Integer
    For ilPjf = LBONE To UBound(tgPjf1Rec) - 1 Step 1
        ilRowNo = tgPjf1Rec(ilPjf).iSaveIndex
        If (tgPjf1Rec(ilPjf).tPjf.iSofCode = 0) Or (tgPjf1Rec(ilPjf).tPjf.iAdfCode = 0) Then
            imRowNo = ilRowNo
            mAnyZeros = SOFFICEINDEX
            Exit Function
        End If
        ilPjf2 = tgPjf1Rec(ilPjf).i2RecIndex
        If ilPjf2 > 0 Then
            If (tgPjf2Rec(ilPjf2).tPjf.iSofCode = 0) Or (tgPjf2Rec(ilPjf2).tPjf.iAdfCode = 0) Then
                imRowNo = ilRowNo
                mAnyZeros = ADVTINDEX
                Exit Function
            End If
        End If
    Next ilPjf
    mAnyZeros = 0
End Function
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
    imPjfChg = False
    lbcPropNo.ListIndex = -1
    lbcSOffice.ListIndex = -1
    lbcAdvt.ListIndex = -1
    lbcProd.ListIndex = -1
    lbcDemo.ListIndex = -1
    lbcVehicle.ListIndex = -1
    lbcNR.ListIndex = -1
    lbcPot.ListIndex = -1
    edcComment.Text = ""
    ReDim tgPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPjf2Rec(0 To 1) As PJF2REC
    ReDim tgPjfDel(0 To 1) As PJF2REC
    ReDim smShow(0 To TOTALINDEX, 0 To 1) As String 'Values shown in program area
    ReDim smSave(0 To 8, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmSave(0 To 5, 0 To 1) As Long 'Values saved (program name) in program area
    ReDim imSave(0 To 3, 0 To 1) As Integer 'Values saved (program name) in program area
    mInitProjCtrls
    'edcName.Text = ""
    'lbcGenre.ListIndex = -1
    'edcComment.Text = ""
    'lbcTime.ListIndex = -1
    'lbcLen.ListIndex = -1
    'lbcProg.ListIndex = -1
    'edcSource.Text = ""
    'edcType(0).Text = ""
    'edcType(1).Text = ""
    'tmEnf.iVefCode = 0  'Force this field to be reset in mMoveCtrlToRec
    'tmEnf.iEtfCode = 0  'Force this field to be reset in mMoveCtrlToRec
    'smComment = ""
    'smGenre = ""
    'mMoveCtrlToRec False
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    '    tmCtrls(ilLoop).iChg = False
    'Next ilLoop
    'imTimeFirst = True
    'imLenFirst = True
    'imProgFirst = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompMonths                     *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Months for Year        *
'*                                                     *
'*******************************************************
Private Sub mCompMonths(ilYear As Integer, ilStartWk() As Integer, ilNoWks() As Integer)
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLoop As Integer
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartCorp(slDate, True)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndCorp(slStart, True)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1    'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    Else                        'Standard
        'Compute start week number for each month
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartStd(slDate)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndStd(slStart)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1      'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Potential Code        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
'
'   mDemoPop
'   Where:
'
    Dim ilRet As Integer
    Dim slDemo As String      'Demo name, saved to determine if changed
    Dim ilIndex As Integer      'Demo name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilIndex = lbcDemo.ListIndex
    If ilIndex > 1 Then
        slDemo = lbcDemo.List(ilIndex)
    End If
    'ilRet = gPopMnfPlusFieldsBox(CntrProj, lbcDemo, Traffic!lbcDemoCode, "D")
    ilRet = gPopMnfPlusFieldsBox(CntrProj, lbcDemo, tgDemoCode(), sgDemoCodeTag, "D")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox)", CntrProj
        On Error GoTo 0
        lbcDemo.AddItem "[None]", 0
        imChgMode = True
        If ilIndex >= 1 Then
            gFindMatch slDemo, 1, lbcDemo
            If gLastFound(lbcDemo) > 0 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
            Else
                lbcDemo.ListIndex = -1
            End If
        Else
            lbcDemo.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim ilLoop As Integer   'For loop control parameter
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilUsePrev As Integer
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcProj.Top + tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True

    Select Case ilBoxNo 'Branch on box type (control)
        Case SOFFICEINDEX
            'mSaleOfficePop
            Screen.MousePointer = vbHourglass  'Wait
            mSaleOfficePop
            Screen.MousePointer = vbDefault
            If imTerminate Then
                Exit Sub
            End If
            lbcSOffice.Height = gListBoxHeight(lbcSOffice.ListCount, 6)
            edcDropDown.Width = 3 * tmCtrls(ilBoxNo).fBoxW \ 2 '- cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcSOffice.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcSOffice.Move edcDropDown.Left, edcDropDown.Top - lbcSOffice.Height
            End If
            imChgMode = True
            gFindMatch smSave(2, imRowNo), 1, lbcSOffice
            If gLastFound(lbcSOffice) > 0 Then
                lbcSOffice.ListIndex = gLastFound(lbcSOffice)
            Else
                lbcSOffice.ListIndex = -1
                If imRowNo = LBONE Then
                    For ilLoop = LBONE To UBound(tmSaleOffice) - 1 Step 1
                        If tmSaleOffice(ilLoop).iCode = tmSlf.iSofCode Then
                            gFindMatch tmSaleOffice(ilLoop).sName, 1, lbcSOffice
                            If gLastFound(lbcSOffice) > 0 Then
                                lbcSOffice.ListIndex = gLastFound(lbcSOffice)
                            End If
                            Exit For
                        End If
                    Next ilLoop
                Else
                    gFindMatch smSave(2, imRowNo - 1), 1, lbcSOffice
                    If gLastFound(lbcSOffice) > 0 Then
                        lbcSOffice.ListIndex = gLastFound(lbcSOffice)
                    Else
                        For ilLoop = LBONE To UBound(tmSaleOffice) - 1 Step 1
                            If tmSaleOffice(ilLoop).iCode = tmSlf.iSofCode Then
                                gFindMatch tmSaleOffice(ilLoop).sName, 1, lbcSOffice
                                If gLastFound(lbcSOffice) > 0 Then
                                    lbcSOffice.ListIndex = gLastFound(lbcSOffice)
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
            End If
            If lbcSOffice.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcSOffice.List(lbcSOffice.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ADVTINDEX
            'mAdvtPop
            Screen.MousePointer = vbHourglass  'Wait
            mAdvtPop
            Screen.MousePointer = vbDefault
            If imTerminate Then
                Exit Sub
            End If
            lbcAdvt.Height = gListBoxHeight(lbcAdvt.ListCount, 6)
            edcDropDown.Width = 3 * tmCtrls(ADVTINDEX).fBoxW \ 2 '- cmcDropDown.Width
            edcDropDown.MaxLength = 50
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcAdvt.Height
            End If
            imChgMode = True
            gFindMatch smSave(3, imRowNo), 1, lbcAdvt
            If gLastFound(lbcAdvt) > 0 Then
                lbcAdvt.ListIndex = gLastFound(lbcAdvt)
            Else
                If smSave(3, imRowNo) <> "" Then
                    lbcAdvt.ListIndex = -1
                    edcDropDown.Text = smSave(3, imRowNo)
                Else
                    lbcAdvt.ListIndex = -1
                    mGetChf lmSave(5, imRowNo)
                    If (smSave(1, imRowNo) <> "") And (tmChf.iAdfCode > 0) Then
                        For ilLoop = 0 To UBound(tmAdvertiser) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
                            slNameCode = tmAdvertiser(ilLoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmChf.iAdfCode Then
                                lbcAdvt.ListIndex = ilLoop + 1
                                edcDropDown.Text = lbcAdvt.List(ilLoop + 1)
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        If imRowNo > LBONE Then
                            gFindMatch smSave(3, imRowNo - 1), 1, lbcAdvt
                            If gLastFound(lbcAdvt) > 0 Then
                                lbcAdvt.ListIndex = gLastFound(lbcAdvt)
                            End If
                        End If
                    End If
                End If
            End If
            If lbcAdvt.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODINDEX
            gFindMatch smSave(3, imRowNo), 1, lbcAdvt
            If gLastFound(lbcAdvt) > 0 Then
                slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) <> tmAdf.iCode Then
                    tmAdfSrchKey.iCode = Val(slCode) 'ilCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmAdf.iCode = 0
                    End If
                End If
            Else
                tmAdf.iCode = 0
            End If
            'mProdPop tmAdf.iCode
            Screen.MousePointer = vbHourglass  'Wait
            mProdPop tmAdf.iCode
            Screen.MousePointer = vbDefault
            If imTerminate Then
                Exit Sub
            End If
            lbcProd.Height = gListBoxHeight(lbcProd.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 35  'tgSpf.iAProd
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcProd.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcProd.Move edcDropDown.Left, edcDropDown.Top - lbcProd.Height
            End If
            imChgMode = True
            gFindMatch smSave(4, imRowNo), 1, lbcProd
            If gLastFound(lbcProd) >= 1 Then
                lbcProd.ListIndex = gLastFound(lbcProd)
                edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
            Else
                If smSave(4, imRowNo) <> "" Then
                    lbcProd.ListIndex = -1
                    edcDropDown.Text = smSave(4, imRowNo)
                Else
                    ilUsePrev = False
                    If (imRowNo > LBONE) Then
                        If smSave(3, imRowNo) = smSave(3, imRowNo - 1) Then
                            ilUsePrev = True
                        End If
                    End If
                    If ilUsePrev Then
                        gFindMatch smSave(4, imRowNo - 1), 1, lbcProd
                        If gLastFound(lbcProd) >= 1 Then
                            lbcProd.ListIndex = gLastFound(lbcProd)
                            edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                        Else
                            'lbcProd.ListIndex = 0
                            lbcProd.ListIndex = -1
                            edcDropDown.Text = smSave(4, imRowNo - 1)
                        End If
                    Else
                        mGetChf lmSave(5, imRowNo)
                        If (smSave(1, imRowNo) <> "") And (Trim$(tmChf.sProduct) <> "") Then
                            gFindMatch tmChf.sProduct, 1, lbcProd
                            If gLastFound(lbcProd) >= 1 Then
                                lbcProd.ListIndex = gLastFound(lbcProd)
                            Else
                                gFindMatch tmAdf.sProduct, 1, lbcProd
                                If gLastFound(lbcProd) >= 1 Then
                                    lbcProd.ListIndex = gLastFound(lbcProd)
                                Else
                                    lbcProd.ListIndex = 0
                                End If
                            End If
                        Else
                            gFindMatch tmAdf.sProduct, 1, lbcProd
                            If gLastFound(lbcProd) >= 1 Then
                                lbcProd.ListIndex = gLastFound(lbcProd)
                            Else
                                lbcProd.ListIndex = 0
                            End If
                        End If
                        edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                    End If
                End If
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PROPNOINDEX
            'mPropNoPop imRowNo
            Screen.MousePointer = vbHourglass  'Wait
            mPropNoPop imRowNo
            Screen.MousePointer = vbDefault
            If imTerminate Then
                Exit Sub
            End If
            lbcPropNo.Height = gListBoxHeight(lbcPropNo.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 0
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcPropNo.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcPropNo.Move edcDropDown.Left, edcDropDown.Top - lbcPropNo.Height
            End If
            imChgMode = True
            gFindMatch smSave(1, imRowNo), 1, lbcPropNo
            If gLastFound(lbcPropNo) > 0 Then
                lbcPropNo.ListIndex = gLastFound(lbcPropNo)
            Else
                lbcPropNo.ListIndex = 0   '[None]
            End If
            imComboBoxIndex = lbcPropNo.ListIndex
            If lbcPropNo.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcPropNo.List(lbcPropNo.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DEMOINDEX
            lbcDemo.Height = gListBoxHeight(lbcDemo.ListCount, 6)
            edcDropDown.Width = 3 * tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcDemo.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDemo.Move edcDropDown.Left, edcDropDown.Top - lbcProd.Height
            End If
            imChgMode = True
            gFindMatch smSave(5, imRowNo), 1, lbcDemo
            If gLastFound(lbcDemo) >= 1 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
                edcDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
            Else
                If smSave(5, imRowNo) <> "" Then
                    lbcDemo.ListIndex = -1
                    edcDropDown.Text = smSave(5, imRowNo)
                Else
                    lbcDemo.ListIndex = -1
                    edcDropDown.Text = ""
                    If imRowNo > 1 Then
                        If (smSave(3, imRowNo) = smSave(3, imRowNo - 1)) And (smSave(4, imRowNo) = smSave(4, imRowNo - 1)) Then
                            gFindMatch smSave(5, imRowNo - 1), 1, lbcDemo
                            If gLastFound(lbcDemo) >= 1 Then
                                lbcDemo.ListIndex = gLastFound(lbcDemo)
                                edcDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
                            End If
                        End If
                    End If
                    If lbcDemo.ListIndex < 0 Then
                        mGetChf lmSave(5, imRowNo)
                        If (smSave(1, imRowNo) <> "") And (tmChf.iMnfDemo(0) > 0) Then
                            For ilLoop = 0 To UBound(tgDemoCode) - 1 Step 1 'Traffic!lbcDemoCode.ListCount - 1 Step 1
                                slNameCode = tgDemoCode(ilLoop).sKey   'Traffic!lbcDemoCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmChf.iMnfDemo(0) Then
                                    lbcDemo.ListIndex = ilLoop + 1
                                    edcDropDown.Text = lbcDemo.List(ilLoop + 1)
                                    Exit For
                                End If
                            Next ilLoop
                        Else
                            For ilLoop = 0 To UBound(tgDemoCode) - 1 Step 1 'Traffic!lbcDemoCode.ListCount - 1 Step 1
                                slNameCode = tgDemoCode(ilLoop).sKey   'Traffic!lbcDemoCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmAdf.iMnfDemo(0) Then
                                    lbcDemo.ListIndex = ilLoop + 1
                                    edcDropDown.Text = lbcDemo.List(ilLoop + 1)
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                    If lbcDemo.ListIndex < 0 Then
                        lbcDemo.ListIndex = 0
                        edcDropDown.Text = lbcDemo.List(0)
                    End If
                End If
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case VEHINDEX
            'mVehPop
            Screen.MousePointer = vbHourglass  'Wait
            mVehPop
            Screen.MousePointer = vbDefault
            If imTerminate Then
                Exit Sub
            End If
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
            edcDropDown.Width = tmCtrls(VEHINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            imChgMode = True
            gFindMatch smSave(6, imRowNo), 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                imComboBoxIndex = lbcVehicle.ListIndex
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            Else
                If imRowNo > 1 Then
                    gFindMatch smSave(6, imRowNo - 1), 0, lbcVehicle
                    If (gLastFound(lbcVehicle) >= 0) And (smSave(3, imRowNo) = smSave(3, imRowNo - 1)) Then
                        If gLastFound(lbcVehicle) < lbcVehicle.ListCount - 1 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle) + 1
                            imComboBoxIndex = lbcVehicle.ListIndex
                            edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                        Else
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            imComboBoxIndex = lbcVehicle.ListIndex
                            edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                        End If
                    Else
                        slName = sgUserDefVehicleName
                        gFindMatch slName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            imComboBoxIndex = lbcVehicle.ListIndex
                            edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                        Else
                            lbcVehicle.ListIndex = 0
                            imComboBoxIndex = lbcVehicle.ListIndex
                            edcDropDown.Text = lbcVehicle.List(0)
                        End If
                    End If
                Else
                    slName = sgUserDefVehicleName
                    gFindMatch slName, 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                        imComboBoxIndex = lbcVehicle.ListIndex
                        edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                    Else
                        lbcVehicle.ListIndex = 0
                        imComboBoxIndex = lbcVehicle.ListIndex
                        edcDropDown.Text = lbcVehicle.List(0)
                    End If
                End If
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case NRINDEX
            lbcNR.Height = gListBoxHeight(lbcNR.ListCount, 2)
            edcDropDown.Width = 3 * tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 7
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcNR.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If imSave(2, imRowNo) < 0 Then
                If imRowNo = LBONE Then
                    lbcNR.ListIndex = 1
                Else
                    lbcNR.ListIndex = imSave(2, imRowNo - 1)
                End If
            Else
                lbcNR.ListIndex = imSave(2, imRowNo)
            End If
            imComboBoxIndex = lbcNR.ListIndex
            If lbcNR.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcNR.List(lbcNR.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case POTINDEX
            lbcPot.Height = gListBoxHeight(lbcPot.ListCount, 6)
            edcDropDown.Width = 2 * tmCtrls(POTINDEX).fBoxW '- cmcDropDwon.Width
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcProj.Value <= vbcProj.LargeChange \ 2 Then
                lbcPot.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcPot.Move edcDropDown.Left, edcDropDown.Top - lbcPot.Height
            End If
            imChgMode = True
            gFindMatch smSave(7, imRowNo), 1, lbcPot
            If gLastFound(lbcPot) >= 1 Then
                lbcPot.ListIndex = gLastFound(lbcPot)
                edcDropDown.Text = lbcPot.List(lbcPot.ListIndex)
            Else
                If smSave(7, imRowNo) <> "" Then
                    lbcPot.ListIndex = -1
                    edcDropDown.Text = smSave(7, imRowNo)
                Else
                    lbcPot.ListIndex = lbcPot.ListCount - 1 '0    '[New]
                    edcDropDown.Text = lbcPot.List(lbcPot.ListIndex)
                    mGetChf lmSave(5, imRowNo)
                    If (smSave(1, imRowNo) <> "") And (tmChf.iMnfPotnType > 0) Then
                        For ilLoop = 0 To UBound(tgPotCode) - 1 Step 1  'Traffic!lbcPotCode.ListCount - 1 Step 1
                            slNameCode = tgPotCode(ilLoop).sKey    'Traffic!lbcPotCode.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmChf.iMnfPotnType Then
                                lbcPot.ListIndex = ilLoop + 1
                                edcDropDown.Text = lbcPot.List(ilLoop + 1)
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        If imRowNo > LBONE Then
                            gFindMatch smSave(7, imRowNo - 1), 1, lbcPot
                            If gLastFound(lbcPot) >= 1 Then
                                lbcPot.ListIndex = gLastFound(lbcPot)
                                edcDropDown.Text = lbcPot.List(lbcPot.ListIndex)
                            End If
                        End If
                    End If
                End If
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMMENTINDEX
            'edcComment.Width = tmCtrls(ilBoxNo).fBoxW
            edcComment.MaxLength = 1000
            edcComment.Move pbcProj.Left + tmCtrls(ilBoxNo).fBoxX, pbcProj.Top + tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            edcComment.Text = smSave(8, imRowNo)
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
        Case PD1INDEX, PD2INDEX, PD3INDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcProj, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15)
            If (imRowNo > LBONE) And (lmSave(ilBoxNo - PD1INDEX + 1, imRowNo) = 0) Then
                edcDropDown.Text = "0"  'Trim$(Str$(lmSave(ilBoxNo - PD1INDEX + 1, imRowNo - 1)))
            Else
                edcDropDown.Text = Trim$(Str$(lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)))
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenProjForPropNo               *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate Projection for a      *
'*                      specified proposal             *
'*                                                     *
'*******************************************************
Private Sub mGenProjForPropNo(llChfCode As Long, ilAsk As Integer)
    Dim ilRes As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilPjf As Integer
    Dim ilPjf2 As Integer
    Dim ilRet As Integer
    Dim ilStartPjf As Integer
    Dim ilUpper As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilCff As Integer
    Dim slDate As String
    Dim slStart As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llPrice As Long
    Dim ilNoSpots As Integer
    Dim llYearEndDate As Long
    Dim llYear1StartDate As Long
    Dim llYear2StartDate As Long
    Dim llMonDate As Long
    Dim ilWkNo As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilGross As Integer
    Dim slAdvtName As String

    If llChfCode <= 0 Then
        Exit Sub
    End If
    If ilAsk Then
        ilRes = MsgBox("Copy Totals from Proposal to Projections", vbYesNo + vbQuestion, "Update")
    Else
        ilRes = vbYes
    End If
    If ilRes = vbYes Then
        'Set values into tgPjf1Rec and tgPjf2Rec so save can be rebuild
        ilRet = mMoveCtrlToRec()
        slDate = "1/15/" & Trim$(Str$(imCurYear))
        llYear1StartDate = gDateValue(gObtainStartStd(slDate))
        slDate = "1/15/" & Trim$(Str$(imCurYear + 1))
        llYear2StartDate = gDateValue(gObtainStartStd(slDate))
        slDate = "12/15/" & Trim$(Str$(imCurYear))
        llYearEndDate = gDateValue(gObtainEndStd(slDate))
        llMonDate = gDateValue(smMonDate)
        'Remove any projections that match advertiser/Product from projection
        'and generate a new projection for each vehicle defined in the proposal
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfCntrProj, tgClfCntrProj(), tgCffCntrProj())
        If Not ilRet Then
            Exit Sub
        End If
        tmAdfSrchKey.iCode = tgChfCntrProj.iAdfCode 'ilCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmAdf.sName = "~~~~~~|||||||||"
            slAdvtName = Trim$(tmAdf.sName)
        Else
            If (tmAdf.sBillAgyDir = "D") Then
                If (Trim$(tmAdf.sAddrID) <> "") Then
                    slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
                Else
                    slAdvtName = Trim$(tmAdf.sName) & "/Direct"
                End If
            Else
                slAdvtName = Trim$(tmAdf.sName)
            End If
        End If
        ilIndex = LBONE 'LBound(smSave, 2)
        ilPjf = LBONE   'LBound(tgPjf1Rec)
        Do While ilPjf <= UBound(tgPjf1Rec)
            ilIndex = tgPjf1Rec(ilPjf).iSaveIndex
            If ilIndex > 0 Then
                If (smSave(3, ilIndex) = Trim$(slAdvtName)) And (smSave(4, ilIndex) = Trim$(tmChf.sProduct)) Then
                    If tgPjf1Rec(ilPjf).iStatus = 1 Then
                        tgPjfDel(UBound(tgPjfDel)).tPjf = tgPjf1Rec(ilPjf).tPjf
                        tgPjfDel(UBound(tgPjfDel)).iStatus = tgPjf1Rec(ilPjf).iStatus
                        tgPjfDel(UBound(tgPjfDel)).lRecPos = tgPjf1Rec(ilPjf).lRecPos
                        ReDim Preserve tgPjfDel(0 To UBound(tgPjfDel) + 1) As PJF2REC
                        ilPjf2 = tgPjf1Rec(ilPjf).i2RecIndex
                        If ilPjf2 > 0 Then
                            If tgPjf2Rec(ilPjf2).iStatus = 1 Then
                                tgPjfDel(UBound(tgPjfDel)).tPjf = tgPjf2Rec(ilPjf2).tPjf
                                tgPjfDel(UBound(tgPjfDel)).iStatus = tgPjf2Rec(ilPjf2).iStatus
                                tgPjfDel(UBound(tgPjfDel)).lRecPos = tgPjf2Rec(ilPjf2).lRecPos
                                ReDim Preserve tgPjfDel(0 To UBound(tgPjfDel) + 1) As PJF2REC
                            End If
                        End If
                    End If
                    If ilPjf <= UBound(tgPjf1Rec) - 1 Then
                        'Remove record from tgRjf1Rec- Leave tgPjf2Rec
                        For ilLoop = ilPjf To UBound(tgPjf1Rec) - 1 Step 1
                            tgPjf1Rec(ilLoop) = tgPjf1Rec(ilLoop + 1)
                        Next ilLoop
                        ReDim Preserve tgPjf1Rec(0 To UBound(tgPjf1Rec) - 1) As PJF1REC
                    Else
                        Exit Do
                    End If
                Else
                    ilPjf = ilPjf + 1
                End If
            Else
                ilPjf = ilPjf + 1
            End If
        Loop
        ilStartPjf = UBound(tgPjf1Rec)
        'Create new record for each vehicle
        For ilLoop = LBound(tgClfCntrProj) To UBound(tgClfCntrProj) - 1 Step 1
            If (tgClfCntrProj(ilLoop).ClfRec.sType <> "O") And (tgClfCntrProj(ilLoop).ClfRec.sType <> "A") Then
                'Determine if record created for proposal and matching vehicle- if so update
                ilIndex = -1
                For ilPjf = ilStartPjf To UBound(tgPjf1Rec) - 1 Step 1
                    If (tgPjf1Rec(ilPjf).tPjf.lChfCode = llChfCode) And ((tgPjf1Rec(ilPjf).tPjf.iVefCode = tgClfCntrProj(ilLoop).ClfRec.iVefCode)) Then
                        ilIndex = ilPjf
                        Exit For
                    End If
                Next ilPjf
                If ilIndex = -1 Then
                    ilUpper = UBound(tgPjf1Rec)
                    tgPjf1Rec(ilUpper).tPjf.iSlfCode = tmSlf.iCode
                    tgPjf1Rec(ilUpper).tPjf.iSofCode = tmSlf.iSofCode
                    tgPjf1Rec(ilUpper).tPjf.iYear = imCurYear
                    tgPjf1Rec(ilUpper).tPjf.iMnfBus = tgChfCntrProj.iMnfPotnType
                    tgPjf1Rec(ilUpper).tPjf.iEffDate(0) = 0
                    tgPjf1Rec(ilUpper).tPjf.iEffDate(1) = 0
                    tgPjf1Rec(ilUpper).tPjf.lCxfChgR = 0
                    tgPjf1Rec(ilUpper).tPjf.iAdfCode = tgChfCntrProj.iAdfCode
                    mProdPop tgChfCntrProj.iAdfCode
                    gFindMatch Trim$(tgChfCntrProj.sProduct), 1, lbcProd
                    If gLastFound(lbcProd) >= 1 Then
                        slNameCode = tgProdCode(gLastFound(lbcProd) - 1).sKey    'Traffic!lbcProdCode.List(gLastFound(lbcProd) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilRet = CP_MSG_NONE Then
                            tgPjf1Rec(ilUpper).tPjf.lPrfCode = Val(slCode)
                        Else
                            tgPjf1Rec(ilUpper).tPjf.lPrfCode = 0
                        End If
                    Else
                        tgPjf1Rec(ilUpper).tPjf.lPrfCode = 0
                    End If
                    tgPjf1Rec(ilUpper).tPjf.iMnfDemo = tgChfCntrProj.iMnfDemo(0)
                    tgPjf1Rec(ilUpper).tPjf.iVefCode = tgClfCntrProj(ilLoop).ClfRec.iVefCode
                    tgPjf1Rec(ilUpper).tPjf.sNewRet = "R"
                    tgPjf1Rec(ilUpper).tPjf.lChfCode = llChfCode
                    tgPjf1Rec(ilUpper).tPjf.iRolloverDate(0) = 0
                    tgPjf1Rec(ilUpper).tPjf.iRolloverDate(1) = 0
                    tgPjf1Rec(ilUpper).tPjf.iEffTime(0) = 1
                    tgPjf1Rec(ilUpper).tPjf.iEffTime(1) = 0
                    tgPjf1Rec(ilUpper).tPjf.iUrfCode = tgUrf(0).iCode
                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                        tgPjf1Rec(ilUpper).sKey = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & tgChfCntrProj.sProduct
                        tgPjf1Rec(ilUpper).sAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                    Else
                        tgPjf1Rec(ilUpper).sKey = tmAdf.sName & tgChfCntrProj.sProduct
                        tgPjf1Rec(ilUpper).sAdvtName = tmAdf.sName
                    End If
                    tgPjf1Rec(ilUpper).sProdName = tgChfCntrProj.sProduct
                    For ilGross = LBound(tgPjf1Rec(ilUpper).tPjf.lGross) To UBound(tgPjf1Rec(ilUpper).tPjf.lGross) Step 1
                        tgPjf1Rec(ilUpper).tPjf.lGross(ilGross) = 0
                    Next ilGross
                    tgPjf1Rec(ilUpper).iStatus = 0
                    tgPjf1Rec(ilUpper).lRecPos = 0
                    If imPrjNoYears > 1 Then
                        tgPjf1Rec(ilUpper).i2RecIndex = UBound(tgPjf2Rec)
                        tgPjf2Rec(UBound(tgPjf2Rec)).tPjf = tgPjf1Rec(ilUpper).tPjf
                        tgPjf2Rec(UBound(tgPjf2Rec)).tPjf.iYear = imCurYear + 1
                        tgPjf2Rec(UBound(tgPjf2Rec)).iStatus = 0
                        tgPjf2Rec(UBound(tgPjf2Rec)).lRecPos = 0
                        ReDim Preserve tgPjf2Rec(0 To UBound(tgPjf2Rec) + 1) As PJF2REC
                    Else
                        tgPjf1Rec(ilUpper).i2RecIndex = 0
                    End If
                    ilIndex = ilUpper
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgPjf1Rec(0 To ilUpper) As PJF1REC
                End If
                'Increment gross values from flight
                ilCff = tgClfCntrProj(ilLoop).iFirstCff
                Do While ilCff >= LBound(tgCffCntrProj)
                    gUnpackDateLong tgCffCntrProj(ilCff).CffRec.iStartDate(0), tgCffCntrProj(ilCff).CffRec.iStartDate(1), llStartDate    'Week Start date
                    gUnpackDateLong tgCffCntrProj(ilCff).CffRec.iEndDate(0), tgCffCntrProj(ilCff).CffRec.iEndDate(1), llEndDate    'Week Start date
                    If llEndDate < llStartDate Then
                        Exit Do
                    End If
                    If tgCffCntrProj(ilCff).CffRec.sPriceType = "T" Then
                        llDate = llStartDate
                        Do While llDate <= llEndDate
                            If llDate >= llMonDate Then
                                If tgCffCntrProj(ilCff).CffRec.sDyWk = "D" Then
                                    ilNoSpots = 0
                                    For ilDay = gWeekDayLong(llDate) To 6 Step 1
                                        If llDate + ilDay <= llEndDate Then
                                            ilNoSpots = ilNoSpots + tgCffCntrProj(ilCff).CffRec.iDay(ilDay)
                                        End If
                                    Next ilDay
                                Else
                                    ilNoSpots = tgCffCntrProj(ilCff).CffRec.iSpotsWk + tgCffCntrProj(ilCff).CffRec.iXSpotsWk
                                End If
                                llPrice = (ilNoSpots * tgCffCntrProj(ilCff).CffRec.lActPrice + 50) \ 100
                                'Compute week index
                                If llDate <= llYearEndDate Then
                                    ilWkNo = (llDate - llYear1StartDate) / 7 + 1
                                    If ilWkNo = 1 Then
                                        slDate = "1/15/" & Trim$(Str$(imCurYear))
                                        slStart = gObtainStartCorp(slDate, True)
                                        ilDay = gWeekDayStr(slStart)
                                        If ilDay = 0 Then
                                            tgPjf1Rec(ilIndex).tPjf.lGross(0) = tgPjf1Rec(ilIndex).tPjf.lGross(0) + 0
                                            tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) = tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) + llPrice
                                        Else
                                            tgPjf1Rec(ilIndex).tPjf.lGross(0) = tgPjf1Rec(ilIndex).tPjf.lGross(0) + (llPrice * ilDay) / 7
                                            tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) = tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) + (llPrice - (llPrice * ilDay) / 7)
                                        End If
                                    'ElseIf ilWkNo = 52 Then
                                    '    slDate = "12/15/" & Trim$(Str$(imCurYear))
                                    '    slStart = gObtainEndCorp(slDate, True)
                                    '    ilDay = gWeekDayStr(slStart)
                                    '    If ilDay = 6 Then
                                    '        tgPjf1Rec(ilIndex).tPjf.lGross(53) = tgPjf1Rec(ilIndex).tPjf.lGross(53) + 0
                                    '        tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) = tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) + llPrice
                                    '    Else
                                    '        ilDay = 7 - ilDay - 1
                                    '        tgPjf1Rec(ilIndex).tPjf.lGross(52) = tgPjf1Rec(ilIndex).tPjf.lGross(52) + (llPrice * ilDay) / 7
                                    '        tgPjf1Rec(ilIndex).tPjf.lGross(53) = tgPjf1Rec(ilIndex).tPjf.lGross(53) + (llPrice - (llPrice * ilDay) / 7)
                                    '    End If
                                    Else
                                        tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) = tgPjf1Rec(ilIndex).tPjf.lGross(ilWkNo) + llPrice
                                    End If

                                Else
                                    ilPjf2 = tgPjf1Rec(ilIndex).i2RecIndex
                                    ilWkNo = (llDate - llYear2StartDate) / 7 + 1
                                    If ilWkNo <= 53 Then
                                        'ilWkNo = (llDate - llYear1StartDate) / 7 + 1
                                        If ilWkNo = 1 Then
                                            slDate = "1/15/" & Trim$(Str$(imCurYear))
                                            slStart = gObtainStartCorp(slDate, True)
                                            ilDay = gWeekDayStr(slStart)
                                            If ilDay = 0 Then
                                                tgPjf2Rec(ilPjf2).tPjf.lGross(0) = tgPjf2Rec(ilPjf2).tPjf.lGross(0) + 0
                                                tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) = tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) + llPrice
                                            Else
                                                tgPjf2Rec(ilPjf2).tPjf.lGross(0) = tgPjf2Rec(ilPjf2).tPjf.lGross(0) + (llPrice * ilDay) / 7
                                                tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) = tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) + (llPrice - (llPrice * ilDay) / 7)
                                            End If
                                        'ElseIf ilWkNo = 52 Then
                                        '    slDate = "12/15/" & Trim$(Str$(imCurYear))
                                        '    slStart = gObtainEndCorp(slDate, True)
                                        '    ilDay = gWeekDayStr(slStart)
                                        '    If ilDay = 6 Then
                                        '        tgPjf2Rec(ilPjf2).tPjf.lGross(53) = tgPjf2Rec(ilPjf2).tPjf.lGross(53) + 0
                                        '        tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) = tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) + llPrice
                                        '    Else
                                        '        ilDay = 7 - ilDay - 1
                                        '        tgPjf2Rec(ilPjf2).tPjf.lGross(52) = tgPjf2Rec(ilPjf2).tPjf.lGross(52) + (llPrice * ilDay) / 7
                                        '        tgPjf2Rec(ilPjf2).tPjf.lGross(53) = tgPjf2Rec(ilPjf2).tPjf.lGross(53) + (llPrice - (llPrice * ilDay) / 7)
                                        '    End If
                                        Else
                                            tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) = tgPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) + llPrice
                                        End If
                                    End If
                                End If
                            End If
                            'Advance to next monday
                            Do
                                llDate = llDate + 1
                            Loop Until gWeekDayLong(llDate) = 0
                        Loop
                    End If
                    ilCff = tgCffCntrProj(ilCff).iNextCff
                Loop
            End If
        Next ilLoop
        'ReMake Sort key for all record
        For ilLoop = LBONE To UBound(tgPjf1Rec) - 1 Step 1
            If tgPjf1Rec(ilLoop).tPjf.iAdfCode <> tmAdf.iCode Then
                tmAdfSrchKey.iCode = tgPjf1Rec(ilLoop).tPjf.iAdfCode 'ilCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            If tgPjf1Rec(ilLoop).tPjf.lPrfCode > 0 Then
                If tgPjf1Rec(ilLoop).tPjf.lPrfCode <> tmPrf.lCode Then
                    tmPrfSrchKey.lCode = tgPjf1Rec(ilLoop).tPjf.lPrfCode 'ilCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Else
                tmPrf.sName = ""
                tmPrf.lCode = 0
            End If
            'tgPjf1Rec(ilLoop).sKey = tmAdf.sName & tmPrf.sName
            'tgPjf1Rec(ilLoop).sAdvtName = tmAdf.sName
            If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                tgPjf1Rec(ilLoop).sKey = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & tmPrf.sName
                tgPjf1Rec(ilLoop).sAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
            Else
                tgPjf1Rec(ilLoop).sKey = tmAdf.sName & tmPrf.sName
                tgPjf1Rec(ilLoop).sAdvtName = tmAdf.sName
            End If
            tgPjf1Rec(ilLoop).sProdName = tmPrf.sName
        Next ilLoop
        If ilUpper > 1 Then
            ArraySortTyp fnAV(tgPjf1Rec(), 1), UBound(tgPjf1Rec) - 1, 0, LenB(tgPjf1Rec(1)), 0, LenB(tgPjf1Rec(1).sKey), 0
        End If
        pbcProj.Cls
        vbcProj.Value = vbcProj.Min
        mMoveRecToCtrl False
        'Find one added and set row number to it
        For ilPjf = LBONE To UBound(tgPjf1Rec) - 1 Step 1
            If tgPjf1Rec(ilPjf).tPjf.lChfCode = llChfCode Then
                imRowNo = tgPjf1Rec(ilPjf).iSaveIndex
                Exit For
            End If
        Next ilPjf
        mGetShowPrices
        mGetShowOPrices
        mGetShowPPrices
        pbcProj_Paint
        imPjfChg = True
        imRowNo = -1
        imBoxNo = -1
        pbcClickFocus.SetFocus
        mSetCommands
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetChf                         *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get contract so product is     *
'*                      obtained                       *
'*                                                     *
'*******************************************************
Private Sub mGetChf(llChfCode As Long)
    Dim ilRet As Integer

    If llChfCode = 0 Then
        tmChf.lCode = 0
        tmChf.iAdfCode = 0
        tmChf.sProduct = ""
        tmChf.iMnfDemo(0) = 0
        tmChf.iMnfPotnType = 0
        Exit Sub
    End If
    If llChfCode <> tmChf.lCode Then
        tmChfSrchKey.lCode = llChfCode 'ilCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmChf.iAdfCode = 0
            tmChf.sProduct = ""
            tmChf.iMnfDemo(0) = 0
            tmChf.iMnfPotnType = 0
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetChfCode                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get contract from Proposal #   *
'*                                                     *
'*******************************************************
Private Function mGetChfCode(slPropNo As String) As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    mGetChfCode = 0
    If Trim$(slPropNo) = "" Then
        Exit Function
    End If
    gFindMatch slPropNo, 1, lbcPropNo
    If gLastFound(lbcPropNo) > 0 Then
        slNameCode = tgCntrCode(gLastFound(lbcPropNo) - 1).sKey  'Traffic!lbcCntrCode.List(gLastFound(lbcPropNo) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmChfSrchKey.lCode = Val(slCode)
        If tmChf.lCode <> tmChfSrchKey.lCode Then
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                mGetChfCode = tmChf.lCode
            End If
        Else
            mGetChfCode = tmChf.lCode
        End If
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetPriorRollDate               *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine rollover date        *
'*                                                     *
'*******************************************************
Private Function mGetPriorRollDate(ilSlfCode As Integer, slRollDate As String) As String
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llRDate As Long
    Dim slDate As String

    mGetPriorRollDate = ""
    tmSlfSrchKey.iCode = ilSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If
    llDate = 0
    tmPjfSrchKey.iSlfCode = ilSlfCode
    '   M     S    M     S    M     S
    '                         |Today
    '  |-------|  |-------|  |-------|
    'scan extra week (it shouldn't be required if rollover every week)
    '
    slDate = gDecOneWeek(gDecOneWeek(gObtainPrevMonday(slRollDate)))
    gPackDate slDate, tmPjfSrchKey.iRolloverDate(0), tmPjfSrchKey.iRolloverDate(1)
    ilRet = btrGetGreaterOrEqual(hmPjf, tmPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) And (tmPjf.iSlfCode = ilSlfCode) Then
        Do
            If (tmPjf.iRolloverDate(0) <> 0) Or (tmPjf.iRolloverDate(1) <> 0) Then
                gUnpackDateLong tmPjf.iRolloverDate(0), tmPjf.iRolloverDate(1), llRDate
                If llRDate > llDate Then
                    llDate = llRDate
                End If
            End If
            ilRet = btrGetNext(hmPjf, tmPjf, imPjfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop While (ilRet = BTRV_ERR_NONE) And (tmPjf.iSlfCode = ilSlfCode)
        If llDate > 0 Then
            mGetPriorRollDate = Format$(llDate, "m/d/yy")
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowDates                   *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowDates(ilCallGetPrice As Integer)
'
'   mGetShowDates
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim slWkEnd As String
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilYearOk As Integer
    Dim ilWkNo As Integer
    Dim ilWkCount As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    Dim slFontName As String
    Dim flFontSize As Single
    'If UBound(tgPjfRec) <= 1 Then
    '    For ilIndex = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
    '        tmPdGroups(ilIndex).iStartWkNo = -1
    '        tmPdGroups(ilIndex).iNoWks = 0
    '        tmPdGroups(ilIndex).iTrueNoWks = 0
    '        tmPdGroups(ilIndex).iFltNo = 0
    '        tmPdGroups(ilIndex).sStartDate = ""
    '        tmPdGroups(ilIndex).sEndDate = ""
    '        gSetShow pbcProj, "", tmWKCtrls(ilIndex)
    '    Next ilIndex
    '    Exit Sub
    'End If
    If imPdYear = 0 Then
        Exit Sub
    End If
    slFontName = pbcProj.FontName
    flFontSize = pbcProj.FontSize
    pbcProj.FontBold = False
    pbcProj.FontSize = 7
    pbcProj.FontName = "Arial"
    pbcProj.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    tmPdGroups(1).iYear = imPdYear
    tmPdGroups(1).iStartWkNo = imPdStartWk

    ilIndex = 1
    Do
        ilFound = False
        If ilIndex > 1 Then
            If tmPdGroups(ilIndex).iYear <> tmPdGroups(ilIndex - 1).iYear Then
                ilYearOk = False
            Else
                ilYearOk = True
            End If
        Else
            ilYearOk = False
        End If
        If Not ilYearOk Then
            mCompMonths tmPdGroups(ilIndex).iYear, ilStartWk(), ilNoWks()
        End If
        If rbcType(0).Value Then        'Quarter
            For ilLoop = LBONE To 12 Step 3
                If (tmPdGroups(ilIndex).iStartWkNo >= ilStartWk(ilLoop)) And (tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(ilLoop) + ilNoWks(ilLoop) + ilNoWks(ilLoop + 1) + ilNoWks(ilLoop + 2) - 1) Then
                    tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop) + ilNoWks(ilLoop + 1) + ilNoWks(ilLoop + 2) - (tmPdGroups(ilIndex).iStartWkNo - ilStartWk(ilLoop))
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf rbcType(1).Value Then    'Month
            For ilLoop = LBONE To 12 Step 1
                If (tmPdGroups(ilIndex).iStartWkNo >= ilStartWk(ilLoop)) And (tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(ilLoop) + ilNoWks(ilLoop) - 1) Then
                    tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop) - (tmPdGroups(ilIndex).iStartWkNo - ilStartWk(ilLoop))
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf rbcType(2).Value Then    'Week
            If tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(12) + ilNoWks(12) - 1 Then
                tmPdGroups(ilIndex).iNoWks = 1
                ilFound = True
            End If
        End If
        If ilFound Then
            If ilIndex <> 3 Then    'was 4
                tmPdGroups(ilIndex + 1).iStartWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks
                tmPdGroups(ilIndex + 1).iYear = tmPdGroups(ilIndex).iYear   'imPdYear
            End If
            ilIndex = ilIndex + 1
        Else
            tmPdGroups(ilIndex).iYear = tmPdGroups(ilIndex).iYear + 1
            tmPdGroups(ilIndex).iStartWkNo = 1
            'Test if year exist
            If tmPdGroups(ilIndex).iYear > imPrjStartYear + imPrjNoYears - 1 Then
                For ilLoop = ilIndex To 3 Step 1    'was 4
                    tmPdGroups(ilLoop).iStartWkNo = -1
                    tmPdGroups(ilLoop).iTrueNoWks = 0
                    tmPdGroups(ilLoop).iNoWks = 0
                Next ilLoop
                Exit Do
            End If
        End If
    Loop Until ilIndex > 3  'was 4
    'Compute Start/End Date if groups
    For ilIndex = LBONE To UBound(tmPdGroups) Step 1
        If tmPdGroups(ilIndex).iStartWkNo > 0 Then
            If rbcShow(0).Value Then    'Corporate
                slDate = "1/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartCorp(slDate, True)
                'slDate = "12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                'slEnd = gObtainEndCorp(slDate, True)
                slDate = slStart
                For ilLoop = 1 To 12 Step 1
                    slEnd = gObtainEndCorp(slDate, True)
                    slDate = gIncOneDay(slEnd)
                Next ilLoop
            Else
                slDate = "1/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartStd(slDate)
                slDate = "12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slEnd = gObtainEndStd(slDate)
            End If
            ilWkNo = 1
            Do
                If ilWkNo = tmPdGroups(ilIndex).iStartWkNo Then
                    tmPdGroups(ilIndex).sStartDate = slStart
                    slWkEnd = gObtainNextSunday(slStart)
                    ilWkCount = 1
                    Do
                        If ilWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks - 1 Then
                            tmPdGroups(ilIndex).sEndDate = slWkEnd
                            tmPdGroups(ilIndex).iTrueNoWks = ilWkCount
                            'slDate = tmPdGroups(ilIndex).sStartDate & "-" & tmPdGroups(ilIndex).sEndDate
                            slDate = tmPdGroups(ilIndex).sStartDate
                            slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                            gSetShow pbcProj, slDate, tmWKCtrls(ilIndex)
                            slStr = Trim$(Str$(ilWkCount))
                            slStr = "# Weeks " & slStr
                            gSetShow pbcProj, slStr, tmNWCtrls(ilIndex)
                            Exit Do
                        Else
                            ilWkNo = ilWkNo + 1
                            ilWkCount = ilWkCount + 1
                            slWkEnd = gIncOneWeek(slWkEnd)
                            If gDateValue(slWkEnd) > gDateValue(slEnd) Then
                                tmPdGroups(ilIndex).sEndDate = slEnd
                                tmPdGroups(ilIndex).iTrueNoWks = ilWkCount - 1
                                'slDate = Left$(tmPdGroups(ilIndex).sStartDate, Len(tmPdGroups(ilIndex).sStartDate) - 3)
                                'slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                slDate = tmPdGroups(ilIndex).sStartDate
                                slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                gSetShow pbcProj, slDate, tmWKCtrls(ilIndex)
                                slStr = Trim$(Str$(ilWkCount - 1))
                                slStr = "# Weeks " & slStr
                                gSetShow pbcProj, slStr, tmNWCtrls(ilIndex)
                                Exit Do
                            End If
                        End If
                    Loop
                    Exit Do
                Else
                    ilWkNo = ilWkNo + 1
                    slStart = gIncOneWeek(slStart)
                End If
            Loop
        Else
            tmPdGroups(ilIndex).sStartDate = ""
            tmPdGroups(ilIndex).sEndDate = ""
            gSetShow pbcProj, "", tmWKCtrls(ilIndex)
            gSetShow pbcProj, "", tmNWCtrls(ilIndex)
        End If
    Next ilIndex
    pbcProj.FontSize = flFontSize
    pbcProj.FontName = slFontName
    pbcProj.FontSize = flFontSize
    pbcProj.FontBold = True
    'mGetShowPrices
    'mGetShowPPrices
    'mGetShowOPrices
    If ilCallGetPrice Then
        mGetShowPrices
        mGetShowPPrices
        mGetShowOPrices
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowOPrices                 *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate prices from original *
'*                                                     *
'*******************************************************
Private Sub mGetShowOPrices()
    Dim ilPjf As Integer
    Dim il2Pjf As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim slStart As String
    Dim ilBox As Integer
    Dim llGross As Long
    For ilBox = LBONE To UBound(lmOAPSave) Step 1
        lmOAPSave(ilBox) = 0
        lmOTSave(ilBox) = 0
    Next ilBox
    'Sum value
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        For ilPjf = LBONE To UBound(tgOPjf1Rec) Step 1 '- 1 Step 1
            If tgOPjf1Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    ''Add in the first part of the standard week
                    'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                    '    lmOTSave(ilGroup) = lmOTSave(ilGroup) + tgOPjf1Rec(ilPjf).tPjf.lGross(0)
                    '    If imRowNo > 0 Then
                    '        If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    '            lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + tgOPjf1Rec(ilPjf).tPjf.lGross(0)
                    '        End If
                    '    End If
                    'End If
                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                        If rbcShow(0).Value Then
                            slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                            mPjfGetGross ilPjf, slStart, tgOPjf1Rec(), tgOPjf2Rec(), llGross
                        Else
                            llGross = tgOPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                        End If
                        lmOTSave(ilGroup) = lmOTSave(ilGroup) + llGross 'tgOPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                        If imRowNo > 0 Then
                            If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                                lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + llGross 'tgOPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                            End If
                        End If
                    Next ilWk
                    'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                    '    lmOTSave(ilGroup) = lmOTSave(ilGroup) + tgOPjf1Rec(ilPjf).tPjf.lGross(53)
                    '    If imRowNo > 0 Then
                    '        If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    '            lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + tgOPjf1Rec(ilPjf).tPjf.lGross(53)
                    '        End If
                    '    End If
                    'End If
                End If
            Else
                il2Pjf = tgOPjf1Rec(ilPjf).i2RecIndex
                If il2Pjf > 0 Then
                    If tgOPjf2Rec(il2Pjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                        If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                            ''Add in the first part of the standard week
                            'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                            '    lmOTSave(ilGroup) = lmOTSave(ilGroup) + tgOPjf2Rec(il2Pjf).tPjf.lGross(0)
                            '    If imRowNo > 0 Then
                            '        If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                            '            lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + tgOPjf2Rec(il2Pjf).tPjf.lGross(0)
                            '        End If
                            '    End If
                            'End If
                            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                If rbcShow(0).Value Then
                                    slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                                    mPjfGetGross ilPjf, slStart, tgOPjf1Rec(), tgOPjf2Rec(), llGross
                                Else
                                    llGross = tgOPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                End If
                                lmOTSave(ilGroup) = lmOTSave(ilGroup) + llGross 'tgOPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                If imRowNo > 0 Then
                                    If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                                        lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + llGross 'tgOPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                    End If
                                End If
                            Next ilWk
                            'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                            '    lmOTSave(ilGroup) = lmOTSave(ilGroup) + tgOPjf2Rec(il2Pjf).tPjf.lGross(53)
                            '    If imRowNo > 0 Then
                            '        If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                            '            lmOAPSave(ilGroup) = lmOAPSave(ilGroup) + tgOPjf2Rec(il2Pjf).tPjf.lGross(53)
                            '        End If
                            '    End If
                            'End If
                        End If
                    End If
                End If
            End If
        Next ilPjf
    Next ilGroup
    'Compute totals
    For ilPjf = LBONE To UBound(tgOPjf1Rec) Step 1 '- 1 Step 1
        If (rbcShow(1).Value) Then
            lmOTSave(GTTOTALINDEX) = lmOTSave(GTTOTALINDEX) + tgOPjf1Rec(ilPjf).tPjf.lGross(0)
            If imRowNo > 0 Then
                If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    lmOAPSave(GTTOTALINDEX) = lmOAPSave(GTTOTALINDEX) + tgOPjf1Rec(ilPjf).tPjf.lGross(0)
                End If
            End If
        End If
        For ilWk = 1 To 53 Step 1
            lmOTSave(GTTOTALINDEX) = lmOTSave(GTTOTALINDEX) + tgOPjf1Rec(ilPjf).tPjf.lGross(ilWk)
            If imRowNo > 0 Then
                If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    lmOAPSave(GTTOTALINDEX) = lmOAPSave(GTTOTALINDEX) + tgOPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                End If
            End If
        Next ilWk
        If (imPrjNoYears > 1) And (tgOPjf1Rec(ilPjf).i2RecIndex > 0) Then
            If (rbcShow(1).Value) Then
                lmOTSave(GTTOTALINDEX) = lmOTSave(GTTOTALINDEX) + tgOPjf2Rec(tgOPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                If imRowNo > 0 Then
                    If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                        lmOAPSave(GTTOTALINDEX) = lmOAPSave(GTTOTALINDEX) + tgOPjf2Rec(tgOPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                    End If
                End If
            End If
            For ilWk = 1 To 53 Step 1
                lmOTSave(GTTOTALINDEX) = lmOTSave(GTTOTALINDEX) + tgOPjf2Rec(tgOPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                If imRowNo > 0 Then
                    If (StrComp(Trim$(tgOPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgOPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                        lmOAPSave(GTTOTALINDEX) = lmOAPSave(GTTOTALINDEX) + tgOPjf2Rec(tgOPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                    End If
                End If
            Next ilWk
        End If
    Next ilPjf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowPPrices                 *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate prices for prior week*
'*                                                     *
'*******************************************************
Private Sub mGetShowPPrices()
    Dim ilPjf As Integer
    Dim il2Pjf As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim slStart As String
    Dim ilBox As Integer
    Dim llGross As Long
    For ilBox = LBONE To UBound(tmPAPCtrls) Step 1
        tmPAPCtrls(ilBox).sShow = ""
    Next ilBox
    For ilBox = LBONE To UBound(tmPTCtrls) Step 1
        tmPTCtrls(ilBox).sShow = ""
    Next ilBox
    For ilBox = LBONE To UBound(lmPAPSave) Step 1
        lmPAPSave(ilBox) = 0
        lmPTSave(ilBox) = 0
    Next ilBox
    'Sum value
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        For ilPjf = LBONE To UBound(tgPPjf1Rec) Step 1 '- 1 Step 1
            If tgPPjf1Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    ''Add in the first part of the standard week
                    'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                    '    lmPTSave(ilGroup) = lmPTSave(ilGroup) + tgPPjf1Rec(ilPjf).tPjf.lGross(0)
                    '    If imRowNo > 0 Then
                    '        If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    '            lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + tgPPjf1Rec(ilPjf).tPjf.lGross(0)
                    '        End If
                    '    End If
                    'End If
                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                        If rbcShow(0).Value Then
                            slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                            mPjfGetGross ilPjf, slStart, tgPPjf1Rec(), tgPPjf2Rec(), llGross
                        Else
                            llGross = tgPPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                        End If
                        lmPTSave(ilGroup) = lmPTSave(ilGroup) + llGross 'tgPPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                        If imRowNo > 0 Then
                            If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                                lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + llGross       'tgPPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                            End If
                        End If
                    Next ilWk
                    'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                    '    lmPTSave(ilGroup) = lmPTSave(ilGroup) + tgPPjf1Rec(ilPjf).tPjf.lGross(53)
                    '    If imRowNo > 0 Then
                    '        If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    '            lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + tgPPjf1Rec(ilPjf).tPjf.lGross(53)
                    '        End If
                    '    End If
                    'End If
                End If
            Else
                il2Pjf = tgPPjf1Rec(ilPjf).i2RecIndex
                If il2Pjf > 0 Then
                    If tgPPjf2Rec(il2Pjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                        If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                            ''Add in the first part of the standard week
                            'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                            '    lmPTSave(ilGroup) = lmPTSave(ilGroup) + tgPPjf2Rec(il2Pjf).tPjf.lGross(0)
                            '    If imRowNo > 0 Then
                            '        If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                            '            lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + tgPPjf2Rec(il2Pjf).tPjf.lGross(0)
                            '        End If
                            '    End If
                            'End If
                            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                If rbcShow(0).Value Then
                                    slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                                    mPjfGetGross ilPjf, slStart, tgPPjf1Rec(), tgPPjf2Rec(), llGross
                                Else
                                    llGross = tgPPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                End If
                                lmPTSave(ilGroup) = lmPTSave(ilGroup) + llGross 'tgPPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                If imRowNo > 0 Then
                                    If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                                        lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + llGross 'tgPPjf2Rec(il2Pjf).tPjf.lGross(ilWk)
                                    End If
                                End If
                            Next ilWk
                            'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                            '    lmPTSave(ilGroup) = lmPTSave(ilGroup) + tgPPjf2Rec(il2Pjf).tPjf.lGross(53)
                            '    If imRowNo > 0 Then
                            '        If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                            '            lmPAPSave(ilGroup) = lmPAPSave(ilGroup) + tgPPjf2Rec(il2Pjf).tPjf.lGross(53)
                            '        End If
                            '    End If
                            'End If
                        End If
                    End If
                End If
            End If
        Next ilPjf
    Next ilGroup
    'Compute totals
    For ilPjf = LBONE To UBound(tgPPjf1Rec) Step 1 '- 1 Step 1
        If (rbcShow(1).Value) Then
            lmPTSave(GTTOTALINDEX) = lmPTSave(GTTOTALINDEX) + tgPPjf1Rec(ilPjf).tPjf.lGross(0)
            If imRowNo > 0 Then
                If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    lmPAPSave(GTTOTALINDEX) = lmPAPSave(GTTOTALINDEX) + tgPPjf1Rec(ilPjf).tPjf.lGross(0)
                End If
            End If
        End If
        For ilWk = 1 To 53 Step 1
            lmPTSave(GTTOTALINDEX) = lmPTSave(GTTOTALINDEX) + tgPPjf1Rec(ilPjf).tPjf.lGross(ilWk)
            If imRowNo > 0 Then
                If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                    lmPAPSave(GTTOTALINDEX) = lmPAPSave(GTTOTALINDEX) + tgPPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                End If
            End If
        Next ilWk
        If (imPrjNoYears > 1) And (tgPPjf1Rec(ilPjf).i2RecIndex > 0) Then
            If (rbcShow(1).Value) Then
                lmPTSave(GTTOTALINDEX) = lmPTSave(GTTOTALINDEX) + tgPPjf2Rec(tgPPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                If imRowNo > 0 Then
                    If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                        lmPAPSave(GTTOTALINDEX) = lmPAPSave(GTTOTALINDEX) + tgPPjf2Rec(tgPPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                    End If
                End If
            End If
            For ilWk = 1 To 53 Step 1
                lmPTSave(GTTOTALINDEX) = lmPTSave(GTTOTALINDEX) + tgPPjf2Rec(tgPPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                If imRowNo > 0 Then
                    If (StrComp(Trim$(tgPPjf1Rec(ilPjf).sAdvtName), Trim$(smSave(3, imRowNo)), 1) = 0) And (StrComp(Trim$(tgPPjf1Rec(ilPjf).sProdName), Trim$(smSave(4, imRowNo)), 1) = 0) Then
                        lmPAPSave(GTTOTALINDEX) = lmPAPSave(GTTOTALINDEX) + tgPPjf2Rec(tgPPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                    End If
                End If
            Next ilWk
        End If
    Next ilPjf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowPrices                  *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowPrices()
    Dim ilLoop As Integer
    Dim ilPjf As Integer
    Dim ilPjf1 As Integer
    Dim ilIndex As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim slStart As String
    Dim slStr As String
    Dim ilBox As Integer
    Dim llGross As Long
    For ilLoop = LBONE To UBound(smShow, 2) Step 1  '- 1 Step 1
        smShow(PD1INDEX, ilLoop) = ""
        smShow(PD2INDEX, ilLoop) = ""
        smShow(PD3INDEX, ilLoop) = ""
        smShow(TOTALINDEX, ilLoop) = ""    'Total for year
        lmSave(PD1INDEX - PD1INDEX + 1, ilLoop) = 0
        lmSave(PD2INDEX - PD1INDEX + 1, ilLoop) = 0
        lmSave(PD3INDEX - PD1INDEX + 1, ilLoop) = 0
        lmSave(TOTALINDEX - PD1INDEX + 1, ilLoop) = 0
    Next ilLoop
    For ilBox = LBONE To UBound(tmCAPCtrls) Step 1
        tmCAPCtrls(ilBox).sShow = ""
    Next ilBox
    For ilBox = LBONE To UBound(tmCTCtrls) Step 1
        tmCTCtrls(ilBox).sShow = ""
    Next ilBox
    For ilBox = LBONE To UBound(lmCAPSave) Step 1
        lmCAPSave(ilBox) = 0
        lmCTSave(ilBox) = 0
    Next ilBox
    If (UBound(smSave, 2) = LBONE) And (smSave(2, UBound(smSave, 2)) = "") Then
        Exit Sub
    End If
    'Sum value
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        For ilIndex = LBONE To UBound(smSave, 2) Step 1 '- 1 Step 1
            ilPjf = imSave(1, ilIndex)
            ilPjf1 = ilPjf
            If ilPjf > 0 Then
                If tgPjf1Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                    If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                        ''Add in the first part of the standard week
                        'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                        '    lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                        '    lmCTSave(ilGroup) = lmCTSave(ilGroup) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                        '    If imRowNo > 0 Then
                        '        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                        '            lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                        '        End If
                        '    End If
                        'End If
                        For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                            If rbcShow(0).Value Then
                                slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                                mPjfGetGross ilPjf1, slStart, tgPjf1Rec(), tgPjf2Rec(), llGross
                            Else
                                llGross = tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                            End If
                            lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + llGross 'tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                            lmCTSave(ilGroup) = lmCTSave(ilGroup) + llGross 'tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                            If imRowNo > 0 Then
                                If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                                    lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + llGross 'tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                                End If
                            End If
                        Next ilWk
                        'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                        '    lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + tgPjf1Rec(ilPjf).tPjf.lGross(53)
                        '    lmCTSave(ilGroup) = lmCTSave(ilGroup) + tgPjf1Rec(ilPjf).tPjf.lGross(53)
                        '    If imRowNo > 0 Then
                        '        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                        '            lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + tgPjf1Rec(ilPjf).tPjf.lGross(53)
                        '        End If
                        '    End If
                        'End If
                    End If
                Else
                    ilPjf = tgPjf1Rec(ilPjf).i2RecIndex
                    If ilPjf > 0 Then
                        If tgPjf2Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                            If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                                ''Add in the first part of the standard week
                                'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                                '    lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + tgPjf2Rec(ilPjf).tPjf.lGross(0)
                                '    lmCTSave(ilGroup) = lmCTSave(ilGroup) + tgPjf2Rec(ilPjf).tPjf.lGross(0)
                                '    If imRowNo > 0 Then
                                '        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                                '            lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + tgPjf2Rec(ilPjf).tPjf.lGross(0)
                                '        End If
                                '    End If
                                'End If
                                For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                    If rbcShow(0).Value Then
                                        slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                                        mPjfGetGross ilPjf1, slStart, tgPjf1Rec(), tgPjf2Rec(), llGross
                                    Else
                                        llGross = tgPjf2Rec(ilPjf).tPjf.lGross(ilWk)
                                    End If
                                    lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + llGross 'tgPjf2Rec(ilPjf).tPjf.lGross(ilWk)
                                    lmCTSave(ilGroup) = lmCTSave(ilGroup) + llGross 'tgPjf2Rec(ilPjf).tPjf.lGross(ilWk)
                                    If imRowNo > 0 Then
                                        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                                            lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + llGross 'tgPjf2Rec(ilPjf).tPjf.lGross(ilWk)
                                        End If
                                    End If
                                Next ilWk
                                'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                                '    lmSave(ilGroup, ilIndex) = lmSave(ilGroup, ilIndex) + tgPjf2Rec(ilPjf).tPjf.lGross(53)
                                '    lmCTSave(ilGroup) = lmCTSave(ilGroup) + tgPjf2Rec(ilPjf).tPjf.lGross(53)
                                '    If imRowNo > 0 Then
                                '        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                                '            lmCAPSave(ilGroup) = lmCAPSave(ilGroup) + tgPjf2Rec(ilPjf).tPjf.lGross(53)
                                '        End If
                                '    End If
                                'End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilIndex
    Next ilGroup
    'Compute totals
    For ilIndex = LBONE To UBound(smSave, 2) Step 1 '- 1 Step 1
        If smSave(2, ilIndex) <> "" Then
            ilPjf = imSave(1, ilIndex)
        Else
            ilPjf = -1
        End If
        If ilPjf > 0 Then
            If (rbcShow(1).Value) Then
                lmSave(UBound(tmPdGroups) + 1, ilIndex) = lmSave(UBound(tmPdGroups) + 1, ilIndex) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                If imRowNo > 0 Then
                    If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                        lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) + tgPjf1Rec(ilPjf).tPjf.lGross(0)
                    End If
                End If
            End If
            For ilWk = 1 To 53 Step 1
                lmSave(UBound(tmPdGroups) + 1, ilIndex) = lmSave(UBound(tmPdGroups) + 1, ilIndex) + tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) + tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                If imRowNo > 0 Then
                    If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                        lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) + tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
                    End If
                End If
            Next ilWk
            If imPrjNoYears > 1 Then
                If (rbcShow(1).Value) Then
                    lmSave(UBound(tmPdGroups) + 1, ilIndex) = lmSave(UBound(tmPdGroups) + 1, ilIndex) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                    lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                    If imRowNo > 0 Then
                        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                            lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(0)
                        End If
                    End If
                End If
                For ilWk = 1 To 53 Step 1
                    lmSave(UBound(tmPdGroups) + 1, ilIndex) = lmSave(UBound(tmPdGroups) + 1, ilIndex) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                    lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                    If imRowNo > 0 Then
                        If (smSave(3, ilIndex) = smSave(3, imRowNo)) And (smSave(4, ilIndex) = smSave(4, imRowNo)) Then
                            lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) + tgPjf2Rec(tgPjf1Rec(ilPjf).i2RecIndex).tPjf.lGross(ilWk)
                        End If
                    End If
                Next ilWk
            End If
        End If
    Next ilIndex
    'For ilWk = 1 To 53 Step 1
    '    lmOSave(GTTOTALINDEX) = lmOSave(GTTOTALINDEX) + tmT1Pjf.lGross(ilWk)
    'Next ilWk
    'If imPrjNoYears > 1 Then
    '    For ilWk = 1 To 53 Step 1
    '        lmOSave(GTTOTALINDEX) = lmOSave(GTTOTALINDEX) + tmT2Pjf.lGross(ilWk)
    '    Next ilWk
    'End If
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        For ilIndex = LBONE To UBound(lmSave, 2) Step 1 '- 1 Step 1
            If imSave(1, ilIndex) > 0 Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    slStr = Trim$(Str$(lmSave(ilGroup, ilIndex)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcProj, slStr, tmCtrls(PD1INDEX)
                    smShow(ilGroup + PD1INDEX - 1, ilIndex) = tmCtrls(PD1INDEX).sShow
                    If ilGroup = LBONE Then
                        slStr = Trim$(Str$(lmSave(UBound(tmPdGroups) + 1, ilIndex)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcProj, slStr, tmCtrls(TOTALINDEX)
                        smShow(TOTALINDEX, ilIndex) = tmCtrls(TOTALINDEX).sShow
                    End If
                Else
                    smShow(ilGroup + PD1INDEX - 1, ilIndex) = ""
                    If ilGroup = LBONE Then
                        smShow(TOTALINDEX, ilIndex) = ""
                    End If
                End If
            Else
                smShow(ilGroup + PD1INDEX - 1, ilIndex) = ""
                If ilGroup = LBONE Then
                    smShow(TOTALINDEX, ilIndex) = ""
                End If
            End If
        Next ilIndex
    Next ilGroup
    'For ilBox = LBound(tmGTCtrls) To UBound(tmGTCtrls) Step 1
    '    slStr = Trim$(Str$(lmGTSave(ilBox)))
    '    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    '    gSetShow pbcProj, slStr, tmGTCtrls(ilBox)
    '    smGTShow(ilBox) = tmGTCtrls(ilBox).sShow
    '    slStr = Trim$(Str$(lmOSave(ilBox)))
    '    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    '    gSetShow pbcProj, slStr, tmOCtrls(ilBox)
    'Next ilBox
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
    imLBCtrls = 1
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    'mParseCmmdLine
    'CntrProj.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    ReDim tgPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPjf2Rec(0 To 1) As PJF2REC
    ReDim tgPjfDel(0 To 1) As PJF2REC
    ReDim smShow(0 To TOTALINDEX, 0 To 1) As String 'Values shown in program area
    ReDim smSave(0 To 8, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmSave(0 To 5, 0 To 1) As Long 'Values saved (program name) in program area
    ReDim imSave(0 To 3, 0 To 1) As Integer 'Values saved (program name) in program area
    ReDim tgPPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPPjf2Rec(0 To 1) As PJF2REC
    ReDim tgOPjf1Rec(0 To UBound(tgPjf1Rec)) As PJF1REC
    ReDim tgOPjf2Rec(0 To UBound(tgPjf2Rec)) As PJF2REC
    imcKey.Picture = IconTraf!imcKey.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    mInitBox
    gCenterStdAlone CntrProj
    imState = 0
    imSlspSelectedIndex = -1
    'CntrProj.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imBoxNo = -1 'Initialize current Box to N/A
    imRowNo = -1
    imPjfChg = False
    imShowIndex = 1 'Standard
    imTypeIndex = 1 'Month
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imIgnoreSetting = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imSettingValue = False
    imDragType = -1
    imPropPopReqd = True
    hmCHF = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", CntrProj
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", CntrProj
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", CntrProj
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
'    hmDsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.Btr)", CntrProj
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    hmPjf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pjf.Btr)", CntrProj
    On Error GoTo 0
    imPjfRecLen = Len(tmPjf)
    ReDim tgPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPjf2Rec(0 To 1) As PJF2REC
    ReDim tgPjfDel(0 To 1) As PJF2REC
    hmSlf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.Btr)", CntrProj
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", CntrProj
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmPrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", CntrProj
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmCxf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", CntrProj
    On Error GoTo 0
    ilRet = gObtainCorpCal()
    Screen.MousePointer = vbHourglass
    lbcSOffice.Clear 'Force list box to be populated
    mSaleOfficePop
    If imTerminate Then
        Exit Sub
    End If
    lbcAdvt.Clear 'Force list box to be populated
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    lbcDemo.Clear 'Force list box to be populated
    mDemoPop
    If imTerminate Then
        Exit Sub
    End If
    lbcVehicle.Clear 'Force list box to be populated
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    lbcNR.AddItem "New"
    lbcNR.AddItem "Return"
    lbcPot.Clear 'Force list box to be populated
    mPotPop
    If imTerminate Then
        Exit Sub
    End If
    lbcPropNo.Clear 'Force list box to be populated
    'mPropNoPop 'Only populate when user select field
    'If imTerminate Then
    '    Exit Sub
    'End If
    cbcSelect.Clear 'Force list box to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    imInHotSpot = False
'    imHotSpot(1, 1) = 4875  'Left
'    imHotSpot(1, 2) = 15    'Top
'    imHotSpot(1, 3) = 4875 + 150 'Right
'    imHotSpot(1, 4) = 15 + 180  'Bottom
'    imHotSpot(2, 1) = 4875 + 150 'Left
'    imHotSpot(2, 2) = 15    'Top
'    imHotSpot(2, 3) = 4875 + 150 + 150 'Right
'    imHotSpot(2, 4) = 15 + 180  'Bottom
'    imHotSpot(3, 1) = 7845  'Left
'    imHotSpot(3, 2) = 15    'Top
'    imHotSpot(3, 3) = 7845 + 150 'Right
'    imHotSpot(3, 4) = 15 + 180  'Bottom
'    imHotSpot(4, 1) = 7845 + 150 'Left
'    imHotSpot(4, 2) = 15    'Top
'    imHotSpot(4, 3) = 7845 + 150 + 150 'Right
'    imHotSpot(4, 4) = 15 + 180  'Bottom
    imHotSpot(1, 1) = pbcLnWkArrow(0).Left  '3945  'Left
    imHotSpot(1, 2) = 15    'Top
    imHotSpot(1, 3) = imHotSpot(1, 1) + 150 '3945 + 150 'Right
    imHotSpot(1, 4) = 15 + 180  'Bottom
    imHotSpot(2, 1) = pbcLnWkArrow(0).Left + 150 '4095  'Left
    imHotSpot(2, 2) = 15    'Top
    imHotSpot(2, 3) = imHotSpot(2, 1) + 150 '4095 + 150 'Right
    imHotSpot(2, 4) = 15 + 180  'Bottom
    imHotSpot(3, 1) = pbcLnWkArrow(1).Left  '7845  'Left
    imHotSpot(3, 2) = 15    'Top
    imHotSpot(3, 3) = imHotSpot(3, 1) + 150 '7845 + 150 'Right
    imHotSpot(3, 4) = 15 + 180  'Bottom
    imHotSpot(4, 1) = pbcLnWkArrow(1).Left + 150 '7995  'Left
    imHotSpot(4, 2) = 15    'Top
    imHotSpot(4, 3) = imHotSpot(4, 1) + 150 '7995 + 150 'Right
    imHotSpot(4, 4) = 15 + 180  'Bottom
    imShowIndex = 1 'Std Month
    imTypeIndex = 0 'Quarter
    If tgSpf.sRUseCorpCal <> "Y" Then
        imIgnoreSetting = True
        rbcShow(1).Value = True
        rbcShow(0).Enabled = False
    End If
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    smMonDate = gObtainPrevMonday(slDate)
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, imCurMonth, imCurYear
    imIgnoreRightMove = False
    Screen.MousePointer = vbDefault
    If tgSpf.sRUseCorpCal = "Y" Then
        If imCurMonth < tgMCof(LBound(tgMCof)).iStartMnthNo - 3 Then
            mTestCorpYear imCurYear
        Else
            mTestCorpYear imCurYear + 1
        End If
    End If
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
    Dim ilLoop As Integer
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcProj.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcProj.Move 165, 630, pbcProj.Width + fgPanelAdj + vbcProj.Width, pbcProj.Height + fgPanelAdj
    pbcProj.Move plcProj.Left + fgBevelX, plcProj.Top + fgBevelY
    vbcProj.Move pbcProj.Left + pbcProj.Width - 15, pbcProj.Top
    pbcArrow.Move plcProj.Left - pbcArrow.Width - 15    'set arrow    'Vehicle
    pbcKey.Move 90, 600
    'Office
    gSetCtrl tmCtrls(SOFFICEINDEX), 30, 420, 615, fgBoxGridH
    'Advertiser
    gSetCtrl tmCtrls(ADVTINDEX), 660, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Product
    gSetCtrl tmCtrls(PRODINDEX), 1560, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'PropNo
    gSetCtrl tmCtrls(PROPNOINDEX), 2460, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Demo
    gSetCtrl tmCtrls(DEMOINDEX), 3360, tmCtrls(SOFFICEINDEX).fBoxY, 210, fgBoxGridH
    'Vehicle
    gSetCtrl tmCtrls(VEHINDEX), 3585, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Revenue Sets
    gSetCtrl tmCtrls(NRINDEX), 4485, tmCtrls(SOFFICEINDEX).fBoxY, 210, fgBoxGridH
    'Potential
    gSetCtrl tmCtrls(POTINDEX), 4710, tmCtrls(SOFFICEINDEX).fBoxY, 210, fgBoxGridH
    'Comment
    gSetCtrl tmCtrls(COMMENTINDEX), 4935, tmCtrls(SOFFICEINDEX).fBoxY, 210, fgBoxGridH
    'Period 1
    gSetCtrl tmCtrls(PD1INDEX), 5160, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Period 2
    gSetCtrl tmCtrls(PD2INDEX), 6060, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Period 3
    gSetCtrl tmCtrls(PD3INDEX), 6960, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Total
    gSetCtrl tmCtrls(TOTALINDEX), 7860, tmCtrls(SOFFICEINDEX).fBoxY, 885, fgBoxGridH
    'Week Dates
    'Week 1
    gSetCtrl tmWKCtrls(WK1INDEX), 5160, 30, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmWKCtrls(WK2INDEX), 6060, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmWKCtrls(WK3INDEX), 6960, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmWKCtrls(WKTINDEX), 7860, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH

    'Number of weeks
    'Week 1
    gSetCtrl tmNWCtrls(WK1INDEX), 5160, 225, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmNWCtrls(WK2INDEX), 6060, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmNWCtrls(WK3INDEX), 6960, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Current Week Advt/Prod
    'Week 1
    gSetCtrl tmCAPCtrls(GTDOLLAR1INDEX), 5160, 3390, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmCAPCtrls(GTDOLLAR2INDEX), 6060, tmCAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmCAPCtrls(GTDOLLAR3INDEX), 6960, tmCAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmCAPCtrls(GTTOTALINDEX), 7860, tmCAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Prior Week Advt/Prod
    'Week 1
    gSetCtrl tmPAPCtrls(GTDOLLAR1INDEX), 5160, 3585, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmPAPCtrls(GTDOLLAR2INDEX), 6060, tmPAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmPAPCtrls(GTDOLLAR3INDEX), 6960, tmPAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmPAPCtrls(GTTOTALINDEX), 7860, tmPAPCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Current Week Totals
    'Week 1
    gSetCtrl tmCTCtrls(GTDOLLAR1INDEX), 5160, 3780, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmCTCtrls(GTDOLLAR2INDEX), 6060, tmCTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmCTCtrls(GTDOLLAR3INDEX), 6960, tmCTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmCTCtrls(GTTOTALINDEX), 7860, tmCTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Prior Week Totals
    'Week 1
    gSetCtrl tmPTCtrls(GTDOLLAR1INDEX), 5160, 3975, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmPTCtrls(GTDOLLAR2INDEX), 6060, tmPTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmPTCtrls(GTDOLLAR3INDEX), 6960, tmPTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmPTCtrls(GTTOTALINDEX), 7860, tmPTCtrls(GTDOLLAR1INDEX).fBoxY, 885, fgBoxGridH

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
            If tmCtrls(ilLoop).fBoxX > 90 Then
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
    pbcProj.Picture = LoadPicture("")
    pbcProj.Width = llMax
    plcProj.Width = llMax + vbcProj.Width + 2 * fgBevelX + 15
    lacFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (CntrProj.Width - 6 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcSave.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcUndo.Left = cmcSave.Left + cmcSave.Width + ilSpaceBetweenButtons
    cmcRollover.Left = cmcUndo.Left + cmcUndo.Width + ilSpaceBetweenButtons
    cmcBlock.Left = cmcRollover.Left + cmcRollover.Width + ilSpaceBetweenButtons
    cmcDone.Top = CntrProj.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    cmcUndo.Top = cmcDone.Top
    cmcRollover.Top = cmcDone.Top
    cmcBlock.Top = cmcDone.Top
    imcTrash.Top = cmcDone.Top + cmcDone.Height - imcTrash.Height
    imcTrash.Left = CntrProj.Width - (3 * imcTrash.Width) / 2
    plcShow.Top = cmcDone.Top - plcShow.Height - 120
    plcType.Top = plcShow.Top
    llAdjTop = plcShow.Top - plcProj.Top - fgBevelY - 120
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    Do While plcProj.Top + llAdjTop + 2 * fgBevelY + 240 < plcShow.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop + 60
    plcProj.Height = llAdjTop + 2 * fgBevelY
    pbcProj.Left = plcProj.Left + fgBevelX
    pbcProj.Top = plcProj.Top + fgBevelY
    pbcProj.Height = plcProj.Height - 2 * fgBevelY
    vbcProj.Left = pbcProj.Left + pbcProj.Width + 15
    vbcProj.Top = pbcProj.Top
    vbcProj.Height = pbcProj.Height
    cbcSelect.Left = plcProj.Left + plcProj.Width - cbcSelect.Width

    For ilLoop = PD1INDEX To TOTALINDEX Step 1
        tmWKCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmWKCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
        tmCAPCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmCAPCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
        tmPAPCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmPAPCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
        tmCTCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmCTCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
        tmPTCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmPTCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
    Next ilLoop
    For ilLoop = PD1INDEX To PD3INDEX Step 1
        tmNWCtrls(ilLoop - PD1INDEX + 1).fBoxX = tmCtrls(ilLoop).fBoxX
        tmNWCtrls(ilLoop - PD1INDEX + 1).fBoxW = tmCtrls(ilLoop).fBoxW
    Next ilLoop

    pbcLnWkArrow(0).Left = tmWKCtrls(WK1INDEX).fBoxX - pbcLnWkArrow(0).Width - 30
    pbcLnWkArrow(0).Top = 15
    pbcLnWkArrow(1).Left = tmWKCtrls(WK1INDEX + 2).fBoxX + tmWKCtrls(WK1INDEX + 2).fBoxW + 60
    pbcLnWkArrow(1).Top = 15
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewProj                    *
'*                                                     *
'*             Created:8/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Projection          *
'*                                                     *
'*******************************************************
Private Sub mInitNewProj()
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, imRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, imRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(imSave, 1) To UBound(imSave, 1) Step 1
        imSave(ilLoop, imRowNo) = 0
    Next ilLoop
    For ilLoop = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
        lmSave(ilLoop, imRowNo) = 0
    Next ilLoop
    imSave(2, imRowNo) = -1
    imSave(3, imRowNo) = False
    ilUpper = UBound(tgPjf1Rec)
    imSave(1, imRowNo) = ilUpper
    tgPjf1Rec(ilUpper).iStatus = 0
    tgPjf1Rec(ilUpper).lRecPos = 0
    tgPjf1Rec(ilUpper).iSaveIndex = imRowNo
    tgPjf1Rec(ilUpper).tPjf.iYear = imCurYear
    tgPjf1Rec(ilUpper).tPjf.iSlfCode = tmSlf.iCode
    tgPjf1Rec(ilUpper).tPjf.lChfCode = 0
    tgPjf1Rec(ilUpper).tPjf.iSofCode = 0
    tgPjf1Rec(ilUpper).tPjf.iAdfCode = 0
    tgPjf1Rec(ilUpper).tPjf.lPrfCode = 0
    tgPjf1Rec(ilUpper).tPjf.iMnfDemo = 0
    tgPjf1Rec(ilUpper).tPjf.iVefCode = 0
    tgPjf1Rec(ilUpper).tPjf.iMnfBus = 0
    tgPjf1Rec(ilUpper).tPjf.lCxfChgR = 0
    For ilLoop = LBound(tgPjf1Rec(ilUpper).tPjf.lGross) To UBound(tgPjf1Rec(ilUpper).tPjf.lGross) Step 1
        tgPjf1Rec(ilUpper).tPjf.lGross(ilLoop) = 0
    Next ilLoop
    If imPrjNoYears > 1 Then
        tgPjf1Rec(ilUpper).i2RecIndex = UBound(tgPjf2Rec)
        tgPjf2Rec(UBound(tgPjf2Rec)).tPjf = tgPjf1Rec(ilUpper).tPjf
        tgPjf2Rec(UBound(tgPjf2Rec)).tPjf.iYear = imCurYear + 1
    Else
        tgPjf1Rec(ilUpper).i2RecIndex = 0
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitProjCtrls                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mInitProjCtrls()
    Dim slDate As String
    Dim slStart As String
    Dim ilYear As Integer
    imPdYear = imCurYear
    imPdStartWk = 1
    imPdStartFltNo = 1
    imPrjStartYear = imPdYear
    ilYear = imPdYear
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartCorp(slDate, True)
    Else                        'Standard
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartStd(slDate)
    End If
    imPrjStartWk = (lmNowDate - gDateValue(slStart)) \ 7 + 1
    imPdStartWk = imPrjStartWk
    'If imCurMonth <> 1 Then
        imPrjNoYears = 2
    'Else
    '    imPrjNoYears = 1
    'End If
    'If UBound(tgPrfRec) > 1 Then
    '    ilLower = LBound(tgPjf1Rec)
    '    imPrjStartYear = tgPjf1Rec(ilLower).tPjf.iYear
    '    imPrjNoYears = 1
    '    'Adjust Period to be viewed
    '    If imPrjStartYear > imPdYear Then
    '        imPdYear = imPrjStartYear
    '    ElseIf imPrjStartYear + imPrjNoYears - 1 < imPdYear Then
    '        imPdYear = imPrjStartYear + imPrjNoYears - 1
    '    End If
    'End If
    'ilUpperBound = UBound(tgPjf1Rec)
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
Private Function mMoveCtrlToRec() As Integer
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim ilPjf As Integer

    mMoveCtrlToRec = 0
    For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
        ilPjf = imSave(1, ilRow)
        tgPjf1Rec(ilPjf).tPjf.iSlfCode = tmSlf.iCode
        'Proposal Number
        'gFindMatch smSave(1, ilRow), 1, lbcPropNo
        'If gLastFound(lbcPropNo) > 0 Then
        '    slNameCode = tgCntrCode(gLastFound(lbcPropNo) - 1).sKey  'Traffic!lbcCntrCode.List(gLastFound(lbcPropNo) - 1)
        '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '    tgPjf1Rec(ilPjf).tPjf.lChfCode = Val(slCode)
        If (smSave(1, ilRow) <> "") And (lmSave(5, ilRow) > 0) Then
            tgPjf1Rec(ilPjf).tPjf.lChfCode = lmSave(5, ilRow)
        Else
            tgPjf1Rec(ilPjf).tPjf.lChfCode = 0
        End If
        'Selling Office
        gFindMatch smSave(2, ilRow), 1, lbcSOffice
        If gLastFound(lbcSOffice) > 0 Then
            slNameCode = tmSOfficeCode(gLastFound(lbcSOffice) - 1).sKey 'lbcSOfficeCode.List(gLastFound(lbcSOffice) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.iSofCode = Val(slCode)
        Else
            imRowNo = ilRow
            mMoveCtrlToRec = SOFFICEINDEX
            Exit Function
        End If
        'Advertiser
        gFindMatch smSave(3, ilRow), 1, lbcAdvt
        If gLastFound(lbcAdvt) > 0 Then
            slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.iAdfCode = Val(slCode)
        Else
            imRowNo = ilRow
            mMoveCtrlToRec = ADVTINDEX
            Exit Function
        End If
        'Product
        mProdPop tgPjf1Rec(ilPjf).tPjf.iAdfCode
        gFindMatch smSave(4, ilRow), 1, lbcProd
        If gLastFound(lbcProd) > 0 Then
            slNameCode = tgProdCode(gLastFound(lbcProd) - 1).sKey    'Traffic!lbcProdCode.List(gLastFound(lbcProd) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.lPrfCode = Val(slCode)
        Else
            If (smSave(4, ilRow) <> "") And (smSave(4, ilRow) <> "[None]") Then
                tgPjf1Rec(ilPjf).tPjf.lPrfCode = mAddProd(tgPjf1Rec(ilPjf).tPjf.iAdfCode, smSave(4, ilRow))
            Else
                tgPjf1Rec(ilPjf).tPjf.lPrfCode = 0
            End If
        End If
        'Demo
        gFindMatch smSave(5, ilRow), 1, lbcDemo
        If gLastFound(lbcDemo) > 0 Then
            slNameCode = tgDemoCode(gLastFound(lbcDemo) - 1).sKey    'Traffic!lbcDemoCode.List(gLastFound(lbcDemo) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.iMnfDemo = Val(slCode)
        Else
            tgPjf1Rec(ilPjf).tPjf.iMnfDemo = 0
        End If
        'Vehicle
        gFindMatch smSave(6, ilRow), 0, lbcVehicle
        If gLastFound(lbcVehicle) >= 0 Then
            slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey  'Traffic!lbcUserVehicle.List(gLastFound(lbcVehicle))
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.iVefCode = Val(slCode)
        Else
            imRowNo = ilRow
            mMoveCtrlToRec = VEHINDEX
            Exit Function
        End If
        If imSave(2, ilRow) = 0 Then
            tgPjf1Rec(ilPjf).tPjf.sNewRet = "N"
        Else
            tgPjf1Rec(ilPjf).tPjf.sNewRet = "R"
        End If
        'Potential Business
        gFindMatch smSave(7, ilRow), 1, lbcPot
        If gLastFound(lbcPot) >= 1 Then    'Bypass [None]
            slNameCode = tgPotCode(gLastFound(lbcPot) - 1).sKey  'Traffic!lbcPotCode.List(gLastFound(lbcPot) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPjf1Rec(ilPjf).tPjf.iMnfBus = Val(slCode)
        Else
            tgPjf1Rec(ilPjf).tPjf.iMnfBus = 0
        End If
        'Comment- processed in the save logic
    Next ilRow
    For ilLoop = LBONE To UBound(tgPjf1Rec) - 1 Step 1
        ilPjf = tgPjf1Rec(ilLoop).i2RecIndex
        If ilPjf > 0 Then
            tgPjf2Rec(ilPjf).tPjf.iSlfCode = tmSlf.iCode
            tgPjf2Rec(ilPjf).tPjf.lChfCode = tgPjf1Rec(ilLoop).tPjf.lChfCode
            tgPjf2Rec(ilPjf).tPjf.iSofCode = tgPjf1Rec(ilLoop).tPjf.iSofCode
            tgPjf2Rec(ilPjf).tPjf.iAdfCode = tgPjf1Rec(ilLoop).tPjf.iAdfCode
            tgPjf2Rec(ilPjf).tPjf.lPrfCode = tgPjf1Rec(ilLoop).tPjf.lPrfCode
            tgPjf2Rec(ilPjf).tPjf.iMnfDemo = tgPjf1Rec(ilLoop).tPjf.iMnfDemo
            tgPjf2Rec(ilPjf).tPjf.iVefCode = tgPjf1Rec(ilLoop).tPjf.iVefCode
            tgPjf2Rec(ilPjf).tPjf.sNewRet = tgPjf1Rec(ilLoop).tPjf.sNewRet
            tgPjf2Rec(ilPjf).tPjf.iMnfBus = tgPjf1Rec(ilLoop).tPjf.iMnfBus
        End If
    Next ilLoop
    Exit Function

    On Error GoTo 0
    'imTerminate = True
    Exit Function
End Function
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
Private Sub mMoveRecToCtrl(ilCallGetPrice As Integer)
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
    Dim ilRowNo As Integer
    Dim ilUpper As Integer
    Dim slNoWks As String
    Dim slStartDate As String
    Dim llEndDate As Long
    ilUpper = UBound(tgPjf1Rec)
    ReDim smShow(0 To TOTALINDEX, 0 To ilUpper) As String 'Values shown in program area
    'ReDim smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
    ReDim smSave(0 To 8, 0 To ilUpper) As String 'Values saved (program name) in program area
    ReDim lmSave(0 To 5, 0 To ilUpper) As Long 'Values saved (program name) in program area
    ReDim imSave(0 To 3, 0 To ilUpper) As Integer 'Values saved (program name) in program area
    'Init value in the case that no records are associated with the salesperson
    If ilUpper = LBONE Then
        ilRowNo = imRowNo
        imRowNo = 1
        mInitNewProj
        imRowNo = ilRowNo
        'imSave(1, LBound(tgPjf1Rec)) = ilUpper
        'tgPjf1Rec(ilUpper).iStatus = 0
        'tgPjf1Rec(ilUpper).lRecPos = 0
        'tgPjf1Rec(ilUpper).iSaveIndex = ilUpper 'imRowNo
        'tgPjf1Rec(ilUpper).tPjf.iYear = imCurYear
        'If imPrjNoYears > 1 Then
        '    tgPjf1Rec(ilUpper).i2RecIndex = UBound(tgPjf2Rec)
        '    tgPjf2Rec(UBound(tgPjf2Rec)).tPjf.iYear = imCurYear
        'Else
        '    tgPjf1Rec(ilUpper).i2RecIndex = 0
        'End If
    End If
    'lmSave(5, ilRowNo) = 0
    For ilRowNo = LBONE To UBound(tgPjf1Rec) - 1 Step 1
        tgPjf1Rec(ilRowNo).iSaveIndex = ilRowNo
        lmSave(5, ilRowNo) = 0
        imSave(1, ilRowNo) = ilRowNo
        imSave(3, ilRowNo) = False
        'Get Proposal Number
        tmChf.sProduct = ""
        tmChf.iMnfDemo(0) = 0
        If tgPjf1Rec(ilRowNo).tPjf.lChfCode > 0 Then
            tmChfSrchKey.lCode = tgPjf1Rec(ilRowNo).tPjf.lChfCode 'ilCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                lmSave(5, ilRowNo) = tmChf.lCode
                If tmChf.iCntRevNo > 0 Then
                    slName = Trim$(Str$(tmChf.lCntrNo)) & " R" & Trim$(Str$(tmChf.iCntRevNo)) & "-" & Trim$(Str$(tmChf.iExtRevNo))
                Else
                    slName = Trim$(Str$(tmChf.lCntrNo)) & " V" & Trim$(Str$(tmChf.iPropVer))
                End If
                Select Case tmChf.sStatus
                    Case "W"
                        If tmChf.iCntRevNo > 0 Then
                            slStr = "Rev Working"
                        Else
                            slStr = "Working"
                        End If
                    Case "D"
                        slStr = "Rejected"
                    Case "C"
                        If tmChf.iCntRevNo > 0 Then
                            slStr = "Rev Completed"
                        Else
                            slStr = "Completed"
                        End If
                    Case "I"
                        If tmChf.iCntRevNo > 0 Then
                            slStr = "Rev Unapproved"
                        Else
                            slStr = "Unapproved"
                        End If
                    Case "G"
                        slStr = "Approved Hold"
                    Case "N"
                        slStr = "Approved Order"
                    Case "H"
                        slStr = "Hold"
                    Case "O"
                        slStr = "Order"
                End Select
                slName = slName & " " & slStr
                slStr = slName & " " & Trim$(tmChf.sProduct)
                'Start Date Plus # weeks
                gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate
                gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), llEndDate
                slNoWks = Str$((llEndDate - gDateValue(slStartDate)) \ 7 + 1)
                slStr = slStr & " " & slStartDate & slNoWks
                tmCxfSrchKey.lCode = tmChf.lCxfInt
                If tmCxfSrchKey.lCode <> 0 Then
                    tmCxf.sComment = ""
                    imCxfRecLen = Len(tmCxf) '5027
                    ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If ilRet = BTRV_ERR_NONE Then
                    '    If tmCxf.iStrLen > 0 Then
                    '        If tmCxf.iStrLen < 40 Then
                    '            slStr = slStr & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                    '        Else
                    '            slStr = slStr & " " & Trim$(Left$(tmCxf.sComment, 40))
                    '        End If
                    '    End If
                    'End If
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = slStr & " " & gStripChr0(Left$(tmCxf.sComment, 40))
                    End If
                End If
            Else
                slStr = ""
            End If
        Else
            slStr = ""
        End If
        smSave(1, ilRowNo) = Trim$(slStr)
        gSetShow pbcProj, slStr, tmCtrls(PROPNOINDEX)
        smShow(PROPNOINDEX, ilRowNo) = tmCtrls(PROPNOINDEX).sShow
        'Get Selling Office
        For ilLoop = LBONE To UBound(tmSaleOffice) - 1 Step 1
            If tmSaleOffice(ilLoop).iCode = tgPjf1Rec(ilRowNo).tPjf.iSofCode Then
                smSave(2, ilRowNo) = Trim$(tmSaleOffice(ilLoop).sName)
                slStr = smSave(2, ilRowNo)
                gSetShow pbcProj, slStr, tmCtrls(SOFFICEINDEX)
                smShow(SOFFICEINDEX, ilRowNo) = tmCtrls(SOFFICEINDEX).sShow
                Exit For
            End If
        Next ilLoop
        'Get advertiser
        smSave(3, ilRowNo) = Trim$(tgPjf1Rec(ilRowNo).sAdvtName)
        slStr = smSave(3, ilRowNo)
        gSetShow pbcProj, slStr, tmCtrls(ADVTINDEX)
        smShow(ADVTINDEX, ilRowNo) = tmCtrls(ADVTINDEX).sShow
        'Get product
        smSave(4, ilRowNo) = Trim$(tgPjf1Rec(ilRowNo).sProdName)
        slStr = smSave(4, ilRowNo)
        gSetShow pbcProj, slStr, tmCtrls(PRODINDEX)
        smShow(PRODINDEX, ilRowNo) = tmCtrls(PRODINDEX).sShow
        'Get Demo
        If tgPjf1Rec(ilRowNo).tPjf.iMnfDemo > 0 Then
            For ilLoop = 0 To UBound(tgDemoCode) - 1 Step 1 'Traffic!lbcDemoCode.ListCount - 1 Step 1
                slNameCode = tgDemoCode(ilLoop).sKey   'Traffic!lbcDemoCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tgPjf1Rec(ilRowNo).tPjf.iMnfDemo Then
                    slStr = Trim$(lbcDemo.List(ilLoop + 1))
                    Exit For
                End If
            Next ilLoop
        Else
            slStr = ""
        End If
        smSave(5, ilRowNo) = slStr
        gSetShow pbcProj, slStr, tmCtrls(DEMOINDEX)
        smShow(DEMOINDEX, ilRowNo) = tmCtrls(DEMOINDEX).sShow
        'Vehicle
        For ilLoop = LBONE To UBound(tmUserVeh) - 1 Step 1
            If tmUserVeh(ilLoop).iCode = tgPjf1Rec(ilRowNo).tPjf.iVefCode Then
                smSave(6, ilRowNo) = Trim$(tmUserVeh(ilLoop).sName)
                slStr = smSave(6, ilRowNo)
                gSetShow pbcProj, slStr, tmCtrls(VEHINDEX)
                smShow(VEHINDEX, ilRowNo) = tmCtrls(VEHINDEX).sShow
                Exit For
            End If
        Next ilLoop
        'New or Return
        If tgPjf1Rec(ilRowNo).tPjf.sNewRet = "N" Then
            imSave(2, ilRowNo) = 0
            slStr = "New"
        Else
            imSave(2, ilRowNo) = 1
            slStr = "Return"
        End If
        gSetShow pbcProj, slStr, tmCtrls(NRINDEX)
        smShow(NRINDEX, ilRowNo) = tmCtrls(NRINDEX).sShow
        'Get Potential
        If tgPjf1Rec(ilRowNo).tPjf.iMnfBus > 0 Then
            For ilLoop = 0 To UBound(tgPotCode) - 1 Step 1  'Traffic!lbcPotCode.ListCount - 1 Step 1
                slNameCode = tgPotCode(ilLoop).sKey    'Traffic!lbcPotCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tgPjf1Rec(ilRowNo).tPjf.iMnfBus Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    slStr = Trim$(slName)
                    Exit For
                End If
            Next ilLoop
            smSave(7, ilRowNo) = slStr
        Else
            slStr = ""
            smSave(7, ilRowNo) = ""
        End If
        gSetShow pbcProj, slStr, tmCtrls(POTINDEX)
        smShow(POTINDEX, ilRowNo) = tmCtrls(POTINDEX).sShow
        'Comment
        tmCxfSrchKey.lCode = tgPjf1Rec(ilRowNo).tPjf.lCxfChgR
        If tgPjf1Rec(ilRowNo).tPjf.lCxfChgR <> 0 Then
            tmCxf.sComment = ""
            imCxfRecLen = Len(tmCxf)
            ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCxf.lCode = 0
                'tmCxf.iStrLen = 0
                tmCxf.sComment = ""
            End If
        Else
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
        End If
        'If tmCxf.iStrLen > 0 Then
        '    smSave(8, ilRowNo) = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
        'Else
        '    smSave(8, ilRowNo) = ""
        'End If
        smSave(8, ilRowNo) = gStripChr0(tmCxf.sComment)
        slStr = smSave(8, ilRowNo)
        gSetShow pbcProj, slStr, tmCtrls(COMMENTINDEX)
        smShow(COMMENTINDEX, ilRowNo) = tmCtrls(COMMENTINDEX).sShow
    Next ilRowNo
    imSettingValue = True
    vbcProj.Min = LBONE
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcProj.LargeChange + 1 Then ' + 1 Then
        vbcProj.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcProj.Max = UBound(smShow, 2) - vbcProj.LargeChange
    End If
    imSettingValue = True
    vbcProj.Value = vbcProj.Min
    imSettingValue = True
    mGetShowDates ilCallGetPrice
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPjfGetPrice                    *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get price for date specified   *
'*                                                     *
'*******************************************************
Private Sub mPjfGetGross(ilPjf1 As Integer, slInDate As String, tlPjf1Rec() As PJF1REC, tlPjf2Rec() As PJF2REC, llGrossAmount As Long)
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    Dim ilPjf2 As Integer
    llGrossAmount = 0
    gObtainMonthYear 0, slInDate, ilMonth, ilYear
    gObtainWkNo 0, slInDate, ilWkNo, ilFirstLastWk
    If (ilWkNo > 0) And (ilWkNo < 54) Then
        If tlPjf1Rec(ilPjf1).tPjf.iYear = ilYear Then
            llGrossAmount = tlPjf1Rec(ilPjf1).tPjf.lGross(ilWkNo)
        Else
            ilPjf2 = tlPjf1Rec(ilPjf1).i2RecIndex
            If ilPjf2 > 0 Then
                If tlPjf2Rec(ilPjf2).tPjf.iYear = ilYear Then
                    llGrossAmount = tlPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo)
                End If
            End If
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPjfSetPrice                    *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get price for date specified   *
'*                                                     *
'*******************************************************
Private Sub mPjfSetGross(ilPjf1 As Integer, slInDate As String, llGrossAmount As Long, tlPjf1Rec() As PJF1REC, tlPjf2Rec() As PJF2REC)
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    Dim ilPjf2 As Integer
    gObtainMonthYear 0, slInDate, ilMonth, ilYear
    gObtainWkNo 0, slInDate, ilWkNo, ilFirstLastWk
    If (ilWkNo > 0) And (ilWkNo < 54) Then
        If tlPjf1Rec(ilPjf1).tPjf.iYear = ilYear Then
            tlPjf1Rec(ilPjf1).tPjf.lGross(ilWkNo) = llGrossAmount
        Else
            ilPjf2 = tlPjf1Rec(ilPjf1).i2RecIndex
            If ilPjf2 > 0 Then
                If tlPjf2Rec(ilPjf2).tPjf.iYear = ilYear Then
                    tlPjf2Rec(ilPjf2).tPjf.lGross(ilWkNo) = llGrossAmount
                End If
            End If
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salesperson list box  *
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
    Dim ilType As Integer
    ilIndex = cbcSelect.ListIndex
    If ilIndex > 1 Then
        slName = cbcSelect.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopCntrProjComboBox(CntrProj, cbcSelect, Traffic!lbcSaleCntrProj, Traffic!cbcSelectCombo, igSlfFirstNameFirst)
    ilType = 0  '0=All;1=Salesperson and Negotiator;4=Salespersons, Negotiator and Planner
    'ilRet = gPopSalespersonBox(CntrProj, ilType, True, True, cbcSelect, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(CntrProj, ilType, True, False, cbcSelect, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopluateErr
        gCPErrorMsg ilRet, "mPopluate (gIMoveListBox)", CntrProj
        On Error GoTo 0
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
    End If
    Exit Sub
mPopluateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPotBranch                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      potential and process          *
'*                      communication back from        *
'*                      potential                      *
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
Private Function mPotBranch() As Integer
'
'   ilRet = mPotBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcPot, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mPotBranch = False
        Exit Function
    End If
    If igWinStatus(POTENTIALCODESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mPotBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(COMPETITIVESLIST)) Then
    '    imDoubleClickName = False
    '    mPotBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    sgMnfCallType = "P"
    igMNmCallSource = CALLSOURCECONTRACT
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CntrProj^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "CntrProj^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CntrProj^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "CntrProj^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'CntrProj.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'CntrProj.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mPotBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcPot.Clear
        sgPotCodeTag = ""
        sgPotMnfStamp = ""
        mPotPop
        If imTerminate Then
            mPotBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcPot
        If gLastFound(lbcPot) > 0 Then
            imChgMode = True
            lbcPot.ListIndex = gLastFound(lbcPot)
            edcDropDown.Text = lbcPot.List(lbcPot.ListIndex)
            imChgMode = False
            mPotBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcPot.ListIndex = 1
            edcDropDown.Text = lbcPot.List(1)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
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
'*      Procedure Name:mPotPop                         *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Potential Code        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mPotPop()
'
'   mPotPop
'   Where:
'
    ReDim ilfilter(0 To 0) As Integer
    ReDim slFilter(0 To 0) As String
    ReDim ilOffSet(0 To 0) As Integer
    Dim ilRet As Integer
    Dim slPot As String      'Potential name, saved to determine if changed
    Dim ilIndex As Integer      'Potential name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "P"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilIndex = lbcPot.ListIndex
    If ilIndex > 1 Then
        slPot = lbcPot.List(ilIndex)
    End If
    'ilRet = gIMoveListBox(CntrProj, lbcPot, Traffic!lbcPotCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CntrProj, lbcPot, tgPotCode(), sgPotCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPotPopErr
        gCPErrorMsg ilRet, "mPotPop (gIMoveListBox)", CntrProj
        On Error GoTo 0
        'lbcPot.AddItem "[None]", 0
        lbcPot.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slPot, 1, lbcPot
            If gLastFound(lbcPot) > 0 Then
                lbcPot.ListIndex = gLastFound(lbcPot)
            Else
                lbcPot.ListIndex = -1
            End If
        Else
            lbcPot.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mPotPopErr:
    On Error GoTo 0
    imTerminate = True
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
'    If imSlspSelectedIndex = 0 Then 'New selected
'        imDoubleClickName = False
'        mProdBranch = False
'        Exit Function
'    End If
    If (Not imDoubleClickName) Or (edcDropDown.Text = "[None]") Then
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
    igAdvtProdCallSource = CALLSOURCECONTRACT
    gFindMatch smSave(3, imRowNo), 1, lbcAdvt
    If gLastFound(lbcAdvt) > 0 Then
        sgAdvtProdName = lbcAdvt.List(gLastFound(lbcAdvt))
    Else
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CntrProj^Test\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        Else
            slStr = "CntrProj^Prod\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CntrProj^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    Else
    '        slStr = "CntrProj^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AdvtProd.Exe " & slStr, 1)
    'CntrProj.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AdvtProd.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'CntrProj.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProdBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
   ' If imUpdateAllowed = False Then
   '     mSendHelpMsg "BF"
   ' Else
   '     mSendHelpMsg "BT"
   ' End If
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
            sgAdvtProdName = ""
            Exit Function
        End If
        sgAdvtProdName = ""
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
    'ilRet = gPopAdvtProdBox(CntrProj, ilAdvtCode, lbcProd, Traffic!lbcProdCode)
    ilRet = gPopAdvtProdBox(CntrProj, ilAdvtCode, lbcProd, tgProdCode(), sgProdCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gIMoveListBox)", CntrProj
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
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPropNoPop                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Proposal Population            *
'*                                                     *
'*******************************************************
Private Sub mPropNoPop(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim slName As String
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilShow As Integer
    Dim ilState As Integer
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    'If Not imPropPopReqd Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    'ilIndex = lbcPropNo.ListIndex
    'If ilIndex > 0 Then
    '    slName = lbcPropNo.List(ilIndex)
    'End If
    slName = smSave(1, ilRowNo)
    'slCntrStatus = "WCI" 'Working; Incomplete; Complete
    'Only allow complete to be specified abc request 1/4/2000
    slCntrStatus = "C" 'Complete
    slCntrType = "C" 'Standard only
    ilCurrent = 1
    ilShow = 1
    ilState = 0
    'ilRet = gPopCntrForAASBox(CntrProj, -1, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcPropNo, Traffic!lbcCntrCode)
    ilAAS = -1
    ilAASCode = 0
    If ilRowNo > 0 Then
        gFindMatch smSave(3, ilRowNo), 1, lbcAdvt
        If gLastFound(lbcAdvt) > 0 Then
            slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAAS = 0   'By Advertiser
            ilAASCode = Val(slCode)
        Else
            ilAAS = 3   'By Salesperson
            ilAASCode = tmSlf.iCode
        End If
    End If
    ilRet = gPopCntrForAASBox(CntrProj, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcPropNo, tgCntrCode(), sgCntrCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPropNoErr
        gCPErrorMsg ilRet, "mPropNo (gPopCntrForAASBox)", CntrProj
        On Error GoTo 0
        lbcPropNo.AddItem "[None]", 0
        imChgMode = True
        If Len(Trim(slName)) > 0 Then
            gFindMatch slName, 1, lbcPropNo
            If gLastFound(lbcPropNo) > 0 Then
                lbcPropNo.ListIndex = gLastFound(lbcPropNo)
            Else
                lbcPropNo.ListIndex = -1
            End If
        Else
            lbcPropNo.ListIndex = -1
        End If
        imChgMode = False
    Else
        If Len(Trim(slName)) > 0 Then
            gFindMatch slName, 1, lbcPropNo
            If gLastFound(lbcPropNo) > 0 Then
                lbcPropNo.ListIndex = gLastFound(lbcPropNo)
            Else
                lbcPropNo.ListIndex = -1
            End If
        Else
            lbcPropNo.ListIndex = -1
        End If
    End If
    imPropPopReqd = False
    Screen.MousePointer = vbDefault
    Exit Sub
mPropNoErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadPjfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadPjfRec(ilSlfCode As Integer) As Integer
'
'   iRet = mReadPjfRec (ilSlfCode)
'   Where:
'       ilSlfCode(I)- Slf Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilTest As Integer
    Dim ilPjf As Integer
    Dim ilFound As Integer
    Dim llGross As Long
    Dim ilWk As Integer
    Dim slDate As String
    Dim slStart As String
    Dim llDate As Long
    Dim ilPrjStartWk As Integer
    Dim ilRollForward As Integer
    Dim ilStartWk As Integer
    'Dim slDate As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tgPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPjf2Rec(0 To 1) As PJF2REC
    ReDim tgPjfDel(0 To 1) As PJF2REC
    imPjfChg = False

    tmSlfSrchKey.iCode = ilSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mReadPjfRec = False
        Exit Function
    End If
    ilUpper = UBound(tgPjf1Rec)
    btrExtClear hmPjf   'Clear any previous extend operation
    ilExtLen = Len(tgPjf1Rec(1).tPjf)  'Extract operation record size
    tmPjfSrchKey.iSlfCode = ilSlfCode
    'slDate = gDecOneWeek(gObtainPrevMonday(slEffDate))
    'gPackDate slDate, tmPjfSrchKey.iEffDate(0), tmPjfSrchKey.iEffDate(1)
    tmPjfSrchKey.iRolloverDate(0) = 0
    tmPjfSrchKey.iRolloverDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmPjf, tgPjf1Rec(1).tPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hmPjf, tgPjf1Rec(1).tPjf, imBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmPjf, llNoRec, -1, "UC", "PJF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Pjf", "PjfSlfCode")
        tlIntTypeBuff.iType = ilSlfCode
        ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        On Error GoTo mReadPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPjfRec (btrExtAddLogicConst):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        'gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Pjf", "PjfEffDate")
        'ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        tlDateTypeBuff.iDate0 = 0
        tlDateTypeBuff.iDate1 = 0
        ilOffSet = gFieldOffset("Pjf", "PjfRolloverDate")
        ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        On Error GoTo mReadPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPjfRec (btrExtAddLogicConst):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        ilRet = btrExtAddField(hmPjf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPjfRec (btrExtAddField):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmPjf, tgPjf1Rec(ilUpper).tPjf, ilExtLen, tgPjf1Rec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadPjfRecErr
            gBtrvErrorMsg ilRet, "mReadPjfRec (btrExtGetNextExt):" & "Pjf.Btr", CntrProj
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tgPjf1Rec(1).tPjf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmPjf, tgPjf1Rec(ilUpper).tPjf, ilExtLen, tgPjf1Rec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                ilRecOK = True
                If tgPjf1Rec(ilUpper).tPjf.iAdfCode <> tmAdf.iCode Then
                    tmAdfSrchKey.iCode = tgPjf1Rec(ilUpper).tPjf.iAdfCode 'ilCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If tgPjf1Rec(ilUpper).tPjf.lPrfCode > 0 Then
                        If tgPjf1Rec(ilUpper).tPjf.lPrfCode <> tmPrf.lCode Then
                            tmPrfSrchKey.lCode = tgPjf1Rec(ilUpper).tPjf.lPrfCode 'ilCode
                            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRecOK = False
                        End If
                    Else
                        tmPrf.sName = ""
                        tmPrf.lCode = 0
                    End If
                Else
                    ilRecOK = False
                End If
                If ilRecOK Then
                    ilFound = False
                    For ilTest = LBONE To UBound(tgPjf1Rec) - 1 Step 1
                        If (tgPjf1Rec(ilTest).tPjf.iSofCode = tgPjf1Rec(ilUpper).tPjf.iSofCode) And (tgPjf1Rec(ilTest).tPjf.iAdfCode = tgPjf1Rec(ilUpper).tPjf.iAdfCode) And (tgPjf1Rec(ilTest).tPjf.lPrfCode = tgPjf1Rec(ilUpper).tPjf.lPrfCode) And (tgPjf1Rec(ilTest).tPjf.iVefCode = tgPjf1Rec(ilUpper).tPjf.iVefCode) And (tgPjf1Rec(ilTest).tPjf.lChfCode = tgPjf1Rec(ilUpper).tPjf.lChfCode) And (tgPjf1Rec(ilTest).tPjf.iYear <> tgPjf1Rec(ilUpper).tPjf.iYear) Then
                            ilFound = True
                            If tgPjf1Rec(ilTest).tPjf.iYear > tgPjf1Rec(ilUpper).tPjf.iYear Then
                                'Swap records
                                tmPjf = tgPjf1Rec(ilTest).tPjf
                                llRecPos = tgPjf1Rec(ilTest).lRecPos
                                tgPjf1Rec(ilTest).tPjf = tgPjf1Rec(ilUpper).tPjf
                                tgPjf1Rec(ilTest).lRecPos = tgPjf1Rec(ilUpper).lRecPos
                                tgPjf1Rec(ilUpper).tPjf = tmPjf
                                tgPjf1Rec(ilUpper).lRecPos = llRecPos
                            End If
                            tgPjf1Rec(ilTest).i2RecIndex = UBound(tgPjf2Rec)    'ilTest
                            tgPjf2Rec(UBound(tgPjf2Rec)).tPjf = tgPjf1Rec(ilUpper).tPjf
                            tgPjf2Rec(UBound(tgPjf2Rec)).iStatus = 1
                            tgPjf2Rec(UBound(tgPjf2Rec)).lRecPos = tgPjf1Rec(ilUpper).lRecPos
                            ReDim Preserve tgPjf2Rec(0 To UBound(tgPjf2Rec) + 1) As PJF2REC
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                            tgPjf1Rec(ilUpper).sKey = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & tmPrf.sName
                            tgPjf1Rec(ilUpper).sAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                        Else
                            tgPjf1Rec(ilUpper).sKey = tmAdf.sName & tmPrf.sName
                            tgPjf1Rec(ilUpper).sAdvtName = tmAdf.sName
                        End If
                        tgPjf1Rec(ilUpper).sProdName = tmPrf.sName
                        tgPjf1Rec(ilUpper).iStatus = 1
                        ilUpper = ilUpper + 1
                        ReDim Preserve tgPjf1Rec(0 To ilUpper) As PJF1REC
                    End If
                End If
                ilRet = btrExtGetNext(hmPjf, tgPjf1Rec(ilUpper).tPjf, ilExtLen, tgPjf1Rec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmPjf, tgPjf1Rec(ilUpper).tPjf, ilExtLen, tgPjf1Rec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgPjf1Rec(), 1), UBound(tgPjf1Rec) - 1, 0, LenB(tgPjf1Rec(1)), 0, LenB(tgPjf1Rec(1).sKey), 0
    End If
    If imPrjNoYears > 1 Then
        For ilLoop = LBONE To UBound(tgPjf1Rec) - 1 Step 1
            If tgPjf1Rec(ilLoop).i2RecIndex <= 0 Then
                tgPjf1Rec(ilLoop).i2RecIndex = UBound(tgPjf2Rec)
                tgPjf2Rec(UBound(tgPjf2Rec)).tPjf.iYear = tgPjf1Rec(ilLoop).tPjf.iYear + 1  'imCurYear + 1
                tgPjf2Rec(UBound(tgPjf2Rec)).iStatus = 0
                tgPjf2Rec(UBound(tgPjf2Rec)).lRecPos = 0
                ReDim Preserve tgPjf2Rec(0 To UBound(tgPjf2Rec) + 1) As PJF2REC
            End If
        Next ilLoop
    End If
    ReDim tgOPjf1Rec(0 To UBound(tgPjf1Rec)) As PJF1REC
    ReDim tgOPjf2Rec(0 To UBound(tgPjf2Rec)) As PJF2REC
    'Roll old weeks forward if in same month
    slDate = "1/15/" & Trim$(Str$(imCurYear))
    'If tgSpf.sRUseCorpCal = "Y" Then
    '    slStart = gObtainStartCorp(slDate, True)
    'Else
        slStart = gObtainStartStd(slDate)
        ilPrjStartWk = (lmNowDate - gDateValue(slStart)) \ 7 + 1
    'End If
    ilStartWk = 1
    ilRollForward = False
    'For ilLoop = 1 To 12 Step 1
    '    If tgSpf.sRUseCorpCal = "Y" Then
    '        slEnd = gObtainEndCorp(slStart, True)
    '    Else
    '        slEnd = gObtainEndStd(slStart)
    '    End If
    '    ilNoWks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
    '    'If first week: then ignore roll-forward
    '    If (ilPrjStartWk - 1 >= ilStartWk) And (ilPrjStartWk - 1 <= ilStartWk + ilNoWks - 1) Then
    '        If (ilPrjStartWk >= ilStartWk) And (ilPrjStartWk <= ilStartWk + ilNoWks - 1) Then
    '            ilRollForward = True
    '        End If
    '        Exit For
    '    End If
    '    ilStartWk = ilStartWk + ilNoWks
    '    slStart = gIncOneDay(slEnd)
    'Next ilLoop
    If tgSpf.sRUseCorpCal = "Y" Then
        'If first week of corp: don't roll-forward
        slDate = Format$(lmNowDate, "m/d/yy")
        slStart = gObtainStartCorp(slDate, True)
        llDate = lmNowDate
        Do While gWeekDayLong(llDate) <> 0
            llDate = llDate - 1
        Loop
        If gDateValue(slStart) <> llDate Then
            ilRollForward = True
        End If
    Else
        'If first week of std: don't roll-forward
        slDate = Format$(lmNowDate, "m/d/yy")
        slStart = gObtainStartStd(slDate)
        llDate = lmNowDate
        Do While gWeekDayLong(llDate) <> 0
            llDate = llDate - 1
        Loop
        If gDateValue(slStart) <> llDate Then
            ilRollForward = True
        End If
    End If
    For ilPjf = LBONE To UBound(tgPjf1Rec) - 1 Step 1
        llGross = 0
        For ilWk = 0 To ilPrjStartWk - 1 Step 1
            llGross = llGross + tgPjf1Rec(ilPjf).tPjf.lGross(ilWk)
            tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = 0
        Next ilWk
        If ilRollForward Then
            tgPjf1Rec(ilPjf).tPjf.lGross(ilPrjStartWk) = llGross + tgPjf1Rec(ilPjf).tPjf.lGross(ilPrjStartWk)
        End If
    Next ilPjf
    For ilPjf = LBound(tgPjf1Rec) To UBound(tgPjf1Rec) - 1 Step 1
        tgOPjf1Rec(ilPjf) = tgPjf1Rec(ilPjf)
    Next ilPjf
    For ilPjf = LBound(tgPjf2Rec) To UBound(tgPjf2Rec) - 1 Step 1
        tgOPjf2Rec(ilPjf) = tgPjf2Rec(ilPjf)
    Next ilPjf
    'mInitBudgetCtrls
    mReadPjfRec = True
    Exit Function
mReadPjfRecErr:
    On Error GoTo 0
    mReadPjfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadPPjfRec                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read prior week records        *
'*                                                     *
'*******************************************************
Private Function mReadPPjfRec(ilSlfCode As Integer, slRollDate As String) As Integer
'
'   iRet = mReadPPjfRec (ilSlfCode, slRollDate)
'   Where:
'       ilSlfCode(I)- Slf Code
'       slRollDate(I)- Rollover Date
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tgPPjf1Rec(0 To 1) As PJF1REC
    ReDim tgPPjf2Rec(0 To 1) As PJF2REC
    If slRollDate = "" Then
        mReadPPjfRec = True
        Exit Function
    End If
    tmSlfSrchKey.iCode = ilSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mReadPPjfRec = False
        Exit Function
    End If
    ilUpper = UBound(tgPPjf1Rec)
    btrExtClear hmPjf   'Clear any previous extend operation
    ilExtLen = Len(tgPPjf1Rec(1).tPjf)  'Extract operation record size
    tmPjfSrchKey.iSlfCode = ilSlfCode
    'slDate = gDecOneWeek(gObtainPrevMonday(slRollDate))
    'gPackDate slDate, tmPjfSrchKey.iEffDate(0), tmPjfSrchKey.iEffDate(1)
    gPackDate slRollDate, tmPjfSrchKey.iRolloverDate(0), tmPjfSrchKey.iRolloverDate(1)
    ilRet = btrGetGreaterOrEqual(hmPjf, tgPPjf1Rec(1).tPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hmPjf, tgPPjf1Rec(1).tPjf, imBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmPjf, llNoRec, -1, "UC", "PJF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Pjf", "PjfSlfCode")
        tlIntTypeBuff.iType = ilSlfCode
        ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        On Error GoTo mReadPPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPPjfRec (btrExtAddLogicConst):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        'gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Pjf", "PjfEffDate")
        'ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        'tlDateTypeBuff.iDate0 = 0
        'tlDateTypeBuff.iDate1 = 0
        gPackDate slRollDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Pjf", "PjfRolloverDate")
        ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        On Error GoTo mReadPPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPPjfRec (btrExtAddLogicConst):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        ilRet = btrExtAddField(hmPjf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadPPjfRecErr
        gBtrvErrorMsg ilRet, "mReadPPjfRec (btrExtAddField):" & "Pjf.Btr", CntrProj
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmPjf, tgPPjf1Rec(ilUpper).tPjf, ilExtLen, tgPPjf1Rec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadPPjfRecErr
            gBtrvErrorMsg ilRet, "mReadPPjfRec (btrExtGetNextExt):" & "Pjf.Btr", CntrProj
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tgPPjf1Rec(1).tPjf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmPjf, tgPPjf1Rec(ilUpper).tPjf, ilExtLen, tgPPjf1Rec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                ilRecOK = True
                If tgPPjf1Rec(ilUpper).tPjf.iAdfCode <> tmAdf.iCode Then
                    tmAdfSrchKey.iCode = tgPPjf1Rec(ilUpper).tPjf.iAdfCode 'ilCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If tgPPjf1Rec(ilUpper).tPjf.lPrfCode > 0 Then
                        If tgPPjf1Rec(ilUpper).tPjf.lPrfCode <> tmPrf.lCode Then
                            tmPrfSrchKey.lCode = tgPPjf1Rec(ilUpper).tPjf.lPrfCode 'ilCode
                            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRecOK = False
                        End If
                    Else
                        tmPrf.sName = ""
                        tmPrf.lCode = 0
                    End If
                Else
                    ilRecOK = False
                End If
                If ilRecOK Then
                    ilFound = False
                    For ilTest = LBONE To UBound(tgPPjf1Rec) - 1 Step 1
                        If (tgPPjf1Rec(ilTest).tPjf.iSofCode = tgPPjf1Rec(ilUpper).tPjf.iSofCode) And (tgPPjf1Rec(ilTest).tPjf.iAdfCode = tgPPjf1Rec(ilUpper).tPjf.iAdfCode) And (tgPPjf1Rec(ilTest).tPjf.lPrfCode = tgPPjf1Rec(ilUpper).tPjf.lPrfCode) And (tgPPjf1Rec(ilTest).tPjf.iVefCode = tgPPjf1Rec(ilUpper).tPjf.iVefCode) And (tgPPjf1Rec(ilTest).tPjf.lChfCode = tgPPjf1Rec(ilUpper).tPjf.lChfCode) Then
                            ilFound = True
                            If tgPPjf1Rec(ilTest).tPjf.iYear > tgPPjf1Rec(ilUpper).tPjf.iYear Then
                                'Swap records
                                tmPjf = tgPPjf1Rec(ilTest).tPjf
                                llRecPos = tgPPjf1Rec(ilTest).lRecPos
                                tgPPjf1Rec(ilTest).tPjf = tgPPjf1Rec(ilUpper).tPjf
                                tgPPjf1Rec(ilTest).lRecPos = tgPPjf1Rec(ilUpper).lRecPos
                                tgPPjf1Rec(ilUpper).tPjf = tmPjf
                                tgPPjf1Rec(ilUpper).lRecPos = llRecPos
                            End If
                            tgPPjf1Rec(ilTest).i2RecIndex = ilTest
                            tgPPjf2Rec(UBound(tgPPjf2Rec)).tPjf = tgPPjf1Rec(ilUpper).tPjf
                            tgPPjf2Rec(UBound(tgPPjf2Rec)).iStatus = 1
                            tgPPjf2Rec(UBound(tgPPjf2Rec)).lRecPos = tgPPjf1Rec(ilUpper).lRecPos
                            ReDim Preserve tgPPjf2Rec(0 To UBound(tgPPjf2Rec) + 1) As PJF2REC
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                            tgPPjf1Rec(ilUpper).sKey = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & tmPrf.sName
                            tgPPjf1Rec(ilUpper).sAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                        Else
                            tgPPjf1Rec(ilUpper).sKey = tmAdf.sName & tmPrf.sName
                            tgPPjf1Rec(ilUpper).sAdvtName = tmAdf.sName
                        End If
                        tgPPjf1Rec(ilUpper).sProdName = tmPrf.sName
                        tgPPjf1Rec(ilUpper).iStatus = 1
                        ilUpper = ilUpper + 1
                        ReDim Preserve tgPPjf1Rec(0 To ilUpper) As PJF1REC
                    End If
                End If
                ilRet = btrExtGetNext(hmPjf, tgPPjf1Rec(ilUpper).tPjf, ilExtLen, tgPPjf1Rec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmPjf, tgPPjf1Rec(ilUpper).tPjf, ilExtLen, tgPPjf1Rec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgPPjf1Rec(), 1), UBound(tgPPjf1Rec) - 1, 0, LenB(tgPPjf1Rec(1)), 0, LenB(tgPPjf1Rec(1).sKey), 0
    End If
    If imPrjNoYears > 1 Then
        For ilLoop = LBONE To UBound(tgPPjf1Rec) - 1 Step 1
            If tgPPjf1Rec(ilLoop).i2RecIndex <= 0 Then
                tgPPjf1Rec(ilLoop).i2RecIndex = UBound(tgPPjf2Rec)
                tgPPjf2Rec(UBound(tgPPjf2Rec)).tPjf.iYear = tgPPjf1Rec(ilLoop).tPjf.iYear + 1   'imCurYear + 1
                tgPPjf2Rec(UBound(tgPPjf2Rec)).iStatus = 0
                tgPPjf2Rec(UBound(tgPPjf2Rec)).lRecPos = 0
                ReDim Preserve tgPPjf2Rec(0 To UBound(tgPPjf2Rec) + 1) As PJF2REC
            End If
        Next ilLoop
    End If
    'mInitBudgetCtrls
    mReadPPjfRec = True
    Exit Function
mReadPPjfRecErr:
    On Error GoTo 0
    mReadPPjfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaleOfficePop                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSaleOfficePop()
    Dim ilRet As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilLoop As Integer
    'ilRet = gPopOfficeSourceBox(CntrProj, lbcSOffice, lbcSOfficeCode)
    ilRet = gPopOfficeSourceBox(CntrProj, lbcSOffice, tmSOfficeCode(), smSOfficeCodeTag)
    'ilRet = gPopUserVehComboBox(Budget, cbcCtrl, lbcSaleOfficeCode, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSaleOfficePopErr
        'gCPErrorMsg ilRet, "mSaleOfficePop (gPopUserVehComboBox: Vehicle/Combo)", Budget
        gCPErrorMsg ilRet, "mSaleOfficePop (gPopOfficeSourceBox: Vehicle)", CntrProj
        On Error GoTo 0
        lbcSOffice.AddItem "[New]", 0  'Force as first item on list
    End If
    'ReDim tmSaleOffice(1 To lbcSOfficeCode.ListCount + 1) As PJSALEOFFICE
    ReDim tmSaleOffice(0 To UBound(tmSOfficeCode) + 1) As PJSALEOFFICE
    For ilLoop = 0 To UBound(tmSOfficeCode) - 1 Step 1 'lbcSOfficeCode.ListCount - 1 Step 1
        slNameCode = tmSOfficeCode(ilLoop).sKey   'lbcSOfficeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", tmSaleOffice(ilLoop + 1).sName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmSaleOffice(ilLoop + 1).iCode = Val(slCode)
    Next ilLoop
    Exit Sub
mSaleOfficePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilRowNo As Integer
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilEffTime0 As Integer
    Dim ilEffTime1 As Integer
    Dim slStr As String
    Dim slMsg As String
    Dim ilPjf As Integer
    Dim ilCxfState As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilPjf2 As Integer
    Dim tlPjf As PJF
    Dim tlPjf1 As MOVEREC
    Dim tlPjf2 As MOVEREC
    slStr = Format$(gNow(), "m/d/yy")
    gPackDate slStr, ilEffDate0, ilEffDate1
    slStr = Format$(gNow(), "h:m:s AM/PM")
    gPackTime slStr, ilEffTime0, ilEffTime1
    mSetShow imBoxNo
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            mSaveRec = False
            imRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    ilRet = mMoveCtrlToRec()
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault    'Default
        Beep
        imBoxNo = ilRet
        ilRet = MsgBox("Field not specified", vbOKOnly + vbExclamation, "Save")
        'imTerminate = True
        mSaveRec = False
        Exit Function
    End If
    ilRet = mAnyZeros()
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault    'Default
        Beep
        imBoxNo = ilRet
        If ilRet = SOFFICEINDEX Then
            ilRet = MsgBox("Sales Office Field not specified", vbOKOnly + vbExclamation, "Save")
        Else
            ilRet = MsgBox("Advertiser Field not specified", vbOKOnly + vbExclamation, "Save")
        End If
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    ilRet = btrBeginTrans(hmPjf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("File in Use [Re-press Save], BeginTran" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
        'imTerminate = True
        mSaveRec = False
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilLoop = 0
    For ilPjf = LBONE To UBound(tgPjf1Rec) - 1 Step 1
        ilRowNo = tgPjf1Rec(ilPjf).iSaveIndex
        If tgPjf1Rec(ilPjf).iStatus = 1 Then
            tmCxfSrchKey.lCode = tgPjf1Rec(ilPjf).tPjf.lCxfChgR
            If tmCxfSrchKey.lCode <> 0 Then
                tmCxf.sComment = ""
                imCxfRecLen = Len(tmCxf) '5027
                ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], GetEqual Cxf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            Else
                tmCxf.lCode = 0
            End If
        End If
        ilCxfState = 0
        Do  'Loop until record updated or added
            tmCxf.sComment = ""
            tmCxf.sComType = "S"
            tmCxf.sShProp = "N"
            tmCxf.sShOrder = "N"
            tmCxf.sShSpot = "N"
            tmCxf.sShInv = "N"
            tmCxf.sShInsertion = "N"
            'tmCxf.iStrLen = Len(Trim$(smSave(8, ilRowNo)))
            tmCxf.sComment = Trim$(smSave(8, ilRowNo)) & Chr$(0) '& Chr$(0) 'sgTB
            If tgPjf1Rec(ilPjf).iStatus = 0 Then
                imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                'If Len(Trim$(tmCxf.sComment)) > 2 Then '2 so the control character at the end is not counted
                If Trim$(smSave(8, ilRowNo)) <> "" Then
                    ilCxfState = 1
                    tmCxf.lCode = 0 'Autoincrement
                    tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmCxf.lAutoCode = tmCxf.lCode
                    ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                Else
                    tmCxf.lCode = 0
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrInsert: Comment)"
            Else
                imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                'If Len(Trim$(tmCxf.sComment)) > 2 Then  '2 so the control character at end is not counted
                If Trim$(smSave(8, ilRowNo)) <> "" Then
                    If tmCxf.lCode = 0 Then
                        ilCxfState = 1
                        tmCxf.lCode = 0 'Autoincrement
                        tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                        tmCxf.lAutoCode = tmCxf.lCode
                        ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                    Else
                        ilCxfState = 2
                        tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                        gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                        gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                        ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                    End If
                Else
                    If tmCxf.lCode <> 0 Then
                        ilCxfState = 3
                        ilRet = btrDelete(hmCxf)
                    End If
                    tmCxf.lCode = 0
                End If
                slMsg = "mSaveRec (btrUpdate: Comment)"
            End If
            If ilRet = BTRV_ERR_CONFLICT Then
                tmCxfSrchKey.lCode = tgPjf1Rec(ilPjf).tPjf.lCxfChgR
                If tmCxfSrchKey.lCode <> 0 Then
                    tmCxf.sComment = ""
                    imCxfRecLen = Len(tmCxf) '5027
                    ilCRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilCRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmPjf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("File in Use [Re-press Save], GetEqual Cxf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        'imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Save], Update/Insert Cxf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
            'imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        tgPjf1Rec(ilPjf).tPjf.lCxfChgR = tmCxf.lCode
        'If tgSpf.sRemoteUsers = "Y" Then
            If ilCxfState = 1 Then
                Do
                    'tmCxfSrchKey.lCode = tmCxf.lCode
                    'imCxfRecLen = Len(tmCxf)
                    'ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    'slMsg = "mSaveRec (btrGetEqual:Personnel)"
                    'On Error GoTo mSaveRecErr
                    'gBtrvErrorMsg ilRet, slMsg, CntrProj
                    'On Error GoTo 0
                    tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmCxf.lAutoCode = tmCxf.lCode
                    tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                    gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                    imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                    ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                    slMsg = "mSaveRec (btrUpdate:Personnel)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], Update" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            ElseIf ilCxfState = 3 Then
'                If tgSpf.sRemoteUsers = "Y" Then
'                    tmDsf.lCode = 0
'                    tmDsf.sFileName = "CXF"
'                    gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                    gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                    tmDsf.iRemoteID = tmCxf.iRemoteID
'                    tmDsf.lAutoCode = tmCxf.lAutoCode
'                    tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'                    tmDsf.lCntrNo = 0
'                    ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        ilCRet = btrAbortTrans(hmPjf)
'                        Screen.MousePointer = vbDefault    'Default
'                        ilRet = MsgBox("File in Use [Re-press Save], Insert Dsf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
'                        'imTerminate = True
'                        mSaveRec = False
'                        Exit Function
'                    End If
'                End If
            End If
        'End If
        Do  'Loop until record updated or added
            If (tgPjf1Rec(ilPjf).iStatus = 0) Then  'New selected
                tgPjf1Rec(ilPjf).tPjf.lCode = 0
                tgPjf1Rec(ilPjf).tPjf.iEffDate(0) = ilEffDate0
                tgPjf1Rec(ilPjf).tPjf.iEffDate(1) = ilEffDate1
                tgPjf1Rec(ilPjf).tPjf.iEffTime(0) = ilEffTime0
                tgPjf1Rec(ilPjf).tPjf.iEffTime(1) = ilEffTime1
                tgPjf1Rec(ilPjf).tPjf.iRolloverDate(0) = 0
                tgPjf1Rec(ilPjf).tPjf.iRolloverDate(1) = 0
                tgPjf1Rec(ilPjf).tPjf.iUrfCode = tgUrf(0).iCode
                tgPjf1Rec(ilPjf).tPjf.iRemoteID = tgUrf(0).iRemoteUserID
                tgPjf1Rec(ilPjf).tPjf.lAutoCode = tgPjf1Rec(ilPjf).tPjf.lCode
                ilRet = btrInsert(hmPjf, tgPjf1Rec(ilPjf).tPjf, imPjfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Projection)"
            Else 'Old record-Update
                slMsg = "mSaveRec (btrGetDirect: Projection)"
                ilRet = btrGetDirect(hmPjf, tlPjf, imPjfRecLen, tgPjf1Rec(ilPjf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], GetDirect Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                'tmRec = tlPjf
                'ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
                'tlPjf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilCRet = btrAbortTrans(hmPjf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("File in Use [Re-press Save], GetByKey Pjf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
                '    'imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                LSet tlPjf1 = tlPjf
                LSet tlPjf2 = tgPjf1Rec(ilPjf).tPjf
                If StrComp(tlPjf1.sChar, tlPjf2.sChar, 0) <> 0 Then
                    tgPjf1Rec(ilPjf).tPjf.iEffDate(0) = ilEffDate0
                    tgPjf1Rec(ilPjf).tPjf.iEffDate(1) = ilEffDate1
                    tgPjf1Rec(ilPjf).tPjf.iEffTime(0) = ilEffTime0
                    tgPjf1Rec(ilPjf).tPjf.iEffTime(1) = ilEffTime1
                    tgPjf1Rec(ilPjf).tPjf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate slSyncDate, tgPjf1Rec(ilPjf).tPjf.iSyncDate(0), tgPjf1Rec(ilPjf).tPjf.iSyncDate(1)
                    gPackTime slSyncTime, tgPjf1Rec(ilPjf).tPjf.iSyncTime(0), tgPjf1Rec(ilPjf).tPjf.iSyncTime(1)
                    tgPjf1Rec(ilPjf).tPjf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hmPjf, tgPjf1Rec(ilPjf).tPjf, imPjfRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Projection)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Save], Update Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        If (tgPjf1Rec(ilPjf).iStatus = 0) Then  'And (tgSpf.sRemoteUsers = "Y") Then  'New selected
            Do
                'tmPjfSrchKey1.lCode = tgPjf1Rec(ilPjf).tPjf.lCode
                'ilRet = btrGetEqual(hmPjf, tgPjf1Rec(ilPjf).tPjf, imPjfRecLen, tmPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmPjf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                '    imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                tgPjf1Rec(ilPjf).tPjf.iRemoteID = tgUrf(0).iRemoteUserID
                tgPjf1Rec(ilPjf).tPjf.lAutoCode = tgPjf1Rec(ilPjf).tPjf.lCode
                tgPjf1Rec(ilPjf).tPjf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate slSyncDate, tgPjf1Rec(ilPjf).tPjf.iSyncDate(0), tgPjf1Rec(ilPjf).tPjf.iSyncDate(1)
                gPackTime slSyncTime, tgPjf1Rec(ilPjf).tPjf.iSyncTime(0), tgPjf1Rec(ilPjf).tPjf.iSyncTime(1)
                'ilRet = btrUpdate(hmPjf, tgPjf1Rec(ilPjf).tPjf, imPjfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPjf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("File in Use [Re-press Save], Update Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                'imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        If (tgPjf1Rec(ilPjf).iStatus = 0) Then  'New selected
            tgPjf1Rec(ilPjf).iStatus = 1
            ilRet = btrGetPosition(hmPjf, tgPjf1Rec(ilPjf).lRecPos)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPjf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("File in Use [Re-press Save], GetPosition Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                'imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        imSave(3, ilRowNo) = False
        ilPjf2 = tgPjf1Rec(ilPjf).i2RecIndex
        If ilPjf2 > 0 Then
            Do  'Loop until record updated or added
                If (tgPjf2Rec(ilPjf2).iStatus = 0) Then  'New selected
                    tgPjf2Rec(ilPjf2).tPjf.lCode = 0
                    tgPjf2Rec(ilPjf2).tPjf.iEffDate(0) = ilEffDate0
                    tgPjf2Rec(ilPjf2).tPjf.iEffDate(1) = ilEffDate1
                    tgPjf2Rec(ilPjf2).tPjf.iEffTime(0) = ilEffTime0
                    tgPjf2Rec(ilPjf2).tPjf.iEffTime(1) = ilEffTime1
                    tgPjf2Rec(ilPjf2).tPjf.iRolloverDate(0) = 0
                    tgPjf2Rec(ilPjf2).tPjf.iRolloverDate(1) = 0
                    tgPjf2Rec(ilPjf2).tPjf.iUrfCode = tgUrf(0).iCode
                    tgPjf2Rec(ilPjf2).tPjf.iRemoteID = tgUrf(0).iRemoteUserID
                    tgPjf2Rec(ilPjf2).tPjf.lAutoCode = tgPjf2Rec(ilPjf2).tPjf.lCode
                    ilRet = btrInsert(hmPjf, tgPjf2Rec(ilPjf2).tPjf, imPjfRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Projection)"
                Else 'Old record-Update
                    slMsg = "mSaveRec (btrGetDirect: Projection)"
                    ilRet = btrGetDirect(hmPjf, tlPjf, imPjfRecLen, tgPjf2Rec(ilPjf2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmPjf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("File in Use [Re-press Save], GetDirect Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        'imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    'tmRec = tlPjf
                    'ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
                    'tlPjf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilCRet = btrAbortTrans(hmPjf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("File in Use [Re-press Save], GetByKey Pjf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
                    '    'imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    LSet tlPjf1 = tlPjf
                    LSet tlPjf2 = tgPjf2Rec(ilPjf2).tPjf
                    If StrComp(tlPjf1.sChar, tlPjf2.sChar, 0) <> 0 Then
                        tgPjf2Rec(ilPjf2).tPjf.iEffDate(0) = ilEffDate0
                        tgPjf2Rec(ilPjf2).tPjf.iEffDate(1) = ilEffDate1
                        tgPjf2Rec(ilPjf2).tPjf.iEffTime(0) = ilEffTime0
                        tgPjf2Rec(ilPjf2).tPjf.iEffTime(1) = ilEffTime1
                        tgPjf2Rec(ilPjf2).tPjf.iUrfCode = tgUrf(0).iCode
                        tgPjf2Rec(ilPjf2).tPjf.iSourceID = tgUrf(0).iRemoteUserID
                        gPackDate slSyncDate, tgPjf2Rec(ilPjf2).tPjf.iSyncDate(0), tgPjf2Rec(ilPjf2).tPjf.iSyncDate(1)
                        gPackTime slSyncTime, tgPjf2Rec(ilPjf2).tPjf.iSyncTime(0), tgPjf2Rec(ilPjf2).tPjf.iSyncTime(1)
                        ilRet = btrUpdate(hmPjf, tgPjf2Rec(ilPjf2).tPjf, imPjfRecLen)
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                    slMsg = "mSaveRec (btrUpdate: Projection)"
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPjf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("File in Use [Re-press Save], Update/Insert Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                'imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            If (tgPjf2Rec(ilPjf2).iStatus = 0) Then 'And (tgSpf.sRemoteUsers = "Y") Then  'New selected
                Do
                    'tmPjfSrchKey1.lCode = tgPjf2Rec(ilPjf2).tPjf.lCode
                    'ilRet = btrGetEqual(hmPjf, tgPjf2Rec(ilPjf2).tPjf, imPjfRecLen, tmPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmPjf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    tgPjf2Rec(ilPjf2).tPjf.iRemoteID = tgUrf(0).iRemoteUserID
                    tgPjf2Rec(ilPjf2).tPjf.lAutoCode = tgPjf2Rec(ilPjf2).tPjf.lCode
                    tgPjf2Rec(ilPjf2).tPjf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate slSyncDate, tgPjf2Rec(ilPjf2).tPjf.iSyncDate(0), tgPjf2Rec(ilPjf2).tPjf.iSyncDate(1)
                    gPackTime slSyncTime, tgPjf2Rec(ilPjf2).tPjf.iSyncTime(0), tgPjf2Rec(ilPjf2).tPjf.iSyncTime(1)
                    'ilRet = btrUpdate(hmPjf, tgPjf2Rec(ilPjf2).tPjf, imPjfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], Update Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
            If (tgPjf2Rec(ilPjf2).iStatus = 0) Then  'New selected
                tgPjf2Rec(ilPjf2).iStatus = 1
                ilRet = btrGetPosition(hmPjf, tgPjf2Rec(ilPjf2).lRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], GetPosition Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
            End If
        End If
    Next ilPjf
    For ilPjf = LBONE To UBound(tgPjfDel) - 1 Step 1
        If tgPjfDel(ilPjf).iStatus = 1 Then
            Do
                tmCxfSrchKey.lCode = tgPjfDel(ilPjf).tPjf.lCxfChgR
                If tmCxfSrchKey.lCode <> 0 Then
                    tmCxf.sComment = ""
                    imCxfRecLen = Len(tmCxf) '5027
                    ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmPjf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("File in Use [Re-press Save], GetEqual Cxf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        'imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                Else
                    tmCxf.lCode = 0
                End If
                If tmCxf.lCode <> 0 Then
                    ilRet = btrDelete(hmCxf)
                End If
                slMsg = "mSaveRec (btrDelete: Comment)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPjf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("File in Use [Re-press Save], Delete Cxf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                'imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            If tmCxf.lCode > 0 Then
'                If tgSpf.sRemoteUsers = "Y" Then
'                    tmDsf.lCode = 0
'                    tmDsf.sFileName = "CXF"
'                    gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                    gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                    tmDsf.iRemoteID = tmCxf.iRemoteID
'                    tmDsf.lAutoCode = tmCxf.lAutoCode
'                    tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'                    tmDsf.lCntrNo = 0
'                    ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        ilCRet = btrAbortTrans(hmPjf)
'                        Screen.MousePointer = vbDefault    'Default
'                        ilRet = MsgBox("File in Use [Re-press Save], Insert Dsf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
'                        'imTerminate = True
'                        mSaveRec = False
'                        Exit Function
'                    End If
'                End If
            End If
            Do
                slMsg = "mSaveRec (btrGetDirect: Projection)"
                ilRet = btrGetDirect(hmPjf, tlPjf, imPjfRecLen, tgPjfDel(ilPjf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPjf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("File in Use [Re-press Save], GetDirect Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                    'imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                'tmRec = tlPjf
                'ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
                'tlPjf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilCRet = btrAbortTrans(hmPjf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("File in Use [Re-press Save], GetByKey Pjf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
                '    'imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                ilRet = btrDelete(hmPjf)
                slMsg = "mSaveRec (btrDelete: Projection)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPjf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("File in Use [Re-press Save], Delete Pjf" & Str$(ilRet), vbOKOnly + vbExclamation, "Save")
                'imTerminate = True
                mSaveRec = False
                Exit Function
            End If
'            If tgSpf.sRemoteUsers = "Y" Then
'                tmDsf.lCode = 0
'                tmDsf.sFileName = "PJF"
'                gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                tmDsf.iRemoteID = tlPjf.iRemoteID
'                tmDsf.lAutoCode = tlPjf.lAutoCode
'                tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'                tmDsf.lCntrNo = 0
'                ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'                If ilRet <> BTRV_ERR_NONE Then
'                    ilCRet = btrAbortTrans(hmPjf)
'                    Screen.MousePointer = vbDefault    'Default
'                    ilRet = MsgBox("File in Use [Re-press Save], Insert Dsf" & Str$(ilRet), vbOkOnly + vbExclamation, "Save")
'                    'imTerminate = True
'                    mSaveRec = False
'                    Exit Function
'                End If
'            End If
        End If
    Next ilPjf
    ilRet = btrEndTrans(hmPjf)
    ReDim tgPjfDel(0 To 1) As PJF2REC
    'mPopulate
    'gFindMatch slNameFac, 0, cbcSelect
    'If gLastFound(cbcSelect) > 0 Then
    '    cbcSelect.ListIndex = gLastFound(cbcSelect)
    'End If
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    'imTerminate = True
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
    Dim ilLoop As Integer
    Dim ilNew As Integer
    If imPjfChg And (UBound(tgPjf1Rec) > LBONE) Or (UBound(tgPjfDel) > LBONE) Then
        If ilAsk Then
            ilNew = True
            For ilLoop = LBONE To UBound(tgPjf1Rec) - 1 Step 1
                If tgPjf1Rec(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            For ilLoop = LBONE To UBound(tgPjfDel) - 1 Step 1
                If tgPjfDel(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            If Not ilNew Then
                slMess = "Save Changes to " & Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
            Else
                slMess = "Add Changes to " & Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
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
        'Case NAMEINDEX 'Name
        '    gSetChgFlag tmEnf.sName, edcName, tmCtrls(ilBoxNo)
        'Case GENREINDEX   'Genre
        '    gSetChgFlag smGenre, lbcGenre, tmCtrls(ilBoxNo)
        'Case COMMENTINDEX   'Comment
        '    gSetChgFlag smComment, edcComment, tmCtrls(ilBoxNo)
        'Case TIMEINDEX 'Time format
        '    If tmEnf.sTimeForm = "" Then
        '        gSetChgFlag tmEnf.sTimeForm, lbcTime, tmCtrls(ilBoxNo)
        '    Else
        '        slStr = lbcTime.List(Val(tmEnf.sTimeForm) - 1)
        '        gSetChgFlag slStr, lbcTime, tmCtrls(ilBoxNo)
        '    End If
        'Case LENINDEX 'Length format
        '    If tmEnf.sLenForm = "" Then
        '        gSetChgFlag tmEnf.sLenForm, lbcLen, tmCtrls(ilBoxNo)
        '    Else
        '        slStr = lbcLen.List(Val(tmEnf.sLenForm) - 1)
        '        gSetChgFlag slStr, lbcLen, tmCtrls(ilBoxNo)
        '    End If
        'Case PROGINDEX 'Selling or Airing or N/At
        '    If tmEnf.sPgmForm = "" Then
        '        gSetChgFlag tmEnf.sPgmForm, lbcProg, tmCtrls(ilBoxNo)
        '    Else
        '        slStr = lbcProg.List(Val(tmEnf.sPgmForm) - 1)
        '        gSetChgFlag slStr, lbcProg, tmCtrls(ilBoxNo)
        '    End If
        'Case SOURCEINDEX 'Source
        '    gSetChgFlag tmEnf.sPgmSource, edcSource, tmCtrls(ilBoxNo)
        'Case TYPEINDEX 'Type
        '    For ilLoop = 0 To 1 Step 1  'Set visibility
        '        gSetChgFlag tmEnf.sType(ilLoop), edcType(ilLoop), tmCtrls(ilBoxNo + ilLoop)
        '    Next ilLoop
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
    ilAltered = imPjfChg    'gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If ilAltered Then
        pbcProj.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        cbcSelect.Enabled = False
    Else
        If imSlspSelectedIndex < 0 Then
            pbcProj.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cbcSelect.Enabled = True
        Else
            pbcProj.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            cbcSelect.Enabled = True
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And ((UBound(tgPjf1Rec) > 1) Or (UBound(tgPjfDel) > LBONE)) Then
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
        cmcRollover.Enabled = False
    Else
        cmcUndo.Enabled = False
        If (tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0) Then
            cmcRollover.Enabled = False
        Else
            If imUpdateAllowed Then
                cmcRollover.Enabled = True
            Else
                cmcRollover.Enabled = False
            End If
        End If
    End If
    If (tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0) Then
        cmcBlock.Enabled = False
    Else
        If imUpdateAllowed Then
            cmcBlock.Enabled = True
        Else
            cmcBlock.Enabled = False
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
        pbcClickFocus.SetFocus
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SOFFICEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ADVTINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PROPNOINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DEMOINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case VEHINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case NRINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case POTINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMMENTINDEX
            edcComment.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetPrice                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from working area  *
'*                      to the record                  *
'*                                                     *
'*******************************************************
Private Sub mSetPrice(ilGroup As Integer, ilRowNo As Integer, slInDollar As String)
    Dim ilWk As Integer
    Dim slStart As String
    Dim ilPjf As Integer
    Dim ilPjf1 As Integer
    'Dim slNoWks As String
    'Dim slNoWks1 As String
    'Dim slAvgDollar As String
    'Dim slEndDollar As String
    Dim llAvgDollar As Long
    Dim llEndDollar As Long
    Dim llDollar As Long
    Dim llInDollar As Long
    llInDollar = Val(slInDollar)
    ilPjf = imSave(1, ilRowNo)
    ilPjf1 = ilPjf
    If tgPjf1Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
        If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
            llAvgDollar = llInDollar / tmPdGroups(ilGroup).iTrueNoWks
            llEndDollar = llInDollar - llAvgDollar * (tmPdGroups(ilGroup).iTrueNoWks - 1)
            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                llDollar = llAvgDollar
                If ilWk = tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Then
                    llDollar = llEndDollar
                End If
                slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                If rbcShow(0).Value Then
                    slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                    mPjfSetGross ilPjf1, slStart, llDollar, tgPjf1Rec(), tgPjf2Rec()
                Else
                    tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                End If
                'If ilWk = 1 Then
                '    If rbcShow(0).Value Then    'If input by corporate, then don't split
                '        tgPjf1Rec(ilPjf).tPjf.lGross(0) = 0
                '        tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                '    Else
                '        slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                '        slStart = gObtainStartCorp(slDate, True)
                '        ilDay = gWeekDayStr(slStart)
                '        'If ilDay = 0 Then
                '            tgPjf1Rec(ilPjf).tPjf.lGross(0) = 0
                '            tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                '        'Else
                '        '    tgPjf1Rec(ilPjf).tPjf.lGross(0) = (llDollar * ilDay) / 7
                '        '    tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar - tgPjf1Rec(ilPjf).tPjf.lGross(0)
                '        'End If
                '    End If
                'ElseIf ilWk = 52 Then
                '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
                '        tgPjf1Rec(ilPjf).tPjf.lGross(53) = 0
                '        tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                '    Else
                '        slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                '        slStart = gObtainEndCorp(slDate, True)
                '        ilDay = gWeekDayStr(slStart)
                '        'If ilDay = 6 Then
                '            tgPjf1Rec(ilPjf).tPjf.lGross(53) = 0
                '            tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                '        'Else
                '        '    ilDay = 7 - ilDay - 1
                '        '    tgPjf1Rec(ilPjf).tPjf.lGross(52) = (llDollar * ilDay) / 7
                '        '    tgPjf1Rec(ilPjf).tPjf.lGross(53) = llDollar - tgPjf1Rec(ilPjf).tPjf.lGross(52)
                '        'End If
                '    End If
                'Else
                '    tgPjf1Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                'End If
            Next ilWk
        End If
    Else
        ilPjf = tgPjf1Rec(ilPjf).i2RecIndex
        If ilPjf > 0 Then
            If tgPjf2Rec(ilPjf).tPjf.iYear = tmPdGroups(ilGroup).iYear Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    llAvgDollar = llInDollar / tmPdGroups(ilGroup).iTrueNoWks
                    llEndDollar = llInDollar - llAvgDollar * (tmPdGroups(ilGroup).iTrueNoWks - 1)
                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                        llDollar = llAvgDollar
                        If ilWk = tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Then
                            llDollar = llEndDollar
                        End If
                        slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                        If rbcShow(0).Value Then
                            slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                            mPjfSetGross ilPjf1, slStart, llDollar, tgPjf1Rec(), tgPjf2Rec()
                        Else
                            tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        End If
                        'If ilWk = 1 Then
                        '    If rbcShow(0).Value Then    'If input by corporate, then don't split
                        '        tgPjf2Rec(ilPjf).tPjf.lGross(0) = 0
                        '        tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        '    Else
                        '        slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                        '        slStart = gObtainStartCorp(slDate, True)
                        '        ilDay = gWeekDayStr(slStart)
                        '        'If ilDay = 0 Then
                        '            tgPjf2Rec(ilPjf).tPjf.lGross(0) = 0
                        '            tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        '        'Else
                        '        '    tgPjf2Rec(ilPjf).tPjf.lGross(0) = (llDollar * ilDay) / 7
                        '        '    tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar - tgPjf2Rec(ilPjf).tPjf.lGross(0)
                        '        'End If
                        '    End If
                        'ElseIf ilWk = 52 Then
                        '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
                        '        tgPjf2Rec(ilPjf).tPjf.lGross(53) = 0
                        '        tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        '    Else
                        '        slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                        '        slStart = gObtainEndCorp(slDate, True)
                        '        ilDay = gWeekDayStr(slStart)
                        '        'If ilDay = 6 Then
                        '            tgPjf2Rec(ilPjf).tPjf.lGross(53) = 0
                        '            tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        '        'Else
                        '        '    ilDay = 7 - ilDay - 1
                        '        '    tgPjf2Rec(ilPjf).tPjf.lGross(52) = (llDollar * ilDay) / 7
                        '        '    tgPjf2Rec(ilPjf).tPjf.lGross(53) = llDollar - tgPjf2Rec(ilPjf).tPjf.lGross(52)
                        '        'End If
                        '    End If
                        'Else
                        '    tgPjf2Rec(ilPjf).tPjf.lGross(ilWk) = llDollar
                        'End If
                    Next ilWk
                End If
            End If
        End If
    End If
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
    Dim slAns As String
    Dim slDollar As String
    Dim llDollar As Long
    Dim llChfCode As Long
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SOFFICEINDEX
            lbcSOffice.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcSOffice.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcSOffice.List(lbcSOffice.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(2, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(2, imRowNo) = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case ADVTINDEX
            lbcAdvt.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcAdvt.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcAdvt.List(lbcAdvt.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(3, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(3, imRowNo) = slStr
                tgPjf1Rec(imSave(1, imRowNo)).sAdvtName = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
                mGetShowPrices
                mGetShowOPrices
                mGetShowPPrices
                pbcProj_Paint
            End If
        Case PRODINDEX
            lbcProd.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            'If lbcProd.ListIndex <= 0 Then
            '    slStr = ""
            'Else
            '    slStr = lbcProd.List(lbcProd.ListIndex)
            'End If
            slStr = edcDropDown.Text
            If slStr = "[None]" Then
                slStr = ""
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(4, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(4, imRowNo) = slStr
                tgPjf1Rec(imSave(1, imRowNo)).sProdName = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
                mGetShowPrices
                mGetShowOPrices
                mGetShowPPrices
                pbcProj_Paint
            End If
        Case PROPNOINDEX
            lbcPropNo.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcPropNo.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcPropNo.List(lbcPropNo.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(1, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(1, imRowNo) = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
                lmSave(5, imRowNo) = mGetChfCode(smSave(1, imRowNo))
                Screen.MousePointer = vbHourglass
                llChfCode = lmSave(5, imRowNo)
                mGenProjForPropNo llChfCode, True
                Screen.MousePointer = vbDefault
            End If
        Case DEMOINDEX
            lbcDemo.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcDemo.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcDemo.List(lbcDemo.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(5, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(5, imRowNo) = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case VEHINDEX
            lbcVehicle.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcVehicle.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcVehicle.List(lbcVehicle.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(6, imRowNo) <> slStr Then
                imSave(3, imRowNo) = True
                smSave(6, imRowNo) = slStr
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case NRINDEX
            lbcNR.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcNR.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcNR.List(lbcNR.ListIndex)
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If lbcNR.ListIndex <> imSave(2, imRowNo) Then
                imSave(3, imRowNo) = True
                imSave(2, imRowNo) = lbcNR.ListIndex
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case POTINDEX
            lbcPot.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            If lbcPot.ListIndex <= 0 Then
                slStr = ""
                slAns = ""
            'ElseIf lbcPot.ListIndex = 1 Then
            '    slStr = ""
            '    slAns = lbcPot.List(lbcPot.ListIndex)
            Else
                slStr = lbcPot.List(lbcPot.ListIndex)
                slAns = slStr
            End If
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(7, imRowNo) <> slAns Then
                imSave(3, imRowNo) = True
                smSave(7, imRowNo) = slAns
                If imRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case COMMENTINDEX
            edcComment.Visible = False  'Set visibility
            slStr = edcComment.Text
            gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(8, imRowNo) <> edcComment.Text Then
                smSave(8, imRowNo) = edcComment.Text
                If imRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imPjfChg = True
                End If
            End If
        Case PD1INDEX, PD2INDEX, PD3INDEX
            edcDropDown.Visible = False
            slDollar = edcDropDown.Text
            llDollar = Val(slDollar)
            If lmSave(ilBoxNo - PD1INDEX + 1, imRowNo) <> llDollar Then
                    imSave(3, imRowNo) = True
                'Recompute total and set weeks
                'If tmPdGroups(1).iYear = tmPdGroups(ilBoxNo - PD1INDEX + 1).iYear Then
                    'Remove old values
                    lmSave(4, imRowNo) = lmSave(4, imRowNo) - lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)
                    lmCTSave(ilBoxNo - PD1INDEX + 1) = lmCTSave(ilBoxNo - PD1INDEX + 1) - lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)
                    lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) - lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)
                    lmCAPSave(ilBoxNo - PD1INDEX + 1) = lmCAPSave(ilBoxNo - PD1INDEX + 1) - lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)
                    lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) - lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)
                    'Set new values into fields
                    lmSave(ilBoxNo - PD1INDEX + 1, imRowNo) = llDollar
                    lmSave(4, imRowNo) = lmSave(4, imRowNo) + llDollar
                    slStr = Trim$(Str$(lmSave(ilBoxNo - PD1INDEX + 1, imRowNo)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow

                    slStr = Trim$(Str$(lmSave(4, imRowNo)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcProj, slStr, tmCtrls(ilBoxNo)
                    smShow(TOTALINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow

                    lmCTSave(ilBoxNo - PD1INDEX + 1) = lmCTSave(ilBoxNo - PD1INDEX + 1) + llDollar
                    'slStr = Trim$(Str$(lmGTSave(ilBoxNo - PD1INDEX + 1)))
                    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    'gSetShow pbcProj, slStr, tmGTCtrls(1)
                    'smGTShow(ilBoxNo - PD1INDEX + 1) = tmGTCtrls(1).sShow
                    lmCTSave(GTTOTALINDEX) = lmCTSave(GTTOTALINDEX) + llDollar
                    'slStr = Trim$(Str$(lmGTSave(GTTOTALINDEX)))
                    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    'gSetShow pbcProj, slStr, tmGTCtrls(GTTOTALINDEX)
                    'smGTShow(GTTOTALINDEX) = tmGTCtrls(GTTOTALINDEX).sShow
                    lmCAPSave(ilBoxNo - PD1INDEX + 1) = lmCAPSave(ilBoxNo - PD1INDEX + 1) + llDollar
                    lmCAPSave(GTTOTALINDEX) = lmCAPSave(GTTOTALINDEX) + llDollar
                'End If
                slStr = edcDropDown.Text
                mSetPrice ilBoxNo - PD1INDEX + 1, imRowNo, slStr
                imPjfChg = True
                'pbcProj.Cls
                pbcProj_Paint
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSOfficeBranch                  *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      office and process             *
'*                      communication back from sales  *
'*                      office                         *
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
Private Function mSOfficeBranch() As Integer
'
'   ilRet = mSSourceBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim ilPos As Integer
    Dim ilSvPos As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcSOffice, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSOfficeBranch = False
        Exit Function
    End If
    If igWinStatus(SALESOFFICESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mSOfficeBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(SALESOFFICESLIST)) Then
    '    imDoubleClickName = False
    '    mSOfficeBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igSofCallSource = CALLSOURCESALESPERSON
    If edcDropDown.Text = "[New]" Then
        sgSofName = ""
    Else
        ilSvPos = 0
        ilPos = 1
        Do While ilPos > 0
            ilPos = InStr(ilPos, slStr, "/")
            If ilPos = 0 Then
                Exit Do
            End If
            ilSvPos = ilPos
            ilPos = ilPos + 1
        Loop
        If ilSvPos > 0 Then
            sgSofName = Left$(slStr, ilSvPos - 1)
        Else
            sgSofName = ""
        End If
        'ilRet = gParseItem(slStr, 1, "/", sgSofName)
        'If ilRet <> CP_MSG_NONE Then
        '    sgSofName = slStr
        'End If
    End If
    ilUpdateAllowed = imUpdateAllowed

'    CntrProj.Enabled = False
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CntrProj^Test\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
        Else
            slStr = "CntrProj^Prod\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CntrProj^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
    '    Else
    '        slStr = "CntrProj^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "SOffice.Exe " & slStr, 1)
    'CntrProj.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    SOffice.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgSofName)
    igSofCallSource = Val(sgSofName)
    ilParse = gParseItem(slStr, 2, "\", sgSofName)
    'CntrProj.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSOfficeBranch = True
    imUpdateAllowed = ilUpdateAllowed
'    gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igSofCallSource = CALLDONE Then  'Done
        igSofCallSource = CALLNONE
'        gSetMenuState True
        lbcSOffice.Clear
        smSOfficeCodeTag = ""
        mSaleOfficePop
        If imTerminate Then
            mSOfficeBranch = False
            Exit Function
        End If
        gFindMatch sgSofName, 1, lbcSOffice
        sgSofName = ""
        If gLastFound(lbcSOffice) > 0 Then
            imChgMode = True
            lbcSOffice.ListIndex = gLastFound(lbcSOffice)
            edcDropDown.Text = lbcSOffice.List(lbcSOffice.ListIndex)
            imChgMode = False
            mSOfficeBranch = False
            mSetChg SOFFICEINDEX
        Else
            imChgMode = True
            lbcSOffice.ListIndex = 0
            edcDropDown.Text = lbcSOffice.List(0)
            imChgMode = False
            mSetChg SOFFICEINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igSofCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igSofCallSource = CALLNONE
        sgSofName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igSofCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igSofCallSource = CALLNONE
        sgSofName = ""
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
    'Unload IconTraf
    igManUnload = YES
    'Unload Traffic
    Unload CntrProj
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestCorpYear                   *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test corporate year            *
'*                                                     *
'*******************************************************
Private Sub mTestCorpYear(ilInYear As Integer)
    Dim ilYear As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    ilYear = ilInYear
    If ilYear < 100 Then
        If ilYear >= 70 Then
            ilYear = 1900 + ilYear
        Else
            ilYear = 2000 + ilYear
        End If
    End If
    If tgSpf.sRUseCorpCal = "Y" Then
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilYear Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            MsgBox "Corporate Year Missing for" & Str$(ilYear), vbOKOnly + vbExclamation, "Rate Card"
            imIgnoreSetting = True
            rbcShow(1).Value = True
            rbcShow(0).Enabled = False
        Else
            rbcShow(0).Enabled = True
        End If
    End If
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
    Dim ilRowNo As Integer
    For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        If smSave(2, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If smSave(3, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If smSave(6, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If imSave(2, ilRowNo) < 0 Then
            mTestFields = NO
            Exit Function
        End If
        If smSave(7, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
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
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim llChfCode As Long
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilPjf As Integer
    gFindMatch smSave(2, ilRowNo), 1, lbcSOffice
    If (smSave(2, ilRowNo) = "") Or (gLastFound(lbcSOffice) <= 0) Then
        Beep
        ilRes = MsgBox("Sales Office must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = SOFFICEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    gFindMatch smSave(3, ilRowNo), 1, lbcAdvt
    If (smSave(3, ilRowNo) = "") Or (gLastFound(lbcAdvt) <= 0) Then
        Beep
        ilRes = MsgBox("Advertiser must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ADVTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If smSave(1, ilRowNo) <> "" Then
        mPropNoPop ilRowNo
        llChfCode = mGetChfCode(smSave(1, ilRowNo))
        gFindMatch smSave(3, ilRowNo), 1, lbcAdvt
        If gLastFound(lbcAdvt) > 0 Then
            slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) <> tmChf.iAdfCode Then
                Beep
                ilRes = MsgBox("Specified Advertiser and Proposal Advertiser Must Match", vbOKOnly + vbExclamation, "Incomplete")
                imBoxNo = ADVTINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If
    gFindMatch smSave(6, ilRowNo), 0, lbcVehicle
    If (smSave(6, ilRowNo) = "") Or (gLastFound(lbcVehicle) < 0) Then
        Beep
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = VEHINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If imSave(2, ilRowNo) < 0 Then
        Beep
        ilRes = MsgBox("New/Return must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = NRINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    gFindMatch smSave(7, ilRowNo), 1, lbcPot
    If (smSave(7, ilRowNo) = "") Or (gLastFound(lbcPot) <= 0) Then
        Beep
        ilRes = MsgBox("Potential must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = POTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    ilPjf = imSave(1, ilRowNo)
    If imSave(3, ilRowNo) And (Len(Trim$(smSave(8, ilRowNo))) = 0) And (tgPjf1Rec(ilPjf).iStatus = 1) Then
        Beep
        ilRes = MsgBox("Comment must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = COMMENTINDEX
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
    Dim slCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim ilLoop As Integer
    'ilRet = gPopUserVehicleBox(CntrProj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(CntrProj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    'ilRet = gPopUserVehComboBox(Budget, cbcCtrl, Traffic!lbcUserVehicle, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gPopUserVehComboBox: Vehicle/Combo)", Budget
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", CntrProj
        On Error GoTo 0
    End If
    'ReDim tmUserVeh(1 To Traffic!lbcUserVehicle.ListCount + 1) As USERVEH
    ReDim tmUserVeh(0 To UBound(tgUserVehicle) + 1) As USERVEH
    For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", tmUserVeh(ilLoop + 1).sName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmUserVeh(ilLoop + 1).iCode = Val(slCode)
    Next ilLoop
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
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcProj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcProj_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer
    Dim ilWk As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    If Button = 2 Then
        Exit Sub
    End If
    'Check if hot spot
    If imInHotSpot Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If (UBound(tgPjf1Rec) > 1) Or ((UBound(tgPjf1Rec) = 1) And (smSave(2, 1) <> "") And (smSave(3, 1) <> "") And (smSave(6, 1) <> "")) Then
        For ilLoop = LBONE To UBound(imHotSpot, 1) Step 1
            If (X >= imHotSpot(ilLoop, 1)) And (X <= imHotSpot(ilLoop, 3)) And (Y >= imHotSpot(ilLoop, 2)) And (Y <= imHotSpot(ilLoop, 4)) Then
                Screen.MousePointer = vbHourglass
                mSetShow imBoxNo
                imBoxNo = -1
                imInHotSpot = True
                Select Case ilLoop
                    Case 1  'Goto Start
                        imPdYear = imPrjStartYear
                        imPdStartWk = imPrjStartWk
                    Case 2  'Reduce by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(1).iYear = imPrjStartYear) And (tmPdGroups(1).iStartWkNo <= imPrjStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 9 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                If (tmPdGroups(1).iStartWkNo >= 12) And (tmPdGroups(1).iStartWkNo <= 14) Then
                                    imPdStartWk = ilStartWk(1)
                                ElseIf (tmPdGroups(1).iStartWkNo >= 25) And (tmPdGroups(1).iStartWkNo <= 27) Then
                                    imPdStartWk = ilStartWk(1)  'Compute start of second quarter
                                    For ilWk = 1 To 3 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                Else
                                    imPdStartWk = ilStartWk(1)  'Compute start of third quarter
                                    For ilWk = 1 To 6 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                End If
                                If (imPdYear = imPrjStartYear) And (imPdStartWk < imPrjStartWk) Then
                                    imPdStartWk = imPrjStartWk
                                End If
                            End If
                        ElseIf rbcType(1).Value Then    'Month
                            If (tmPdGroups(1).iYear = imPrjStartYear) And (tmPdGroups(1).iStartWkNo <= imPrjStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 11 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 2 To 12 Step 1
                                    If tmPdGroups(1).iStartWkNo = ilStartWk(ilWk) Then
                                        Exit For
                                    End If
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk - 1)
                                Next ilWk
                            End If
                            If (imPdYear = imPrjStartYear) And (imPdStartWk < imPrjStartWk) Then
                                imPdStartWk = imPrjStartWk
                            End If
                        ElseIf rbcType(2).Value Then    'Week
                            If (tmPdGroups(1).iYear = imPrjStartYear) And (tmPdGroups(1).iStartWkNo <= imPrjStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 12 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                                imPdStartWk = imPdStartWk - 1
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                imPdStartWk = tmPdGroups(1).iStartWkNo - 1
                            End If
                            If (imPdYear = imPrjStartYear) And (imPdStartWk < imPrjStartWk) Then
                                imPdStartWk = imPrjStartWk
                            End If
                        End If
                    Case 3  'Increase by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(3).iYear = imPrjStartYear + imPrjNoYears - 1) And (tmPdGroups(3).iStartWkNo > 39) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths tmPdGroups(3).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(3).iYear = imPrjStartYear + imPrjNoYears - 1) And (tmPdGroups(3).iStartWkNo >= ilStartWk(12)) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths tmPdGroups(3).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(3).iYear = imPrjStartYear + imPrjNoYears - 1) And (tmPdGroups(3).iStartWkNo >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        End If
                    Case 4  'GoTo End
                        imPdYear = imPrjStartYear + imPrjNoYears - 1
                        If rbcType(0).Value Then    'Quarter
                            imPdStartWk = 1
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(10)  'At end
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(12) + ilNoWks(12) - 3
                        End If
                End Select
                pbcProj.Cls
                mGetShowDates True
                pbcProj_Paint
                Screen.MousePointer = vbDefault
                imInHotSpot = False
                Exit Sub
            End If
        Next ilLoop
    End If
    Screen.MousePointer = vbDefault
    ilCompRow = vbcProj.LargeChange + 1
    If UBound(tgPjf1Rec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgPjf1Rec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcProj.Value - 1
                    If (ilRowNo < LBONE) Or (ilRowNo > UBound(smSave, 2)) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If ilBox > 2 Then
                        If smSave(3, ilRowNo) = "" Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    End If
                    mSetShow imBoxNo
                    imRowNo = ilRow + vbcProj.Value - 1
                    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
                        Beep
                        Exit Sub
                    End If
                    If (imRowNo = UBound(smSave, 2)) And (imSave(1, imRowNo) <= 0) Then
                        mInitNewProj
                    End If
                    mGetShowPrices
                    mGetShowOPrices
                    mGetShowPPrices
                    pbcProj_Paint
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcProj_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintProjTitle
    llColor = pbcProj.ForeColor
    slFontName = pbcProj.FontName
    flFontSize = pbcProj.FontSize
    pbcProj.ForeColor = BLUE
    pbcProj.FontBold = False
    pbcProj.FontSize = 7
    pbcProj.FontName = "Arial"
    pbcProj.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = LBONE To UBound(tmWKCtrls) Step 1
        'gPaintArea pbcProj, tmWKCtrls(ilBox).fBoxX, tmWKCtrls(ilBox).fBoxY, tmWKCtrls(ilBox).fBoxW - 15, tmWKCtrls(ilBox).fBoxH - 15, WHITE
        pbcProj.CurrentX = tmWKCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcProj.CurrentY = tmWKCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcProj.Print tmWKCtrls(ilBox).sShow
    Next ilBox
    For ilBox = LBONE To UBound(tmNWCtrls) Step 1
        'gPaintArea pbcProj, tmNWCtrls(ilBox).fBoxX, tmNWCtrls(ilBox).fBoxY, tmNWCtrls(ilBox).fBoxW - 15, tmNWCtrls(ilBox).fBoxH - 15, LIGHTBLUE
        pbcProj.CurrentX = tmNWCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcProj.CurrentY = tmNWCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcProj.Print tmNWCtrls(ilBox).sShow
    Next ilBox
    pbcProj.FontSize = flFontSize
    pbcProj.FontName = slFontName
    pbcProj.FontSize = flFontSize
    pbcProj.ForeColor = llColor
    pbcProj.FontBold = True
    ilStartRow = vbcProj.Value '+ 1  'Top location
    ilEndRow = vbcProj.Value + vbcProj.LargeChange ' + 1
    If ilEndRow > UBound(smSave, 2) Then
        If smSave(2, UBound(smSave, 2)) <> "" Then
            ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smSave, 2) - 1
        End If
    End If
    llColor = pbcProj.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smSave, 2) Then
            pbcProj.ForeColor = DARKPURPLE
        Else
            pbcProj.ForeColor = llColor
        End If
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcProj.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcProj.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = smShow(ilBox, ilRow)
            pbcProj.Print slStr
            'pbcProj.ForeColor = llColor
        Next ilBox
        pbcProj.ForeColor = llColor
    Next ilRow
    If imSlspSelectedIndex >= 0 Then
        For ilBox = LBONE To UBound(tmCAPCtrls) Step 1
            If imState = 0 Then
                slStr = Trim$(Str$(lmCAPSave(ilBox)))
            Else
                slStr = Trim$(Str$(lmCAPSave(ilBox) - lmOAPSave(ilBox)))
            End If
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcProj, slStr, tmCAPCtrls(ilBox)
            gPaintArea pbcProj, tmCAPCtrls(ilBox).fBoxX, tmCAPCtrls(ilBox).fBoxY, tmCAPCtrls(ilBox).fBoxW - 15, tmCAPCtrls(ilBox).fBoxH - 15, WHITE
            pbcProj.CurrentX = tmCAPCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcProj.CurrentY = tmCAPCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
            pbcProj.Print tmCAPCtrls(ilBox).sShow
        Next ilBox
        For ilBox = LBONE To UBound(tmPAPCtrls) Step 1
            If imState = 0 Then
                slStr = Trim$(Str$(lmPAPSave(ilBox)))
            Else
                slStr = Trim$(Str$(lmCAPSave(ilBox) - lmPAPSave(ilBox)))
            End If
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcProj, slStr, tmPAPCtrls(ilBox)
            gPaintArea pbcProj, tmPAPCtrls(ilBox).fBoxX, tmPAPCtrls(ilBox).fBoxY, tmPAPCtrls(ilBox).fBoxW - 15, tmPAPCtrls(ilBox).fBoxH - 15, WHITE
            pbcProj.CurrentX = tmPAPCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcProj.CurrentY = tmPAPCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
            pbcProj.Print tmPAPCtrls(ilBox).sShow
        Next ilBox
        For ilBox = LBONE To UBound(tmCTCtrls) Step 1
            If imState = 0 Then
                slStr = Trim$(Str$(lmCTSave(ilBox)))
            Else
                slStr = Trim$(Str$(lmCTSave(ilBox) - lmOTSave(ilBox)))
            End If
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcProj, slStr, tmCTCtrls(ilBox)
            gPaintArea pbcProj, tmCTCtrls(ilBox).fBoxX, tmCTCtrls(ilBox).fBoxY, tmCTCtrls(ilBox).fBoxW - 15, tmCTCtrls(ilBox).fBoxH - 15, WHITE
            pbcProj.CurrentX = tmCTCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcProj.CurrentY = tmCTCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
            pbcProj.Print tmCTCtrls(ilBox).sShow
        Next ilBox
        For ilBox = LBONE To UBound(tmPTCtrls) Step 1
            If imState = 0 Then
                slStr = Trim$(Str$(lmPTSave(ilBox)))
            Else
                slStr = Trim$(Str$(lmCTSave(ilBox) - lmPTSave(ilBox)))
            End If
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcProj, slStr, tmPTCtrls(ilBox)
            gPaintArea pbcProj, tmPTCtrls(ilBox).fBoxX, tmPTCtrls(ilBox).fBoxY, tmPTCtrls(ilBox).fBoxW - 15, tmPTCtrls(ilBox).fBoxH - 15, WHITE
            pbcProj.CurrentX = tmPTCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcProj.CurrentY = tmPTCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
            pbcProj.Print tmPTCtrls(ilBox).sShow
        Next ilBox
    Else
        For ilBox = LBONE To UBound(tmCAPCtrls) Step 1
            gPaintArea pbcProj, tmCAPCtrls(ilBox).fBoxX, tmCAPCtrls(ilBox).fBoxY, tmCAPCtrls(ilBox).fBoxW - 15, tmCAPCtrls(ilBox).fBoxH - 15, WHITE
        Next ilBox
        For ilBox = LBONE To UBound(tmPAPCtrls) Step 1
            gPaintArea pbcProj, tmPAPCtrls(ilBox).fBoxX, tmPAPCtrls(ilBox).fBoxY, tmPAPCtrls(ilBox).fBoxW - 15, tmPAPCtrls(ilBox).fBoxH - 15, WHITE
        Next ilBox
        For ilBox = LBONE To UBound(tmCTCtrls) Step 1
            gPaintArea pbcProj, tmCTCtrls(ilBox).fBoxX, tmCTCtrls(ilBox).fBoxY, tmCTCtrls(ilBox).fBoxW - 15, tmCTCtrls(ilBox).fBoxH - 15, WHITE
        Next ilBox
        For ilBox = LBONE To UBound(tmPTCtrls) Step 1
            gPaintArea pbcProj, tmPTCtrls(ilBox).fBoxX, tmPTCtrls(ilBox).fBoxY, tmPTCtrls(ilBox).fBoxW - 15, tmPTCtrls(ilBox).fBoxH - 15, WHITE
        Next ilBox
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    If imBoxNo = SOFFICEINDEX Then
        If mSOfficeBranch() Then
            Exit Sub
        End If
    End If
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
    If imBoxNo = POTINDEX Then
        If mPotBranch() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smSave, 2) = 1) And (imSave(1, 1) = 0) Then
                imTabDirection = 0  'Set-Left to right
                imRowNo = 1
                mInitNewProj
                mGetShowPrices
                mGetShowOPrices
                mGetShowPPrices
                pbcProj_Paint
            Else
                If UBound(smSave, 2) <= vbcProj.LargeChange Then 'was <=
                    vbcProj.Max = LBONE 'LBound(smSave, 2)
                Else
                    vbcProj.Max = UBound(smSave, 2) - vbcProj.LargeChange '- 1
                End If
                imRowNo = 1
                If imRowNo >= UBound(smSave, 2) Then
                    mInitNewProj
                End If
                imSettingValue = True
                vbcProj.Value = vbcProj.Min
                imSettingValue = False
                mGetShowPrices
                mGetShowOPrices
                mGetShowPPrices
                pbcProj_Paint
            End If
            ilBox = SOFFICEINDEX 'PROPNOINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case SOFFICEINDEX, 0
            mSetShow imBoxNo
            If (imBoxNo < 1) And (imRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = PD3INDEX    'COMMENTINDEX
            If imRowNo <= 1 Then
                imBoxNo = -1
                imRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imRowNo = imRowNo - 1
            If imRowNo < vbcProj.Value Then
                imSettingValue = True
                vbcProj.Value = vbcProj.Value - 1
                imSettingValue = False
            End If
            mGetShowPrices
            mGetShowOPrices
            mGetShowPPrices
            pbcProj_Paint
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case ADVTINDEX
            mSetShow imBoxNo
            If smSave(1, imRowNo) <> "" Then
                llChfCode = mGetChfCode(smSave(1, imRowNo))
                gFindMatch smSave(3, imRowNo), 1, lbcAdvt
                If gLastFound(lbcAdvt) > 0 Then
                    slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> tmChf.iAdfCode Then
                        Beep
                        imBoxNo = ADVTINDEX
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            End If
            ilBox = imBoxNo - 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case PROPNOINDEX
            If (lbcPropNo.ListIndex > 0) And (smSave(3, imRowNo) <> "") Then 'Test that advertiser name match
                slStr = lbcPropNo.List(lbcPropNo.ListIndex)
                If smSave(1, imRowNo) <> slStr Then
                    lmSave(5, imRowNo) = mGetChfCode(slStr)
                    gFindMatch smSave(3, imRowNo), 1, lbcAdvt
                    If gLastFound(lbcAdvt) > 0 Then
                        slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) <> tmChf.iAdfCode Then
                            Beep
                            mEnableBox imBoxNo
                            Exit Sub
                        End If
                    End If
                End If
            End If
            ilBox = imBoxNo - 1
        Case DEMOINDEX
            If smSave(3, imRowNo) <> "" Then
                ilBox = imBoxNo - 1
            Else
                ilBox = imBoxNo - 2
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
Private Sub pbcState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("T") Or (KeyAscii = Asc("t")) Then
        imState = 0
        pbcState_Paint
        pbcProj_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        imState = 1
        pbcState_Paint
        pbcProj_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imState = 0 Then
            imState = 1
            pbcState_Paint
            pbcProj_Paint
        ElseIf imState = 1 Then
            imState = 0
            pbcState_Paint
            pbcProj_Paint
        End If
    End If
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState = 0 Then
        imState = 1
    ElseIf imState = 1 Then
        imState = 0
    End If
    pbcState_Paint
    pbcProj_Paint
End Sub
Private Sub pbcState_Paint()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = pbcState.Height / 3 - pbcState.TextHeight("A") 'fgBoxInsetY
    If imState = 0 Then
        pbcState.Print "TOTALS"
    ElseIf imState = 1 Then
        pbcState.Print "DIFFERENCES"
    Else
        pbcState.Print "   "
    End If
    llColor = pbcState.ForeColor
    slFontName = pbcState.FontName
    flFontSize = pbcState.FontSize
    pbcState.FontBold = False
    pbcState.FontSize = 7
    pbcState.FontName = "Arial"
    pbcState.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = pbcState.Height - 2 * pbcState.TextHeight("A") 'fgBoxInsetY
    pbcState.Print "Prior Week Rollover"
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = pbcState.Height - pbcState.TextHeight("A") - 30 'fgBoxInsetY
    pbcState.Print smRolloverDate
    pbcState.FontSize = flFontSize
    pbcState.FontName = slFontName
    pbcState.FontSize = flFontSize
    'pbcState.ForeColor = llColor
    pbcState.FontBold = True
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    If imBoxNo = SOFFICEINDEX Then
        If mSOfficeBranch() Then
            Exit Sub
        End If
    End If
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
    If imBoxNo = POTINDEX Then
        If mPotBranch() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            If UBound(smSave, 2) > LBONE Then
                ilBox = COMMENTINDEX
                imRowNo = UBound(smSave, 2) - 1
            Else
                ilBox = SOFFICEINDEX
                imRowNo = LBONE 'LBound(smSave, 2)
                mInitNewProj
            End If
            imSettingValue = True
            If imRowNo <= vbcProj.LargeChange + 1 Then
                vbcProj.Value = 1
            Else
                vbcProj.Value = imRowNo - vbcProj.LargeChange - 1
            End If
            imSettingValue = False
            mGetShowPrices
            mGetShowOPrices
            mGetShowPPrices
            pbcProj_Paint
        Case PD3INDEX   'COMMENTINDEX 'Last control within header
            mSetShow imBoxNo
            If mTestSaveFields(imRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imPjfChg = True
                ReDim Preserve smShow(0 To TOTALINDEX, 0 To imRowNo + 1) As String 'Values shown in program area
                ReDim Preserve smSave(0 To 8, 0 To imRowNo + 1) As String 'Values saved (program name) in program area
                ReDim Preserve lmSave(0 To 5, 0 To imRowNo + 1) As Long 'Values saved (program name) in program area
                ReDim Preserve imSave(0 To 3, 0 To imRowNo + 1) As Integer 'Values saved (program name) in program area
                ReDim Preserve tgPjf1Rec(0 To UBound(tgPjf1Rec) + 1) As PJF1REC
                If imPrjNoYears > 1 Then
                    ReDim Preserve tgPjf2Rec(0 To UBound(tgPjf2Rec) + 1) As PJF2REC
                End If
            End If
            If imRowNo >= UBound(smSave, 2) - 1 Then
                imRowNo = imRowNo + 1
                mInitNewProj
                imSettingValue = True
                If UBound(smSave, 2) <= vbcProj.LargeChange Then 'was <=
                    vbcProj.Max = LBONE 'LBound(smSave, 2) '- 1
                Else
                    vbcProj.Max = UBound(smSave, 2) - vbcProj.LargeChange '- 1
                End If
                imSettingValue = False
            Else
                imRowNo = imRowNo + 1
            End If
            If imRowNo > vbcProj.Value + vbcProj.LargeChange Then
                imSettingValue = True
                vbcProj.Value = vbcProj.Value + 1
                imSettingValue = False
            End If
            mGetShowPrices
            mGetShowOPrices
            mGetShowPPrices
            pbcProj_Paint
            If imRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                mSetCommands
                'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcProj.Top + tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = SOFFICEINDEX 'PROPNOINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        'Case PD1INDEX To PD3INDEX   'DOLLARINDEX, PCTINVINDEX 'Last control within header
        '    ilEnd = False
        '    If imBoxNo - PD1INDEX + 2 >= 4 Then
        '        ilEnd = True
        '    Else
        '        If tmPdGroups(imBoxNo - PD1INDEX + 2).iStartWkNo < 0 Then
        '            ilEnd = True
        '        End If
        '    End If
        '    If ilEnd Then
        '        mSetShow imBoxNo
        '        If mTestSaveFields(imRowNo) = NO Then
        '            mEnableBox imBoxNo
        '            Exit Sub
        '        End If
        '        If imRowNo + 1 >= UBound(tgPjf1Rec) Then
        '            cmcDone.SetFocus
        '            Exit Sub
        '        End If
        '        imRowNo = imRowNo + 1
        '        If imRowNo > vbcProj.Value + vbcProj.LargeChange Then
        '            imSettingValue = True
        '            vbcProj.Value = vbcProj.Value + 1
        '            imSettingValue = False
        '        End If
        '        ilBox = PD1INDEX
        '        imBoxNo = ilBox
        '        mEnableBox ilBox
        '        Exit Sub
        '    Else
        '        ilBox = imBoxNo + 1
        '    End If
        Case 0
            ilBox = SOFFICEINDEX
        Case PROPNOINDEX
            If (lbcPropNo.ListIndex > 0) And (smSave(3, imRowNo) <> "") Then 'Test that advertiser name match
                slStr = lbcPropNo.List(lbcPropNo.ListIndex)
                If smSave(1, imRowNo) <> slStr Then
                    lmSave(5, imRowNo) = mGetChfCode(slStr)
                    gFindMatch smSave(3, imRowNo), 1, lbcAdvt
                    If gLastFound(lbcAdvt) > 0 Then
                        slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) <> tmChf.iAdfCode Then
                            Beep
                            mEnableBox imBoxNo
                            Exit Sub
                        End If
                    End If
                End If
            End If
            mSetShow imBoxNo
            If (imBoxNo < 1) And (imRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = imBoxNo + 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case ADVTINDEX
            mSetShow imBoxNo
            If smSave(1, imRowNo) <> "" Then
                llChfCode = mGetChfCode(smSave(1, imRowNo))
                gFindMatch smSave(3, imRowNo), 1, lbcAdvt
                If gLastFound(lbcAdvt) > 0 Then
                    slNameCode = tmAdvertiser(gLastFound(lbcAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> tmChf.iAdfCode Then
                        Beep
                        imBoxNo = ADVTINDEX
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
            End If
            ilBox = imBoxNo + 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case PRODINDEX
            ilBox = imBoxNo + 2 'Skip over Proposal field
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
Private Sub plcDetailSum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcProj_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub plcShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcDetailSum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slDate As String
    Dim slStart As String
    Dim ilYear As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If imIgnoreSetting Then
        imIgnoreSetting = False
        Exit Sub
    End If
    If Value Then
        Screen.MousePointer = vbHourglass
        pbcProj.Cls
        If imPrjStartYear <> 0 Then
            ilYear = imCurYear
            If Index = 0 Then    'Corporate
                slDate = "1/15/" & Trim$(Str$(ilYear))
                slStart = gObtainStartCorp(slDate, True)
            Else                        'Standard
                slDate = "1/15/" & Trim$(Str$(ilYear))
                slStart = gObtainStartStd(slDate)
            End If
            imPrjStartWk = (lmNowDate - gDateValue(slStart)) \ 7 + 1
            imPdStartWk = imPrjStartWk
            mGetShowDates True
            imPdYear = tmPdGroups(1).iYear
            If imTypeIndex = 1 Then 'By month
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(1).iStartWkNo
                ilFound = False
                Do
                    For ilLoop = LBONE To 12 Step 1
                        If imPdStartWk = ilStartWk(ilLoop) Then
                            ilFound = True
                            Exit Do
                        End If
                    Next ilLoop
                    If imPdStartWk <= 1 Then
                        imPdStartWk = 1
                        ilFound = True
                        Exit Do
                    End If
                    imPdStartWk = imPdStartWk - 1
                Loop Until ilFound
            ElseIf imTypeIndex = 2 Then 'Weeks- make sure not pass end
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(3).iStartWkNo
                ilFound = False
                Do
                    If (imPdStartWk > ilStartWk(12) + ilNoWks(12) - 6) Then
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Else
                        ilFound = True
                        Exit Do
                    End If
                Loop Until ilFound
            End If
        End If
        imShowIndex = Index
        mGetShowDates True
        pbcProj_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcShow_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub rbcShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcProj.Cls
        If imPrjStartYear <> 0 Then
            imPdYear = tmPdGroups(1).iYear
            If Index = 0 Then   'Change to Quarter
                imPdStartWk = imPrjStartWk
            ElseIf Index = 1 Then   'Month
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    If (imPdYear = imPrjStartYear) Then
                        imPdStartWk = imPrjStartWk
                    Else
                        imPdStartWk = 1
                    End If
                Else    'by week- back up to start of month
                    mCompMonths imPdYear, ilStartWk(), ilNoWks()
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                    ilFound = False
                    Do
                        For ilLoop = LBONE To 12 Step 1
                            If imPdStartWk = ilStartWk(ilLoop) Then
                                ilFound = True
                                Exit Do
                            End If
                        Next ilLoop
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        If (imPdYear = imPrjStartYear) And (imPdStartWk = imPrjStartWk) Then
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Loop Until ilFound
                End If
            ElseIf Index = 2 Then   'Week
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    If (imPdYear = imPrjStartYear) Then
                        imPdStartWk = imPrjStartWk
                    Else
                        imPdStartWk = 1
                    End If
                Else  'Month
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                End If
            Else    'Flight
                imPdStartFltNo = 1
            End If
        End If
        imTypeIndex = Index
        mGetShowDates True
        pbcProj_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub rbcType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case PROPNOINDEX
            lbcPropNo.Visible = Not lbcPropNo.Visible
        Case SOFFICEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
        Case ADVTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        Case PRODINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        Case DEMOINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcDemo, edcDropDown, imChgMode, imLbcArrowSetting
        Case VEHINDEX
        Case NRINDEX
        Case POTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcPot, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcProj.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(PROPNOINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(PROPNOINDEX).fBoxY + tmCtrls(PROPNOINDEX).fBoxH)) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcProj.Value - 1
                    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcProj.Top + tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame.Drag vbBeginDrag
                    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcProj_Change()
    If imSettingValue Then
        pbcProj.Cls
        pbcProj_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        imBoxNo = -1
        imRowNo = -1
        pbcProj.Cls
        pbcProj_Paint
        'If (igWinStatus(PROPOSALSJOB) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        '    mEnableBox imBoxNo
        'End If
    End If
End Sub
Private Sub vbcProj_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show by"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Projection"
End Sub

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
Private Sub mPaintProjTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer
    Dim llWidth As Long

    ilHalfY = tmCtrls(1).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop

    llColor = pbcProj.ForeColor
    slFontName = pbcProj.FontName
    flFontSize = pbcProj.FontSize
    ilFillStyle = pbcProj.FillStyle
    llFillColor = pbcProj.FillColor
    pbcProj.ForeColor = BLUE
    pbcProj.FontBold = False
    pbcProj.FontSize = 7
    pbcProj.FontName = "Arial"
    pbcProj.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    pbcProj.Line (tmCtrls(SOFFICEINDEX).fBoxX - 15, 15)-Step(tmCtrls(SOFFICEINDEX).fBoxW + 15, tmCtrls(SOFFICEINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(SOFFICEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Office"
    pbcProj.Line (tmCtrls(ADVTINDEX).fBoxX - 15, 15)-Step(tmCtrls(ADVTINDEX).fBoxW + 15, tmCtrls(ADVTINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(ADVTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Advertiser"
    pbcProj.Line (tmCtrls(PRODINDEX).fBoxX - 15, 15)-Step(tmCtrls(PRODINDEX).fBoxW + 15, tmCtrls(PRODINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(PRODINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Product"
    pbcProj.Line (tmCtrls(PROPNOINDEX).fBoxX - 15, 15)-Step(tmCtrls(PROPNOINDEX).fBoxW + 15, tmCtrls(PROPNOINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(PROPNOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Proposal"
    pbcProj.Line (tmCtrls(DEMOINDEX).fBoxX - 15, 15)-Step(tmCtrls(DEMOINDEX).fBoxW + 15, tmCtrls(DEMOINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(DEMOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "D"
    pbcProj.Line (tmCtrls(VEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(VEHINDEX).fBoxW + 15, tmCtrls(VEHINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Vehicle"
    pbcProj.Line (tmCtrls(NRINDEX).fBoxX - 15, 15)-Step(tmCtrls(NRINDEX).fBoxW + 15, tmCtrls(NRINDEX).fBoxY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(NRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "N/"
    pbcProj.CurrentX = tmCtrls(NRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = ilHalfY + 15
    pbcProj.Print "R"
    pbcProj.Line (tmCtrls(POTINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(POTINDEX).fBoxW + 15, tmCtrls(POTINDEX).fBoxY - ilHalfY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(POTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = ilHalfY + 15
    pbcProj.Print "%"
    pbcProj.Line (tmCtrls(COMMENTINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(COMMENTINDEX).fBoxW + 15, tmCtrls(COMMENTINDEX).fBoxY - ilHalfY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(COMMENTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = ilHalfY + 15
    pbcProj.Print "C"
    pbcProj.Line (tmCtrls(PD1INDEX).fBoxX - 15, 15)-Step(tmCtrls(PD1INDEX).fBoxW + 15, tmCtrls(PD1INDEX).fBoxY - 30), BLUE, B
    pbcProj.Line (tmCtrls(PD2INDEX).fBoxX - 15, 15)-Step(tmCtrls(PD2INDEX).fBoxW + 15, tmCtrls(PD2INDEX).fBoxY - 30), BLUE, B
    pbcProj.Line (tmCtrls(PD3INDEX).fBoxX - 15, 15)-Step(tmCtrls(PD3INDEX).fBoxW + 15, tmCtrls(PD3INDEX).fBoxY - 30), BLUE, B
    pbcProj.Line (tmCtrls(TOTALINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(TOTALINDEX).fBoxW + 15, tmCtrls(TOTALINDEX).fBoxY - ilHalfY - 30), BLUE, B
    pbcProj.CurrentX = tmCtrls(TOTALINDEX).fBoxX + 15  'fgBoxInsetX
    pbcProj.CurrentY = ilHalfY + 15
    pbcProj.Print "Total"


    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            pbcProj.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + 4 * (tmCtrls(1).fBoxH + 15) + 60 < pbcProj.Height
    vbcProj.LargeChange = ilLineCount - 1
    llTop = llTop + 30
    pbcState.Top = llTop
    llWidth = pbcProj.TextWidth("Current Week Advertiser/Product      ")
    pbcState.Left = tmCAPCtrls(GTDOLLAR1INDEX).fBoxX - llWidth - pbcState.Width - 30
    pbcProj.Line (tmCAPCtrls(GTDOLLAR1INDEX).fBoxX - llWidth - 15, llTop - 15)-Step(llWidth + 15, tmCAPCtrls(GTDOLLAR1INDEX).fBoxH + 15), BLUE, B
    pbcProj.CurrentX = tmCAPCtrls(GTDOLLAR1INDEX).fBoxX - llWidth + 15  'fgBoxInsetX
    pbcProj.CurrentY = llTop + 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Current Week Advertiser/Product"
    tmCAPCtrls(GTDOLLAR1INDEX).fBoxY = llTop
    tmCAPCtrls(GTDOLLAR2INDEX).fBoxY = llTop
    tmCAPCtrls(GTDOLLAR3INDEX).fBoxY = llTop
    tmCAPCtrls(GTTOTALINDEX).fBoxY = llTop
    For ilLoop = GTDOLLAR1INDEX To GTTOTALINDEX Step 1
        pbcProj.Line (tmCAPCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCAPCtrls(ilLoop).fBoxW + 15, tmCAPCtrls(ilLoop).fBoxH + 15), BLUE, B
    Next ilLoop
    llTop = llTop + tmCtrls(1).fBoxH + 15
    pbcProj.Line (tmPAPCtrls(GTDOLLAR1INDEX).fBoxX - llWidth - 15, llTop - 15)-Step(llWidth + 15, tmPAPCtrls(GTDOLLAR1INDEX).fBoxH + 15), BLUE, B
    pbcProj.CurrentX = tmCAPCtrls(GTDOLLAR1INDEX).fBoxX - llWidth + 15  'fgBoxInsetX
    pbcProj.CurrentY = llTop + 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Prior Week Advertiser/Product"
    tmPAPCtrls(GTDOLLAR1INDEX).fBoxY = llTop
    tmPAPCtrls(GTDOLLAR2INDEX).fBoxY = llTop
    tmPAPCtrls(GTDOLLAR3INDEX).fBoxY = llTop
    tmPAPCtrls(GTTOTALINDEX).fBoxY = llTop
    For ilLoop = GTDOLLAR1INDEX To GTTOTALINDEX Step 1
        pbcProj.Line (tmPAPCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmPAPCtrls(ilLoop).fBoxW + 15, tmPAPCtrls(ilLoop).fBoxH + 15), BLUE, B
    Next ilLoop
    llTop = llTop + tmCtrls(1).fBoxH + 15
    pbcProj.Line (tmCTCtrls(GTDOLLAR1INDEX).fBoxX - llWidth - 15, llTop - 15)-Step(llWidth + 15, tmCTCtrls(GTDOLLAR1INDEX).fBoxH + 15), BLUE, B
    pbcProj.CurrentX = tmCTCtrls(GTDOLLAR1INDEX).fBoxX - llWidth + 15  'fgBoxInsetX
    pbcProj.CurrentY = llTop + 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Current Week Totals"
    tmCTCtrls(GTDOLLAR1INDEX).fBoxY = llTop
    tmCTCtrls(GTDOLLAR2INDEX).fBoxY = llTop
    tmCTCtrls(GTDOLLAR3INDEX).fBoxY = llTop
    tmCTCtrls(GTTOTALINDEX).fBoxY = llTop
    For ilLoop = GTDOLLAR1INDEX To GTTOTALINDEX Step 1
        pbcProj.Line (tmCTCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCTCtrls(ilLoop).fBoxW + 15, tmCTCtrls(ilLoop).fBoxH + 15), BLUE, B
    Next ilLoop
    llTop = llTop + tmCtrls(1).fBoxH + 15
    pbcProj.Line (tmPTCtrls(GTDOLLAR1INDEX).fBoxX - llWidth - 15, llTop - 15)-Step(llWidth + 15, tmPTCtrls(GTDOLLAR1INDEX).fBoxH + 15), BLUE, B
    pbcProj.CurrentX = tmPTCtrls(GTDOLLAR1INDEX).fBoxX - llWidth + 15  'fgBoxInsetX
    pbcProj.CurrentY = llTop + 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcProj.Print "Prior Week Totals"
    tmPTCtrls(GTDOLLAR1INDEX).fBoxY = llTop
    tmPTCtrls(GTDOLLAR2INDEX).fBoxY = llTop
    tmPTCtrls(GTDOLLAR3INDEX).fBoxY = llTop
    tmPTCtrls(GTTOTALINDEX).fBoxY = llTop
    For ilLoop = GTDOLLAR1INDEX To GTTOTALINDEX Step 1
        pbcProj.Line (tmPTCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmPTCtrls(ilLoop).fBoxW + 15, tmPTCtrls(ilLoop).fBoxH + 15), BLUE, B
    Next ilLoop
    pbcProj.FontSize = flFontSize
    pbcProj.FontName = slFontName
    pbcProj.FontSize = flFontSize
    pbcProj.ForeColor = llColor
    pbcProj.FontBold = True
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
Private Sub mParseCmmdLine()

    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim slCSIName As String
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer

    
    sgCommandStr = Command$
    slStartIn = CurDir$
    slCSIName = ""
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
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        'Signon.Show vbModal
        MsgBox "Contract Projection must be run from Traffic->Proposals"
        igExitTraffic = True
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        slStr = sgUserName
        sgCallAppName = "Traffic"
    Else
        igSportsSystem = 0
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        'ilRet = gParseItem(slCommand, 3, "\", slStr)
        'igRptCallType = Val(slStr)
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            slDate = Format$(Now(), "m/d/yy")
            slMonth = Month(slDate)
            slYear = Year(slDate)
            llValue = Val(slMonth) * Val(slYear)
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            llValue = ilValue
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            slStr = Trim$(Str$(ilValue))
            Do While Len(slStr) < 4
                slStr = "0" & slStr
            Loop
            sgSpecialPassword = slStr
            slCSIName = "CSI"
            sgUserName = "Guide"
        End If
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        If StrComp(slCSIName, "CSI", vbTextCompare) = 0 Then
            gExpandGuideAsUser tgUrf(0)
        End If
        mGetUlfCode
    End If
    'End If
    DoEvents
'    gInitStdAlone ReportList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "CntrProj.Frm"
    If igWinStatus(PROPOSALSJOB) = 0 Then
        imTerminate = True
    End If
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    
End Sub
Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
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

