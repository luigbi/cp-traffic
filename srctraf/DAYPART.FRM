VERSION 5.00
Begin VB.Form Daypart 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   1890
   ClientTop       =   1650
   ClientWidth     =   6030
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
   ScaleHeight     =   4650
   ScaleWidth      =   6030
   Begin VB.ListBox lbcDPGroup 
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
      ItemData        =   "DAYPART.frx":0000
      Left            =   570
      List            =   "DAYPART.frx":0002
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox pbcOverridePrt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4560
      ScaleHeight     =   210
      ScaleWidth      =   1140
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      Picture         =   "DAYPART.frx":0004
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2445
      TabIndex        =   1
      Top             =   165
      Width           =   3120
   End
   Begin VB.ListBox lbcAvail 
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
      Left            =   645
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1785
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.ListBox lbcLibNm 
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
      Left            =   3105
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2475
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ListBox lbcLibNm 
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
      Left            =   3135
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3030
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ListBox lbcLibNm 
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
      Left            =   3255
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2790
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Left            =   3000
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2970
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
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
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "DAYPART.frx":030E
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "DAYPART.frx":0FCC
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
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
      End
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
      HelpContextID   =   1
      Left            =   3690
      TabIndex        =   37
      Top             =   4245
      Width           =   945
   End
   Begin VB.PictureBox pbcActive 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4470
      ScaleHeight     =   210
      ScaleWidth      =   990
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5325
      Top             =   2865
   End
   Begin VB.CommandButton cmcLibDropDown 
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
      Left            =   2760
      Picture         =   "DAYPART.frx":12D6
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcLibDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1755
      MaxLength       =   10
      TabIndex        =   31
      Top             =   2220
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   1500
      Picture         =   "DAYPART.frx":13D0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcSpecDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   555
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1065
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ListBox lbcDPNameCode 
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
      Left            =   5190
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   3675
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pbcHrs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5445
      ScaleHeight     =   210
      ScaleWidth      =   285
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox pbcAvType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4260
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcTDTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   405
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   34
      Top             =   3735
      Width           =   135
   End
   Begin VB.PictureBox pbcTDSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   420
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   16
      Top             =   2055
      Width           =   105
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
      HelpContextID   =   2
      Left            =   2475
      TabIndex        =   36
      Top             =   4245
      Width           =   945
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
      HelpContextID   =   1
      Left            =   1230
      TabIndex        =   35
      Top             =   4245
      Width           =   945
   End
   Begin VB.TextBox edcTDStartEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   735
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3375
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox edcTDPCSpot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3075
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmcTDStartEnd 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1725
      Picture         =   "DAYPART.frx":14CA
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcTDDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   3735
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   210
      Begin VB.CheckBox ckcTDDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   15
         TabIndex        =   22
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4005
      Width           =   75
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   165
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   12
      Top             =   1215
      Width           =   135
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   540
      Width           =   105
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
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1575
   End
   Begin VB.PictureBox pbcSpec 
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
      Height          =   1080
      Left            =   435
      Picture         =   "DAYPART.frx":15C4
      ScaleHeight     =   1080
      ScaleWidth      =   5085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   585
      Width           =   5085
   End
   Begin VB.PictureBox plcSpec 
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   360
      ScaleHeight     =   1080
      ScaleWidth      =   5145
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   5205
   End
   Begin VB.PictureBox pbcLib 
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
      Height          =   1080
      Left            =   1680
      Picture         =   "DAYPART.frx":134E6
      ScaleHeight     =   1080
      ScaleWidth      =   2445
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox plcLib 
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   1605
      ScaleHeight     =   1140
      ScaleWidth      =   2520
      TabIndex        =   29
      Top             =   2130
      Width           =   2580
   End
   Begin VB.PictureBox pbcTD 
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
      Height          =   1725
      Left            =   675
      Picture         =   "DAYPART.frx":16398
      ScaleHeight     =   1725
      ScaleWidth      =   4320
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.PictureBox plcTD 
      ForeColor       =   &H00000000&
      Height          =   1830
      Left            =   615
      ScaleHeight     =   1770
      ScaleWidth      =   4395
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2130
      Width           =   4455
   End
   Begin VB.ListBox lbcLibNmCode 
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
      Left            =   4515
      Sorted          =   -1  'True
      TabIndex        =   40
      Top             =   3990
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lbcAvailCode 
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
      Left            =   4545
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   4275
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4905
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5415
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5295
      TabIndex        =   42
      Top             =   3285
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton rbcType 
      Caption         =   "Times"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   2250
      TabIndex        =   14
      Top             =   1875
      Width           =   840
   End
   Begin VB.OptionButton rbcType 
      Caption         =   "Libraries"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3165
      TabIndex        =   15
      Top             =   1875
      Width           =   1305
   End
   Begin VB.Label lacType 
      Caption         =   "Daypart by"
      Height          =   210
      Left            =   1095
      TabIndex        =   13
      Top             =   1875
      Width           =   960
   End
End
Attribute VB_Name = "Daypart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Daypart.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Daypart input screen code
'   This screen is used within RateCard and Traffic(Contract)
Option Explicit
Option Compare Text
'Dayparts
Dim hmRdf As Integer        'file handle
Dim tmRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer  'RDF record length
Dim tmAvailCode() As SORTCODE
Dim smAvailCodeTag As String
Dim tmDPNameCode() As SORTCODE
Dim smDPNameCodeTag As String
Dim tmDPGroupCode() As SORTCODE
Dim smDPGroupCodeTag As String

'Library Names
Dim hmLtf As Integer        'Library title file handle
Dim tmLtf As LTF            'LTF record image
Dim tmLtfSrchKey As INTKEY0 'LTF key record image
Dim imLtfRecLen As Integer  'LTF record length
'Specification Area
Dim tmCtrls(0 To 10) As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer
Dim smSave(0 To 3) As String  'Values saved (1=Daypart Name; 2=Avail name; 3=Daypart Group name)
Dim imSave(0 To 7) As Integer 'Values saved (1=Base daypart; 2=Report Daypart; 3=Avail type;
                              '4=Hour override; 5=Sort Code; 6=Active/Dormant; 7=Override Print Rule)
'Time/Day Field Areas
Dim tmTDCtrls(0 To 10)  As FIELDAREA    'Time/Days
Dim imLBTDCtrls As Integer
Dim imTDBoxNo As Integer
Dim imTDRowNo As Integer
Dim imMaxTDRows As Integer  'Max number of times defined for each program
Dim smTDShow(0 To 7, 0 To 10) As String 'Values shown in time/day area
                            'imMaxTDRows and first dimension should be the same
Dim smTDSave(0 To 7, 0 To 10) As String 'Values retained from time/day area
Dim imTDChg As Integer
'Library
Dim tmLibCtrls(0 To 3)  As FIELDAREA    'Library
Dim imLBLibCtrls As Integer
Dim imLibBoxNo As Integer
Dim smLibSave(0 To 3) As String
Dim tmLibNmCode() As SORTCODE
Dim smLibNmCodeTag As String
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassFocus As Integer
Dim imUpdateAllowed As Integer
Dim imLbcArrowSetting As Integer
Dim imTabDirection As Integer
Dim imDirProcess As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer
Dim imRdfChg As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imFirstFocus As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imAnyZeroPct As Integer      'Any negative percentage
Dim smOrigDPGroup As String

Const LBONE = 1

Const NAMEINDEX = 1          'Daypart Name control/field
Const BASEINDEX = 2      'Base Daypart control/field
Const RPTINDEX = 3         'Report Daypart control/field
Const AVAILINDEX = 4        'Avail name control/field
Const HRSINDEX = 6          'Hour override control/field
Const OVERRIDEPRTINDEX = 7          'Override Print rule control/field
Const DPGROUPINDEX = 8
Const SORTINDEX = 9       'Sort control/field
Const ACTIVEINDEX = 10
Const TDSTARTINDEX = 1      'Start time control/field
Const TDENDINDEX = 2        'End time control/field
Const TDDAYINDEX = 3        'Day control/field
Const TDPCSPOTINDEX = 10    'Percent spot control/field
Const LIB1INDEX = 1
Const LIB2INDEX = 2
Const LIB3INDEX = 3

Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then  'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcSpec.Cls
    pbcLib.Cls
    pbcTD.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
        mInitShow
    Else
        imSelectedIndex = 0
        If slStr <> "[New]" Then
            edcSpecDropDown.MaxLength = 20
            edcSpecDropDown.Text = slStr
            mSetShow NAMEINDEX
        End If
    End If
    pbcSpec_Paint
    pbcTD_Paint
    pbcLib_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
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
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    pbcArrow.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'mInitDDE
        If igDPNameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgDPName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgDPName    'New name
            End If
            cbcSelect_Change
            If sgDPName <> "" Then
                mSetCommands
                gFindMatch sgDPName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSpecSTab.Enabled Then
                pbcSpecSTab.SetFocus
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
        If pbcSpecSTab.Enabled Then
            pbcSpecSTab.SetFocus
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
Private Sub ckcTDDay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_Click()
    If igDPNameCallSource <> CALLNONE Then
        igDPNameCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igDPNameCallSource <> CALLNONE Then
        sgDPName = smSave(1) 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgDPName = "[New]"
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
    If igDPNameCallSource <> CALLNONE Then
        If sgDPName = "[New]" Then
            igDPNameCallSource = CALLCANCELLED
        Else
            igDPNameCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcLibDropDown_Click()
    Select Case imLibBoxNo
        Case LIB1INDEX
            lbcLibNm(0).Visible = Not lbcLibNm(0).Visible
        Case LIB2INDEX
            lbcLibNm(1).Visible = Not lbcLibNm(1).Visible
        Case LIB3INDEX
            lbcLibNm(2).Visible = Not lbcLibNm(2).Visible
    End Select
    edcLibDropDown.SelStart = 0
    edcLibDropDown.SelLength = Len(edcLibDropDown.Text)
    edcLibDropDown.SetFocus
End Sub
Private Sub cmcLibDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSpecDropDown_Click()
    Select Case imBoxNo
        Case NAMEINDEX
        Case AVAILINDEX + 1
            lbcAvail.Visible = Not lbcAvail.Visible
        Case DPGROUPINDEX
            lbcDPGroup.Visible = Not lbcDPGroup.Visible
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcTDStartEnd_Click()
    Select Case imTDBoxNo
        Case TDSTARTINDEX
            plcTme.Visible = Not plcTme.Visible
        Case TDENDINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcTDStartEnd.SelStart = 0
    edcTDStartEnd.SelLength = Len(edcTDStartEnd.Text)
    edcTDStartEnd.SetFocus
End Sub
Private Sub cmcTDStartEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = cbcSelect.Text   'Save name
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
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    ilCode = tmRdf.iCode
    cbcSelect.Clear
    'lbcDPNameCode.Tag = ""
    smDPNameCodeTag = ""
    mPopulate
    For ilLoop = 0 To UBound(tmDPNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tmDPNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
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
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    Else
        cmcCancel.SetFocus
    End If
End Sub
Private Sub edcLibDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case LIB1INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcLibDropDown, lbcLibNm(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcLibNm(0).ListIndex = 0
            End If
        Case LIB1INDEX + 1
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcLibDropDown, lbcLibNm(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcLibNm(1).ListIndex = 0
            End If
        Case LIB1INDEX + 2
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcLibDropDown, lbcLibNm(2), imBSMode, slStr)
            If ilRet = 1 Then
                lbcLibNm(2).ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcLibDropDown_GotFocus()
    Select Case imBoxNo
        Case LIB1INDEX + 0
            If lbcLibNm(0).ListCount = 1 Then
                lbcLibNm(0).ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcTDSTab.SetFocus
                'Else
                '    pbcTDTab.SetFocus
                'End If
                Exit Sub
            End If
        Case LIB1INDEX + 1
            If lbcLibNm(1).ListCount = 1 Then
                lbcLibNm(1).ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcTDSTab.SetFocus
                'Else
                '    pbcTDTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case LIB1INDEX + 2
            If lbcLibNm(2).ListCount = 1 Then
                lbcLibNm(2).ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcTDSTab.SetFocus
                'Else
                '    pbcTDTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcLibDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcLibDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcLibDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLibDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTDTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case LIB1INDEX
                gProcessArrowKey Shift, KeyCode, lbcLibNm(0), imLbcArrowSetting
            Case LIB1INDEX + 1
                gProcessArrowKey Shift, KeyCode, lbcLibNm(1), imLbcArrowSetting
            Case LIB1INDEX + 2
                gProcessArrowKey Shift, KeyCode, lbcLibNm(2), imLbcArrowSetting
        End Select
        edcLibDropDown.SelStart = 0
        edcLibDropDown.SelLength = Len(edcLibDropDown.Text)
    End If
End Sub
Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case NAMEINDEX
        Case AVAILINDEX + 1
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcSpecDropDown, lbcAvail, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAvail.ListIndex = 0
            End If
        Case DPGROUPINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcSpecDropDown, lbcDPGroup, imBSMode, slStr)
            If ilRet = 1 Then
                lbcDPGroup.ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo, False
End Sub
Private Sub edcSpecDropDown_DblClick()
    If imBoxNo = AVAILINDEX + 1 Then
        imDoubleClickName = True    'Double click event foolowed by mouse up
    End If
    If imBoxNo = DPGROUPINDEX Then
        imDoubleClickName = True    'Double click event foolowed by mouse up
    End If
End Sub
Private Sub edcSpecDropDown_GotFocus()
    Select Case imBoxNo
        Case NAMEINDEX
        Case AVAILINDEX + 1
            If lbcAvail.ListCount = 1 Then
                lbcAvail.ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcSpecSTab.SetFocus
                'Else
                '    pbcSpecTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case DPGROUPINDEX
            If (lbcDPGroup.ListCount = 2) And (lbcDPGroup.ListIndex = -1) Then
                lbcDPGroup.ListIndex = 1
            End If
    End Select
End Sub
Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpecDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If

    ilKey = KeyAscii
    '4/17/13: Allow forward slash as the problem was two daypart with similar names but only one defined in rate card and the other defined in line
    ''4/16/13: Disallow forward slash in the daypart name
    'If imBoxNo = NAMEINDEX Then
    '    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
    '        KeyAscii = 0
    '        Exit Sub
    '    End If
    'Else
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    'End If
End Sub
Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcSpecTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case NAMEINDEX
            Case AVAILINDEX + 1
                gProcessArrowKey Shift, KeyCode, lbcAvail, imLbcArrowSetting
            Case DPGROUPINDEX
                gProcessArrowKey Shift, KeyCode, lbcDPGroup, imLbcArrowSetting
        End Select
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    End If
End Sub
Private Sub edcSpecDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case AVAILINDEX + 1
                If imTabDirection = -1 Then  'Right To Left
                    pbcSpecSTab.SetFocus
                Else
                    pbcSpecTab.SetFocus
                End If
                Exit Sub
            Case DPGROUPINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSpecSTab.SetFocus
                Else
                    pbcSpecTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcTDPCSpot_GotFocus()
    gCtrlGotFocus edcTDPCSpot
End Sub
Private Sub edcTDPCSpot_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcTDPCSpot.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcTDPCSpot.Text, ".")    'Disallow multi-decimal points
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
    If Not imAnyZeroPct Then
        slStr = edcTDPCSpot.Text
        slStr = Left$(slStr, edcTDPCSpot.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcTDPCSpot.SelStart - edcTDPCSpot.SelLength)
        If gCompNumberStr(slStr, "100.00") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcTDStartEnd_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTDStartEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTDStartEnd_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
            If KeyAscii = igLegalTime(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcTDStartEnd_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imTDBoxNo
            Case TDSTARTINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
                edcTDStartEnd.SelStart = 0
                edcTDStartEnd.SelLength = Len(edcTDStartEnd.Text)
            Case TDENDINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
                edcTDStartEnd.SelStart = 0
                edcTDStartEnd.SelLength = Len(edcTDStartEnd.Text)
        End Select
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
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpec.Enabled = False
        pbcTD.Enabled = False
        pbcLib.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcTDSTab.Enabled = False
        pbcTDTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcSpec.Enabled = True
        pbcTD.Enabled = True
        pbcLib.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        pbcTDSTab.Enabled = True
        pbcTDTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Daypart.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcTme.Visible = False
        If (cbcSelect.Enabled) And ((imBoxNo > 0) Or (imTDBoxNo > 0) Or (imLibBoxNo > 0)) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imTDBoxNo > 0 Then
            mTDEnableBox imTDBoxNo
        ElseIf imLibBoxNo > 0 Then
            mLibEnableBox imLibBoxNo
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
    Cancel = 0
    
    Erase tmAvailCode
    Erase tmLibNmCode
    Erase tmDPNameCode
    ilRet = btrClose(hmLtf)
    btrDestroy hmLtf
    btrExtClear hmRdf   'Clear any previous extend operation
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    Set Daypart = Nothing   'Remove data segment

End Sub
Private Sub lbcAvail_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAvail, edcSpecDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcAvail_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcAvail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcAvail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcAvail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAvail, edcSpecDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSpecSTab.SetFocus
        Else
            pbcSpecTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcDPGroup_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcDPGroup, edcSpecDropDown, imChgMode, imLbcArrowSetting
    End If

End Sub

Private Sub lbcDPGroup_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcDPGroup_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcDPGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcDPGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcDPGroup, edcSpecDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSpecSTab.SetFocus
        Else
            pbcSpecTab.SetFocus
        End If
    End If

End Sub

Private Sub lbcLibNm_Click(Index As Integer)
    gProcessLbcClick lbcLibNm(Index), edcLibDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLibNm_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailsBranch                   *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Avail  *
'*                      names and process communication*
'*                      back from avail names          *
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
Private Function mAvailsBranch()
'
'   ilRet = mAvailsBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcSpecDropDown, lbcAvail, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mAvailsBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(AVAILNAMESLIST)) Then
    '    mAvailsBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igANmCallSource = CALLSOURCERATECARD
    If edcSpecDropDown.Text = "[New]" Then
        sgANmName = ""
    Else
        sgANmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igDPCallSource = 0 Then
        'Traffic!edcLinkSrceHelpMsg.Text = ""
        If igTestSystem Then
            slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
        Else
            slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
        End If
        'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
        'Traffic.Enabled = False
        'Screen.MousePointer = vbDefault  'Wait
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        sgCommandStr = slStr
        AName.Show vbModal
        slStr = sgDoneMsg
        ilParse = gParseItem(slStr, 1, "\", sgANmName)
        igANmCallSource = Val(sgANmName)
        ilParse = gParseItem(slStr, 2, "\", sgANmName)
        'Traffic.Enabled = True
        'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
    Else
        'RateCard!edcLinkSrceDoneMsg.Text = ""
        'If (Not igStdAloneMode) And (igShowHelpMsg) Then
            If igTestSystem Then
                slStr = "RateCard^Test\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
            Else
                slStr = "RateCard^Prod\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
            End If
        'Else
        '    If igTestSystem Then
        '        slStr = "RateCard^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
        '    Else
        '        slStr = "RateCard^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
        '    End If
        'End If
        'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
        'Daypart.Enabled = False
        ''Screen.MousePointer = vbDefault  'Wait
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        sgCommandStr = slStr
        AName.Show vbModal
        slStr = sgDoneMsg
        ilParse = gParseItem(slStr, 1, "\", sgANmName)
        igANmCallSource = Val(sgANmName)
        ilParse = gParseItem(slStr, 2, "\", sgANmName)
        'Daypart.Enabled = True
        ''Traffic!edcLinkSrceHelpMsg.Text = "Ok"
        'RateCard!edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    End If
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAvailsBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igANmCallSource = CALLDONE Then  'Done
        igANmCallSource = CALLNONE
        lbcAvail.Clear
        smAvailCodeTag = ""
        sgAvailAnfStamp = ""
        mAvailsPop
        If imTerminate Then
            mAvailsBranch = False
            Exit Function
        End If
        gFindMatch sgANmName, 1, lbcAvail
        sgANmName = ""
        If gLastFound(lbcAvail) > 0 Then
            imChgMode = True
            lbcAvail.ListIndex = gLastFound(lbcAvail)
            edcSpecDropDown.Text = lbcAvail.List(lbcAvail.ListIndex)
            imChgMode = False
            mAvailsBranch = False
        Else
            imChgMode = True
            lbcAvail.ListIndex = 1
            edcSpecDropDown.Text = lbcAvail.List(lbcAvail.ListIndex)
            imChgMode = False
            edcSpecDropDown.SetFocus
            Exit Function
        End If
    End If
    If igANmCallSource = CALLCANCELLED Then  'Cancelled
        igANmCallSource = CALLNONE
        sgANmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igANmCallSource = CALLTERMINATED Then
        igANmCallSource = CALLNONE
        sgANmName = ""
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
'*      Procedure Name:mAvailsPop                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAvailsPop()
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
    ilIndex = lbcAvail.ListIndex
    If ilIndex > 0 Then
        slName = lbcAvail.List(ilIndex)
    End If
'    ilFilter(0) = NOFILTER
'    slFilter(0) = ""
'    ilOffset(0) = 0
    ilFilter(0) = CHARFILTERNOT
    slFilter(0) = "F"
    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'ilRet = gIMoveListBox(Daypart, lbcAvail, lbcAvailCode, "Anf.Btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Daypart, lbcAvail, tmAvailCode(), smAvailCodeTag, "Anf.Btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To lbcAvail.ListCount - 1 Step 1
            slStr = Trim$(lbcAvail.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                lbcAvail.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tmAvailCode) - 1 Step 1
                    tmAvailCode(ilLp) = tmAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tmAvailCode(LBound(tmAvailCode) To UBound(tmAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mAvailsErr
        gCPErrorMsg ilRet, "mAvailsPop (gIMoveListBox: Avails)", Daypart
        On Error GoTo 0
        lbcAvail.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 2, lbcAvail
            If gLastFound(lbcAvail) > 1 Then
                lbcAvail.ListIndex = gLastFound(lbcAvail)
            Else
                lbcAvail.ListIndex = -1
            End If
        Else
            lbcAvail.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAvailsErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilMax As Integer
    imTDChg = False
    lbcAvail.ListIndex = -1
    lbcDPGroup.ListIndex = -1
    smOrigDPGroup = ""
    For ilLoop = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilLoop) = ""
    Next ilLoop
    For ilLoop = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilLoop) = -1
    Next ilLoop
    imSave(5) = 0   'Sort code
    For ilRowNo = 1 To 7 Step 1
        For ilBoxNo = 1 To 10 Step 1
            smTDSave(ilRowNo, ilBoxNo) = ""
            smTDShow(ilRowNo, ilBoxNo) = ""
        Next ilBoxNo
    Next ilRowNo
    smLibSave(1) = ""
    smLibSave(2) = ""
    smLibSave(3) = ""
    tmLibCtrls(1).sShow = ""
    tmLibCtrls(2).sShow = ""
    tmLibCtrls(3).sShow = ""
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).sShow = ""
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    tmRdf.sName = ""
    tmRdf.sBase = ""
    tmRdf.sReport = ""
    tmRdf.sTimeOver = ""
    tmRdf.sState = ""
    For ilLoop = 1 To imMaxTDRows Step 1
        slStr = ""
        gPackTime slStr, tmRdf.iStartTime(0, ilLoop - 1), tmRdf.iStartTime(1, ilLoop - 1)
        slStr = ""
        gPackTime slStr, tmRdf.iEndTime(0, ilLoop - 1), tmRdf.iEndTime(1, ilLoop - 1)
        tmRdf.iEndTime(0, ilLoop - 1) = 1
        tmRdf.iEndTime(1, ilLoop - 1) = 0
        For ilMax = 1 To 7 Step 1
            tmRdf.sWkDays(ilLoop - 1, ilMax - 1) = ""
        Next ilMax
        'slStr = ""
        'gStrToPDN slStr, 2, 3, tmRdf.sSpotPct(ilLoop)
        tmRdf.iSpotPct(ilLoop - 1) = 0
    Next ilLoop
    imAnyZeroPct = False
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
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    pbcArrow.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcSpecDropDown.Text = smSave(1)
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case BASEINDEX
            If imSave(1) < 0 Then
                imSave(1) = 0    'Yes
                mSetChg imBoxNo, True
            End If
            pbcHrs.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcHrs, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcHrs_Paint
            pbcHrs.Visible = True
            pbcHrs.SetFocus
        Case RPTINDEX
            If imSave(2) < 0 Then
                imSave(2) = 0    'Yes
                tmCtrls(BASEINDEX).iChg = True
                mSetChg imBoxNo, True
            End If
            pbcHrs.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcHrs, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcHrs_Paint
            pbcHrs.Visible = True
            pbcHrs.SetFocus
        Case AVAILINDEX
            If imSave(3) < 0 Then
                imSave(3) = 0    'N/A
                tmCtrls(AVAILINDEX).iChg = True
                mSetChg imBoxNo, True
            End If
            pbcAvType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcAvType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcAvType_Paint
            pbcAvType.Visible = True
            pbcAvType.SetFocus
        Case AVAILINDEX + 1
            mAvailsPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAvail.Height = gListBoxHeight(lbcAvail.ListCount, 10)
            edcSpecDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 20
            If imSave(3) = 1 Then
                gMoveFormCtrl pbcSpec, edcSpecDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            Else
                '                                                                                        Alt + 0216 or 0248
                gMoveFormCtrl pbcSpec, edcSpecDropDown, tmCtrls(ilBoxNo).fBoxX + pbcSpec.TextWidth("Ø") + 15, tmCtrls(ilBoxNo).fBoxY
            End If
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcAvail.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
'            If lbcAvail.ListIndex < 0 Then
'                lbcAvail.ListIndex = 0   'First one within list
'            End If
            imChgMode = True
            gFindMatch smSave(2), 0, lbcAvail
            If gLastFound(lbcAvail) >= 0 Then
                lbcAvail.ListIndex = gLastFound(lbcAvail)
                edcSpecDropDown.Text = lbcAvail.List(lbcAvail.ListIndex)
            Else
                If lbcAvail.ListCount > 1 Then
                    lbcAvail.ListIndex = 1
                    edcSpecDropDown.Text = lbcAvail.List(lbcAvail.ListIndex)
                Else
                    lbcAvail.ListIndex = -1
                    edcSpecDropDown.Text = ""
                End If
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case HRSINDEX    'Source Index
            If imSave(4) < 0 Then
                imSave(4) = 0    'All avails
                mSetChg imBoxNo, True
            End If
            pbcHrs.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcHrs, tmCtrls(HRSINDEX).fBoxX, tmCtrls(HRSINDEX).fBoxY
            pbcHrs_Paint
            pbcHrs.Visible = True
            pbcHrs.SetFocus
        Case OVERRIDEPRTINDEX    'Source Index
            If imSave(7) < 0 Then
                imSave(7) = 0    'None
                mSetChg imBoxNo, True
            End If
            pbcOverridePrt.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcOverridePrt, tmCtrls(OVERRIDEPRTINDEX).fBoxX, tmCtrls(OVERRIDEPRTINDEX).fBoxY
            pbcOverridePrt_Paint
            pbcOverridePrt.Visible = True
            pbcOverridePrt.SetFocus
        Case DPGROUPINDEX
            mDPGroupPop
            If imTerminate Then
                Exit Sub
            End If
            lbcDPGroup.Height = gListBoxHeight(lbcDPGroup.ListCount, 10)
            edcSpecDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcDPGroup.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
'            If lbcDPGroup.ListIndex < 0 Then
'                lbcDPGroup.ListIndex = 0   'First one within list
'            End If
            imChgMode = True
            gFindMatch smSave(3), 0, lbcDPGroup
            If gLastFound(lbcDPGroup) >= 0 Then
                lbcDPGroup.ListIndex = gLastFound(lbcDPGroup)
                edcSpecDropDown.Text = lbcDPGroup.List(lbcDPGroup.ListIndex)
            Else
                If lbcDPGroup.ListCount > 1 Then
                    lbcDPGroup.ListIndex = 1
                    edcSpecDropDown.Text = lbcDPGroup.List(lbcDPGroup.ListIndex)
                Else
                    lbcDPGroup.ListIndex = -1
                    edcSpecDropDown.Text = ""
                End If
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case SORTINDEX
            edcSpecDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 4
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            If imSave(5) < 0 Then
                imSave(5) = 0
            End If
            edcSpecDropDown.Text = Trim$(Str$(imSave(5)))
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case ACTIVEINDEX
            If imSave(6) < 0 Then
                imSave(6) = 0    'All avails
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcActive.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcActive, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcActive_Paint
            pbcActive.Visible = True
            pbcActive.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetTime                        *
'*                                                     *
'*             Created:7/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from record to     *
'*                      working area                   *
'*                                                     *
'*******************************************************
Private Sub mGetTime()
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilMax As Integer
    For ilLoop = 1 To imMaxTDRows Step 1  'Row
        smTDSave(ilLoop, 1) = ""
        smTDShow(ilLoop, TDSTARTINDEX) = ""
        smTDSave(ilLoop, 2) = ""
        smTDShow(ilLoop, TDENDINDEX) = ""
        For ilMax = 1 To 7 Step 1
            smTDSave(ilLoop, 2 + ilMax) = ""
        Next ilMax
        For ilMax = 0 To 6 Step 1
            smTDShow(ilLoop, TDDAYINDEX + ilMax) = ""
        Next ilMax
        smTDSave(ilLoop, 10) = ""
        smTDShow(ilLoop, TDPCSPOTINDEX) = ""
    Next ilLoop
    imAnyZeroPct = False
    ilIndex = 1
    For ilLoop = 1 To imMaxTDRows Step 1  'Row
        If (tmRdf.iStartTime(0, ilLoop - 1) <> 1) Or (tmRdf.iStartTime(1, ilLoop - 1) <> 0) Then
            gUnpackTime tmRdf.iStartTime(0, ilLoop - 1), tmRdf.iStartTime(1, ilLoop - 1), "A", "1", smTDSave(ilIndex, 1)
            slStr = smTDSave(ilIndex, 1)
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcTD, slStr, tmTDCtrls(TDSTARTINDEX)
            smTDShow(ilIndex, TDSTARTINDEX) = tmTDCtrls(TDSTARTINDEX).sShow
            gUnpackTime tmRdf.iEndTime(0, ilLoop - 1), tmRdf.iEndTime(1, ilLoop - 1), "A", "1", smTDSave(ilIndex, 2)
            slStr = smTDSave(ilIndex, 2)
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcTD, slStr, tmTDCtrls(TDENDINDEX)
            smTDShow(ilIndex, TDENDINDEX) = tmTDCtrls(TDENDINDEX).sShow
            For ilMax = 1 To 7 Step 1
                smTDSave(ilIndex, 2 + ilMax) = tmRdf.sWkDays(ilLoop - 1, ilMax - 1)
            Next ilMax
            For ilMax = 0 To 6 Step 1
                If smTDSave(ilIndex, 3 + ilMax) = "Y" Then
                    slStr = "4"
                Else
                    slStr = "  "
                End If
                gSetShow pbcTD, slStr, tmTDCtrls(TDDAYINDEX + ilMax)
                smTDShow(ilIndex, TDDAYINDEX + ilMax) = tmTDCtrls(TDDAYINDEX + ilMax).sShow
            Next ilMax
            'gPDNToStr tmRdf.sSpotPct(ilLoop), 2, slStr
            If tmRdf.iSpotPct(ilLoop - 1) = 0 Then
                imAnyZeroPct = True
                slStr = "0"
                smTDSave(ilIndex, 10) = slStr
            Else
                slStr = gIntToStrDec(tmRdf.iSpotPct(ilLoop - 1), 2)
                If slStr = "0.00" Then
                    slStr = ""
                    smTDSave(ilIndex, 10) = ""
                Else
                    smTDSave(ilIndex, 10) = slStr
                    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
                End If
            End If
            gSetShow pbcTD, slStr, tmTDCtrls(TDPCSPOTINDEX)
            smTDShow(ilIndex, TDPCSPOTINDEX) = tmTDCtrls(TDPCSPOTINDEX).sShow
            ilIndex = ilIndex + 1
        End If
    Next ilLoop
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
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBTDCtrls = 1
    imLBLibCtrls = 1
    imFirstActivate = True
    igDPAltered = False
    imTerminate = False
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    imBoxNo = -1 'Initialize current Box to N/A
    imBypassSetting = False
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imBoxNo = -1
    imTDBoxNo = -1
    imLibBoxNo = -1
    imTDChg = False
    imAnyZeroPct = False
    'imMaxTDRows = UBound(tmRdf.iStartTime, 2)
    imMaxTDRows = UBound(smTDShow) 'Dimension of Rcf.iStartTime
    hmRdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "RDF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RDF.Btr)", Daypart
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)
    hmLtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLtf, "", sgDBPath & "LTF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RDF.Btr)", Daypart
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)
    Screen.MousePointer = vbHourglass  'Wait
    lbcAvail.Clear 'Force list box to be populated
    mAvailsPop
    If imTerminate Then
        Exit Sub
    End If
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    lbcDPGroup.Clear 'Force list box to be populated
    mDPGroupPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    rbcType(0).Value = True
    If igVefCode <= 0 Then  'If vehicle not defined then only allow times input
        'plcType.Visible = False
        lacType.Visible = False
        rbcType(0).Visible = False
        rbcType(1).Visible = False
    End If
    Daypart.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterModalForm Daypart
    mInitBox
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
    flTextHeight = pbcSpec.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSpec.Move 360, 555, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    'Daypart Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2175, fgBoxStH
    'Base Daypart
    gSetCtrl tmCtrls(BASEINDEX), 2220, tmCtrls(NAMEINDEX).fBoxY, 1425, fgBoxStH
    'Report Daypart
    gSetCtrl tmCtrls(RPTINDEX), 3660, tmCtrls(NAMEINDEX).fBoxY, 1410, fgBoxStH
    'Avails
    gSetCtrl tmCtrls(AVAILINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2175, fgBoxStH
    gSetCtrl tmCtrls(AVAILINDEX + 1), tmCtrls(AVAILINDEX).fBoxX, tmCtrls(AVAILINDEX).fBoxY, 2175, fgBoxStH
    tmCtrls(AVAILINDEX + 1).iReq = False
    'Override Hour flag
    gSetCtrl tmCtrls(HRSINDEX), 2220, tmCtrls(AVAILINDEX).fBoxY, 1425, fgBoxStH
    'Report Daypart
    gSetCtrl tmCtrls(OVERRIDEPRTINDEX), 3660, tmCtrls(AVAILINDEX).fBoxY, 1410, fgBoxStH
    'Daypart Group Name
    gSetCtrl tmCtrls(DPGROUPINDEX), 30, tmCtrls(AVAILINDEX).fBoxY + fgStDeltaY, 2175, fgBoxStH
    tmCtrls(DPGROUPINDEX).iReq = False
    'Sort Code
    gSetCtrl tmCtrls(SORTINDEX), 2220, tmCtrls(DPGROUPINDEX).fBoxY, 1425, fgBoxStH
    'Active/Dormant
    gSetCtrl tmCtrls(ACTIVEINDEX), 3660, tmCtrls(SORTINDEX).fBoxY, 1410, fgBoxStH
    
    plcTD.Move 615, 2130, pbcTD.Width + fgPanelAdj, pbcTD.Height + fgPanelAdj
    pbcTD.Move plcTD.Left + fgBevelX, plcTD.Top + fgBevelY
    pbcArrow.Move plcTD.Left - pbcArrow.Width - 15
    'Star/End Time
    gSetCtrl tmTDCtrls(TDSTARTINDEX), 30, 360, 960, fgBoxGridH
    gSetCtrl tmTDCtrls(TDENDINDEX), 975, tmTDCtrls(TDSTARTINDEX).fBoxY, 960, fgBoxGridH
    'Days of the week
    For ilLoop = 0 To 6 Step 1
        gSetCtrl tmTDCtrls(TDDAYINDEX + ilLoop), 1950 + 225 * (ilLoop), tmTDCtrls(TDSTARTINDEX).fBoxY, 210, fgBoxGridH
        tmTDCtrls(TDDAYINDEX + ilLoop).iReq = False
    Next ilLoop
    'Percent spots
    gSetCtrl tmTDCtrls(TDPCSPOTINDEX), 3525, tmTDCtrls(TDSTARTINDEX).fBoxY, 780, fgBoxGridH
    tmTDCtrls(TDPCSPOTINDEX).iReq = False
    plcLib.Move 1605, plcTD.Top, pbcLib.Width + fgPanelAdj, pbcLib.Height + fgPanelAdj
    pbcLib.Move plcLib.Left + fgBevelX, plcLib.Top + fgBevelY
    gSetCtrl tmLibCtrls(LIB1INDEX), 30, 30, 2400, fgBoxStH
    'Library 2
    gSetCtrl tmLibCtrls(LIB2INDEX), 30, tmLibCtrls(LIB1INDEX).fBoxY + fgStDeltaY, 2400, fgBoxStH
    tmLibCtrls(LIB2INDEX).iReq = False
    'Library 3
    gSetCtrl tmLibCtrls(LIB3INDEX), 30, tmLibCtrls(LIB2INDEX).fBoxY + fgStDeltaY, 2400, fgBoxStH
    tmLibCtrls(LIB3INDEX).iReq = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    For ilBoxNo = NAMEINDEX To ACTIVEINDEX
        Select Case ilBoxNo 'Branch on box type (control)
            Case NAMEINDEX
                slStr = smSave(1)
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case BASEINDEX
                If imSave(1) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case RPTINDEX
                If imSave(2) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case AVAILINDEX
                Select Case imSave(3)
                    Case 0  'All avails
                        slStr = "All avails"
                        smSave(2) = ""
                    Case 1  'Book into selected avail
                        slStr = smSave(2)
                    Case 2  'Exclused selected avail
                        slStr = "Ø" & smSave(2)    'alt + 0216 or 0248
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case AVAILINDEX + 1
            Case HRSINDEX    'Source Index
                If imSave(4) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case OVERRIDEPRTINDEX    'Override Print Rule Index
                If imSave(7) = 1 Then
                    slStr = "Daypart Name"
                ElseIf imSave(7) = 2 Then
                    slStr = "Avail Name"
                Else
                    slStr = "None"
                End If
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case DPGROUPINDEX
                slStr = Trim$(smSave(3))
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case SORTINDEX
                slStr = Trim$(Str$(imSave(5)))
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
            Case ACTIVEINDEX
                Select Case imSave(6)
                    Case 0  'Active
                        slStr = "Active"
                    Case 1  'Book into selected avail
                        slStr = "Dormant"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
    If (smLibSave(1) = "") And (smLibSave(2) = "") And (smLibSave(3) = "") Then
        For ilRowNo = 1 To imMaxTDRows Step 1
            For ilBoxNo = TDSTARTINDEX To TDPCSPOTINDEX Step 1
                Select Case ilBoxNo 'Branch on box type (control)
                    Case TDSTARTINDEX 'Vehicle
                        slStr = smTDSave(ilRowNo, TDSTARTINDEX)
                        If gValidTime(slStr) Then
                            If slStr <> "" Then
                                slStr = gFormatTime(slStr, "A", "1")
                            End If
                            gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                            smTDShow(ilRowNo, TDSTARTINDEX) = tmTDCtrls(ilBoxNo).sShow
                        Else
                            smTDShow(ilRowNo, TDSTARTINDEX) = ""
                        End If
                    Case TDENDINDEX 'Vehicle
                        slStr = smTDSave(ilRowNo, TDENDINDEX)
                        If gValidTime(slStr) Then
                            If slStr <> "" Then
                                slStr = gFormatTime(slStr, "A", "1")
                            End If
                            gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                            smTDShow(ilRowNo, TDENDINDEX) = tmTDCtrls(ilBoxNo).sShow
                        Else
                            smTDShow(ilRowNo, TDENDINDEX) = tmTDCtrls(ilBoxNo).sShow
                        End If
                    Case TDDAYINDEX To TDDAYINDEX + 6 'Day
                        slStr = smTDSave(ilRowNo, ilBoxNo)
                        If slStr = "Y" Then
                            slStr = "4"
                        Else
                            slStr = " "
                        End If
                        gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                        smTDShow(ilRowNo, ilBoxNo) = tmTDCtrls(ilBoxNo).sShow
                    Case TDPCSPOTINDEX 'Percent Spot
                        slStr = smTDSave(ilRowNo, TDPCSPOTINDEX)
                        If Not imAnyZeroPct Then
                            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
                        End If
                        gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                        smTDShow(ilRowNo, TDPCSPOTINDEX) = tmTDCtrls(ilBoxNo).sShow
                End Select
            Next ilBoxNo
        Next ilRowNo
        tmLibCtrls(1).sShow = ""
        tmLibCtrls(2).sShow = ""
        tmLibCtrls(3).sShow = ""
    Else
        For ilLoop = 1 To 3 Step 1
            slStr = smLibSave(ilLoop)
            gSetShow pbcLib, slStr, tmLibCtrls(ilBoxNo)
        Next ilLoop
        For ilRowNo = 1 To 7 Step 1
            For ilBoxNo = 1 To 10 Step 1
                smTDShow(ilRowNo, ilBoxNo) = ""
            Next ilBoxNo
        Next ilRowNo
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibEnableBox                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mLibEnableBox(ilBoxNo As Integer)
'
'   mTDEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBLibCtrls Or ilBoxNo > UBound(tmLibCtrls) Then
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcLib.Left - pbcArrow.Width - 15, plcLib.Top + tmLibCtrls(1).fBoxY + (ilBoxNo - 1) * (fgBoxStH + 15) + fgBoxStH \ 2
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case LIB1INDEX, LIB1INDEX + 1, LIB1INDEX + 2 'Library name
            If ilBoxNo = LIB1INDEX Then
                mLibPop
            End If
            If imTerminate Then
                Exit Sub
            End If
            lbcLibNm(ilBoxNo - LIB1INDEX).Height = gListBoxHeight(lbcLibNm(ilBoxNo - LIB1INDEX).ListCount, 6)
            edcLibDropDown.Width = tmLibCtrls(ilBoxNo).fBoxW - cmcLibDropDown.Width
            edcLibDropDown.MaxLength = 20
            If ilBoxNo = LIB1INDEX Then
                gMoveFormCtrl pbcLib, edcLibDropDown, tmLibCtrls(ilBoxNo).fBoxX, tmLibCtrls(ilBoxNo).fBoxY
            ElseIf ilBoxNo = LIB1INDEX + 1 Then
                gMoveFormCtrl pbcLib, edcLibDropDown, tmLibCtrls(ilBoxNo).fBoxX, tmLibCtrls(ilBoxNo).fBoxY
            Else
                gMoveFormCtrl pbcLib, edcLibDropDown, tmLibCtrls(ilBoxNo).fBoxX, tmLibCtrls(ilBoxNo).fBoxY
            End If
            cmcLibDropDown.Move edcLibDropDown.Left + edcLibDropDown.Width, edcLibDropDown.Top
            'If imRowNo - vbcRateCard.Value <= vbcRateCard.LargeChange \ 2 Then
                lbcLibNm(ilBoxNo - LIB1INDEX).Move edcLibDropDown.Left, edcLibDropDown.Top + edcLibDropDown.Height
            'Else
            '    lbcLibNm(ilBoxNo - LIB1INDEX).Move edcLibDropDown.Left, edcLibDropDown.Top - lbcLibNm(ilBoxNo - LIB1INDEX).Height
            'End If
            imChgMode = True
            gFindMatch smLibSave(1 + ilBoxNo - LIB1INDEX), 0, lbcLibNm(ilBoxNo - LIB1INDEX)
            If gLastFound(lbcLibNm(ilBoxNo - LIB1INDEX)) >= 0 Then
                lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex = gLastFound(lbcLibNm(ilBoxNo - LIB1INDEX))
            Else
                lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex = -1
            End If
            If lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex >= 0 Then
                edcLibDropDown.Text = lbcLibNm(ilBoxNo - LIB1INDEX).List(lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex)
            Else
                edcLibDropDown.Text = ""
            End If
            imChgMode = False
            edcLibDropDown.SelStart = 0
            edcLibDropDown.SelLength = Len(edcLibDropDown.Text)
            edcLibDropDown.Visible = True
            cmcLibDropDown.Visible = True
            edcLibDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLibPop()
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    'gUnpackDate tgRcfI.iStartDate(0), tgRcfI.iStartDate(1), slStr
    If igVefCode <= 0 Then
        lbcLibNm(0).Clear
        lbcLibNm(1).Clear
        lbcLibNm(2).Clear
        Exit Sub
    End If
    slStr = Format$(gNow(), "m/d/yy")
    llStartDate = gDateValue(slStr)
    llEndDate = 0
    'ilRet = gPopLibByDateBox(Daypart, 1, "A", llStartDate, llEndDate, igVefCode, lbcLibNm(0), lbcLibNmCode)
    ilRet = gPopLibByDateBox(Daypart, 1, "A", llStartDate, llEndDate, igVefCode, lbcLibNm(0), tmLibNmCode(), smLibNmCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLibPopErr
        gCPErrorMsg ilRet, "mLibPop (gPopLibByDateBox: Library)", Daypart
        On Error GoTo 0
        lbcLibNm(0).AddItem "[None]", 0
        lbcLibNm(1).Clear
        lbcLibNm(2).Clear
        For ilLoop = 0 To lbcLibNm(0).ListCount - 1 Step 1
            lbcLibNm(1).AddItem lbcLibNm(0).List(ilLoop)
            lbcLibNm(2).AddItem lbcLibNm(0).List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibSetFocus                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mLibSetFocus(ilBoxNo As Integer)
'
'   mLibSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBLibCtrls Or ilBoxNo > UBound(tmLibCtrls) Then
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcLib.Left - pbcArrow.Width - 15, plcLib.Top + tmLibCtrls(1).fBoxY + (ilBoxNo - 1) * (fgBoxStH + 15) + fgBoxStH \ 2
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case LIB1INDEX, LIB1INDEX + 1, LIB1INDEX + 2 'Library name
            edcLibDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTDSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mLibSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    If ilBoxNo < imLBLibCtrls Or ilBoxNo > UBound(tmLibCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LIB1INDEX, LIB1INDEX + 1, LIB1INDEX + 2 'Library name
            lbcLibNm(ilBoxNo - LIB1INDEX).Visible = False
            edcLibDropDown.Visible = False
            cmcLibDropDown.Visible = False
            If lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex > 0 Then
                slStr = lbcLibNm(ilBoxNo - LIB1INDEX).List(lbcLibNm(ilBoxNo - LIB1INDEX).ListIndex)
            Else
                slStr = ""
            End If
            If smLibSave(1 + ilBoxNo - LIB1INDEX) <> slStr Then
                imRdfChg = True
            End If
            smLibSave(1 + ilBoxNo - LIB1INDEX) = slStr
            gSetShow pbcLib, slStr, tmLibCtrls(ilBoxNo)
    End Select
    mSetCommands
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
'   mMoveCtrlToRec iTest
'   Where:
'
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilLib As Integer
    Dim ilLibDefined As Integer
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    '4/12/11: Temporarily set to N
    tmRdf.sOverridePrtRules = "N"
    
    tmRdf.sName = smSave(1)    'Name
    If imSave(1) = 0 Then
        tmRdf.sBase = "Y"
    Else
        tmRdf.sBase = "N"
    End If
    If imSave(2) = 0 Then
        tmRdf.sReport = "Y"
    Else
        tmRdf.sReport = "N"
    End If
    Select Case imSave(3)
        Case 0  'All avails
            tmRdf.sInOut = "N"
        Case 1  'Book into avail
            tmRdf.sInOut = "I"
        Case 2  'Exclude avail
            tmRdf.sInOut = "O"
    End Select
    If (imSave(3) = 1) Or (imSave(3) = 2) Then
        gFindMatch smSave(2), 1, lbcAvail
        ilIndex = gLastFound(lbcAvail)
        slNameCode = tmAvailCode(ilIndex - 1).sKey 'lbcAvailCode.List(ilIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Daypart
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmRdf.ianfCode = CInt(slCode)
    Else
        tmRdf.ianfCode = 0
    End If
    If imSave(4) = 1 Then
        tmRdf.sTimeOver = "N"
    Else
        tmRdf.sTimeOver = "Y"
    End If
    If imSave(7) = 1 Then
        tmRdf.sOverridePrtRules = "D"
    ElseIf (imSave(7) = 2) And (imSave(3) > 0) Then
        tmRdf.sOverridePrtRules = "A"
    Else
        tmRdf.sOverridePrtRules = "N"
    End If
    If smSave(3) <> "" And smSave(3) <> "[New]" And smSave(3) <> "[None]" Then
        gFindMatch smSave(3), 1, lbcDPGroup
        ilIndex = gLastFound(lbcDPGroup)
        slNameCode = tmDPGroupCode(ilIndex - 2).sKey 'lbcAvailCode.List(ilIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Daypart
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmRdf.iDPGroupMnfCode = CInt(slCode)
    Else
        tmRdf.iDPGroupMnfCode = 0
    End If
    tmRdf.iSortCode = imSave(5)
    If imSave(6) = 1 Then
        tmRdf.sState = "D"
    Else
        tmRdf.sState = "A"
    End If
    ilLibDefined = False
    For ilLib = 1 To 3 Step 1
        If Trim$(smLibSave(ilLib)) <> "" Then
            gFindMatch smLibSave(ilLib), 0, lbcLibNm(ilLib)
            ilIndex = gLastFound(lbcLibNm(ilLib))
            slNameCode = tmLibNmCode(ilIndex).sKey 'lbcLibNmCode.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", Daypart
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmRdf.iLtfCode(ilLib - 1) = CInt(slCode)
            ilLibDefined = True
        Else
            tmRdf.iLtfCode(ilLib - 1) = 0
        End If
    Next ilLib
    If ilLibDefined Then
        For ilRow = 1 To imMaxTDRows Step 1
            tmRdf.iStartTime(0, ilRow - 1) = 1
            tmRdf.iStartTime(1, ilRow - 1) = 0
            tmRdf.iEndTime(0, ilRow - 1) = 1
            tmRdf.iEndTime(1, ilRow - 1) = 0
        Next ilRow
        For ilCol = 1 To 7 Step 1
            For ilRow = 1 To imMaxTDRows Step 1
                tmRdf.sWkDays(ilRow - 1, ilCol - 1) = "N"
            Next ilRow
        Next ilCol
        For ilCol = 1 To imMaxTDRows Step 1
            'slStr = ""
            'gStrToPDN slStr, 2, 3, tmRdf.sSpotPct(ilCol)
            tmRdf.iSpotPct(ilCol - 1) = 0
        Next ilCol
    Else
        mSetTime
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
    Dim ilTest As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilLib As Integer
    Dim ilLoop As Integer
    
    smSave(1) = Trim$(tmRdf.sName)
    If tmRdf.sBase = "Y" Then
        imSave(1) = 0
    ElseIf tmRdf.sBase = "N" Then
        imSave(1) = 1
    Else
        imSave(1) = 1
    End If
    If tmRdf.sReport = "Y" Then
        imSave(2) = 0
    ElseIf tmRdf.sReport = "N" Then
        imSave(2) = 1
    Else
        imSave(2) = 1
    End If
    If tmRdf.sInOut = "N" Then
        imSave(3) = 0
    ElseIf tmRdf.sInOut = "I" Then
        imSave(3) = 1
    ElseIf tmRdf.sInOut = "O" Then
        imSave(3) = 2
    Else
        imSave(3) = -1
    End If
    smSave(2) = ""
    If (tmRdf.sInOut = "I") Or (tmRdf.sInOut = "O") Then
        slRecCode = Trim$(Str$(tmRdf.ianfCode))
        For ilTest = 0 To UBound(tmAvailCode) - 1 Step 1 'lbcAvailCode.ListCount - 1 Step 1
            slNameCode = tmAvailCode(ilTest).sKey  'lbcAvailCode.List(ilTest)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Daypart
            On Error GoTo 0
            If slRecCode = slCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smSave(2))
                On Error GoTo mMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 1)", Daypart
                On Error GoTo 0
                Exit For
            End If
        Next ilTest
        smSave(2) = Trim$(smSave(2))
    End If
    If tmRdf.sTimeOver = "N" Then
        imSave(4) = 1
    Else
        imSave(4) = 0
    End If
    If tmRdf.sOverridePrtRules = "D" Then
        imSave(7) = 1
    ElseIf tmRdf.sOverridePrtRules = "A" Then
        imSave(7) = 2
    Else
        imSave(7) = 0
    End If
    smSave(3) = ""
    smOrigDPGroup = ""
    If tmRdf.iDPGroupMnfCode > 0 Then
        For ilLoop = 0 To UBound(tmDPGroupCode) - 1 Step 1 'lbcAnnCode.ListCount - 1 Step 1
            slNameCode = tmDPGroupCode(ilLoop).sKey  'lbcAnnCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmRdf.iDPGroupMnfCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smSave(3))
                imChgMode = True
                lbcDPGroup.ListIndex = ilLoop + 2
                imChgMode = False
                smOrigDPGroup = smSave(3)
                Exit For
            End If
        Next ilLoop
    End If

    imSave(5) = tmRdf.iSortCode
    If tmRdf.sState = "D" Then
        imSave(6) = 1
    Else
        imSave(6) = 0
    End If
    For ilLib = 1 To 3 Step 1
        smLibSave(ilLib) = ""
        If tmRdf.iLtfCode(ilLib - 1) > 0 Then
            tmLtfSrchKey.iCode = tmRdf.iLtfCode(ilLib - 1)
            ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smLibSave(ilLib) = Trim$(tmLtf.sName)
            End If
        End If
    Next ilLib
    mGetTime
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
    If smSave(1) <> "" Then    'Test name
        slStr = Trim$(smSave(1))
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                slStr = Trim$(smSave(1))
                If slStr = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Daypart already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    smSave(1) = Trim$(tmRdf.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo, True
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
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    ilFilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    imPopReqd = False
    'ilRet = gIMoveListBox(Daypart, cbcSelect, lbcDPNameCode, "Rdf.Btr", gFieldOffset("Rdf", "RdfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Daypart, cbcSelect, tmDPNameCode(), smDPNameCodeTag, "Rdf.Btr", gFieldOffset("Rdf", "RdfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox: Daypart)", Daypart
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
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
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

    slNameCode = tmDPNameCode(ilSelectIndex - 1).sKey  'lbcDPNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", Daypart
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmRdfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Daypart
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
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim slSyncDate As String
    Dim slSyncTime As String
    mSetShow imBoxNo
    mTDSetShow imTDBoxNo
    mLibSetShow imLibBoxNo
    If mTestSaveFields(SHOWMSG, True) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    igDPAltered = True
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "RDF.Btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec
        If imSelectedIndex = 0 Then 'New selected
            tmRdf.iCode = 0  'Autoincrement
            tmRdf.sContrRef = "N"
            tmRdf.iRemoteID = tgUrf(0).iRemoteUserID
            tmRdf.iAutoCode = tmRdf.iCode
            ilRet = btrInsert(hmRdf, tmRdf, imRdfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Daypart)"
        Else 'Old record-Update
            tmRdf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmRdf.iSyncDate(0), tmRdf.iSyncDate(1)
            gPackTime slSyncTime, tmRdf.iSyncTime(0), tmRdf.iSyncTime(1)
            ilRet = btrUpdate(hmRdf, tmRdf, imRdfRecLen)
            slMsg = "mSaveRec (btrUpdate:Daypart)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Daypart
    On Error GoTo 0
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmRdfSrchKey.iCode = tmRdf.iCode
            'ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Daypart)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, Daypart
            'On Error GoTo 0
            tmRdf.iRemoteID = tgUrf(0).iRemoteUserID
            tmRdf.iAutoCode = tmRdf.iCode
            tmRdf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmRdf.iSyncDate(0), tmRdf.iSyncDate(1)
            gPackTime slSyncTime, tmRdf.iSyncTime(0), tmRdf.iSyncTime(1)
            ilRet = btrUpdate(hmRdf, tmRdf, imRdfRecLen)
            slMsg = "mSaveRec (btrUpdate:Daypart)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Daypart
        On Error GoTo 0
        tgMRdf(UBound(tgMRdf)) = tmRdf
        'ReDim Preserve tgMRdf(1 To UBound(tgMRdf) + 1) As RDF
        ReDim Preserve tgMRdf(0 To UBound(tgMRdf) + 1) As RDF
        'Sort by code so that binary search can be used
        'If UBound(tgMRdf) - 1 > 1 Then
        If UBound(tgMRdf) > 0 Then
            'ArraySortTyp fnAV(tgMRdf(), 1), UBound(tgMRdf) - 1, 0, LenB(tgMRdf(1)), 0, -1, 0
            ArraySortTyp fnAV(tgMRdf(), 0), UBound(tgMRdf), 0, LenB(tgMRdf(0)), 0, -1, 0
        End If
    Else
        'TTP 6563
        'ilRet = gBinarySearchRdf(tmAdf.iCode)
        ilRet = gBinarySearchRdf(tmRdf.iCode)
        If ilRet <> -1 Then
            tgMRdf(ilRet) = tmRdf
        End If
    End If
    'If imSelectedIndex <> 0 Then
    '    lbcDPNameCode.RemoveItem imSelectedIndex - 1
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = Trim$(tmRdf.sName)
    'cbcSelect.AddItem slName
    'slName = tmRdf.sName
    'slName = slName + "\" + LTrim$(Str$(tmRdf.iCode))
    'lbcDPNameCode.AddItem slName
    'cbcSelect.AddItem "[New]", 0
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
    If mTestSaveFields(NOMSG, True) = YES Then  'No Then
        If (ilAltered = YES) Or (imTDChg = True) Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & smSave(1)
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
Private Sub mSetChg(ilBoxNo As Integer, ilUseSave As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    Dim slInitStr As String
    If ilBoxNo < imLBCtrls Or (ilBoxNo > UBound(tmCtrls)) Then
        'mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            If ilUseSave Then
                slStr = smSave(1)
            Else
                slStr = edcSpecDropDown.Text
            End If
            gSetChgFlagStr tmRdf.sName, slStr, tmCtrls(ilBoxNo)
        Case BASEINDEX
            If imSave(1) = 0 Then
                slStr = "Y"
            ElseIf imSave(1) = 1 Then
                slStr = "N"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sBase, slStr, tmCtrls(ilBoxNo)
        Case RPTINDEX
            If imSave(2) = 0 Then
                slStr = "Y"
            ElseIf imSave(2) = 1 Then
                slStr = "N"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sReport, slStr, tmCtrls(ilBoxNo)
        Case AVAILINDEX
            If imSave(3) = 0 Then
                slStr = "N"
            ElseIf imSave(3) = 1 Then
                slStr = "I"
            ElseIf imSave(3) = 2 Then
                slStr = "O"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sInOut, slStr, tmCtrls(ilBoxNo)
        Case AVAILINDEX + 1
            gSetChgFlag smSave(2), lbcAvail, tmCtrls(ilBoxNo)
        Case HRSINDEX
            If imSave(4) = 0 Then
                slStr = "Y"
            ElseIf imSave(4) = 1 Then
                slStr = "N"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sTimeOver, slStr, tmCtrls(ilBoxNo)
        Case OVERRIDEPRTINDEX
            If imSave(7) = 0 Then
                slStr = "N"
            ElseIf imSave(7) = 1 Then
                slStr = "D"
            ElseIf imSave(7) = 2 Then
                slStr = "A"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sOverridePrtRules, slStr, tmCtrls(ilBoxNo)
        Case DPGROUPINDEX
            If ilUseSave Then
                slStr = Trim$(smSave(3))
            Else
                slStr = Trim$(edcSpecDropDown.Text)
            End If
            gSetChgFlagStr smOrigDPGroup, slStr, tmCtrls(ilBoxNo)
        Case SORTINDEX
            If ilUseSave Then
                slStr = Trim$(Str$(imSave(5)))
            Else
                slStr = edcSpecDropDown.Text
            End If
            slInitStr = Trim$(Str$(tmRdf.iSortCode))
            gSetChgFlagStr slInitStr, slStr, tmCtrls(ilBoxNo)
        Case ACTIVEINDEX
            If imSave(6) = 0 Then
                slStr = "A"
            ElseIf imSave(6) = 1 Then
                slStr = "D"
            Else
                slStr = " "
            End If
            gSetChgFlagStr tmRdf.sState, slStr, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    'Update button set if all mandatory fields have data and any field altered
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestSaveFields(NOMSG, False) = YES) And ((ilAltered = YES) Or (imTDChg = True)) Then
        If imUpdateAllowed Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    'If ilAltered Then
    '    cmcUndo.Enabled = True
    'Else
    '    cmcUndo.Enabled = False
    'End If
    'Erase button set if any field contains a value or change mode
    'If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
    '    cmcErase.Enabled = True
    'Else
    '    cmcErase.Enabled = False
    'End If
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
'*             Created:6/28/93       By:D. LeVine      *
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

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.SetFocus
        Case BASEINDEX
            pbcHrs.SetFocus
        Case RPTINDEX
            pbcHrs.SetFocus
        Case AVAILINDEX
            pbcAvType.SetFocus
        Case AVAILINDEX + 1
            edcSpecDropDown.SetFocus
        Case HRSINDEX    'Source Index
            pbcHrs.SetFocus
        Case DPGROUPINDEX
            edcSpecDropDown.SetFocus
        Case SORTINDEX
            edcSpecDropDown.SetFocus
        Case ACTIVEINDEX
            pbcActive.SetFocus
    End Select
    mSetCommands
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
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    pbcArrow.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Visible = False
            slStr = Trim$(edcSpecDropDown.Text)
            smSave(1) = slStr
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case BASEINDEX
            pbcHrs.Visible = False
            If imSave(1) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case RPTINDEX
            pbcHrs.Visible = False
            If imSave(2) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case AVAILINDEX 'Avail
            pbcAvType.Visible = False  'Set visibility
            Select Case imSave(3)
                Case 0  'All avails
                    slStr = "All avails"
                    smSave(2) = ""
                Case 1  'Book into selected avail
                    slStr = smSave(2)
                Case 2  'Excluse selected avail
                    slStr = "Ø" & smSave(2)    'Alt+0216 or 0248
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmCtrls(AVAILINDEX)
            If imSave(3) <= 0 Then
                If imSave(7) = 2 Then
                    pbcSpec.Cls
                    imSave(7) = 0
                    gSetShow pbcSpec, "None", tmCtrls(OVERRIDEPRTINDEX)
                End If
            End If
        Case AVAILINDEX + 1
            lbcAvail.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            If lbcAvail.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcAvail.List(lbcAvail.ListIndex)
            End If
            smSave(2) = slStr
            Select Case imSave(3)
                Case 0  'All avails
                    slStr = "All avails"
                    smSave(2) = ""
                Case 1  'Book into selected avail
                    slStr = smSave(2)
                Case 2  'Exclused selected avail
                    slStr = "Ø" & smSave(2)    'alt + 0216 or 0248
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmCtrls(AVAILINDEX)
        Case HRSINDEX 'Hour flag
            pbcHrs.Visible = False
            If imSave(4) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case OVERRIDEPRTINDEX 'Hour flag
            pbcOverridePrt.Visible = False
            If imSave(7) = 1 Then
                slStr = "Daypart Name"
            ElseIf imSave(7) = 2 Then
                slStr = "Avail Name"
            Else
                slStr = "None"
            End If
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case DPGROUPINDEX
            lbcDPGroup.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            If lbcDPGroup.ListIndex <= 1 Then
                slStr = ""
            Else
                slStr = lbcDPGroup.List(lbcDPGroup.ListIndex)
            End If
            smSave(3) = slStr
            gSetShow pbcSpec, slStr, tmCtrls(DPGROUPINDEX)
        Case SORTINDEX
            edcSpecDropDown.Visible = False
            slStr = Trim$(edcSpecDropDown.Text)
            imSave(5) = Val(slStr)
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
        Case ACTIVEINDEX
            pbcActive.Visible = False  'Set visibility
            Select Case imSave(6)
                Case 0  'Active
                    slStr = "Active"
                Case 1  'Book into selected avail
                    slStr = "Dormant"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetTime                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from working area  *
'*                      to the record                  *
'*                                                     *
'*******************************************************
Private Sub mSetTime()
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilMax As Integer
    Dim ilCount As Integer
    For ilLoop = 1 To imMaxTDRows Step 1
        slStr = ""
        gPackTime slStr, tmRdf.iStartTime(0, ilLoop - 1), tmRdf.iStartTime(1, ilLoop - 1)
        tmRdf.iStartTime(0, ilLoop - 1) = 1 'used as indicator that time not defined
        tmRdf.iStartTime(1, ilLoop - 1) = 0
        gPackTime slStr, tmRdf.iEndTime(0, ilLoop - 1), tmRdf.iEndTime(1, ilLoop - 1)
        tmRdf.iEndTime(0, ilLoop - 1) = 1
        tmRdf.iEndTime(1, ilLoop - 1) = 0
        For ilMax = 1 To 7 Step 1
            tmRdf.sWkDays(ilLoop - 1, ilMax - 1) = ""
        Next ilMax
        'slStr = ""
        'gStrToPDN slStr, 2, 3, tmRdf.sSpotPct(ilLoop)
        tmRdf.iSpotPct(ilLoop - 1) = 0
    Next ilLoop
    ilCount = 0
    For ilLoop = 1 To imMaxTDRows Step 1
        If smTDSave(ilLoop, 1) <> "" Then
            ilCount = ilCount + 1
        End If
    Next ilLoop
    If ilCount = 0 Then
        Exit Sub
    End If
    ilIndex = imMaxTDRows + 1 - ilCount 'Start index
    For ilLoop = 1 To imMaxTDRows Step 1  'Row
        If smTDSave(ilLoop, 1) <> "" Then
            gPackTime smTDSave(ilLoop, 1), tmRdf.iStartTime(0, ilIndex - 1), tmRdf.iStartTime(1, ilIndex - 1)
            If smTDSave(ilLoop, 2) <> "" Then
                gPackTime smTDSave(ilLoop, 2), tmRdf.iEndTime(0, ilIndex - 1), tmRdf.iEndTime(1, ilIndex - 1)
            End If
            For ilMax = 1 To 7 Step 1
                If (StrComp(Trim$(smTDSave(ilLoop, 2 + ilMax)), "Y", 1) = 0) Or (StrComp(Trim$(smTDSave(ilLoop, 2 + ilMax)), "N", 1) = 0) Then
                    tmRdf.sWkDays(ilIndex - 1, ilMax - 1) = smTDSave(ilLoop, 2 + ilMax)
                Else
                    tmRdf.sWkDays(ilIndex - 1, ilMax - 1) = "N"
                End If
            Next ilMax
            slStr = smTDSave(ilLoop, 10)
            If imAnyZeroPct Then
                tmRdf.iSpotPct(ilIndex - 1) = 0
            Else
                'gStrToPDN slStr, 2, 3, tmRdf.sSpotPct(ilIndex)
                tmRdf.iSpotPct(ilIndex - 1) = gStrDecToLong(slStr, 2)
            End If
            ilIndex = ilIndex + 1
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTDEnableBox                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mTDEnableBox(ilBoxNo As Integer)
'
'   mTDEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slTotalPct As String
    Dim ilBlankPct As Integer
    If ilBoxNo < imLBTDCtrls Or ilBoxNo > UBound(tmTDCtrls) Then
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcTD.Left - pbcArrow.Width - 15, plcTD.Top + tmTDCtrls(1).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case TDSTARTINDEX 'Start/End
            edcTDStartEnd.MaxLength = 10
            gMoveTableCtrl pbcTD, edcTDStartEnd, tmTDCtrls(TDSTARTINDEX).fBoxX, tmTDCtrls(TDSTARTINDEX).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15)
            cmcTDStartEnd.Move edcTDStartEnd.Left + edcTDStartEnd.Width, edcTDStartEnd.Top
            If imTDRowNo <= 2 Then
                plcTme.Move edcTDStartEnd.Left, edcTDStartEnd.Top + edcTDStartEnd.Height
            Else
                plcTme.Move edcTDStartEnd.Left, edcTDStartEnd.Top - plcTme.Height
            End If
            edcTDStartEnd.Text = smTDSave(imTDRowNo, TDSTARTINDEX)
            plcTme.Visible = False
            edcTDStartEnd.Visible = True  'Set visibility
            cmcTDStartEnd.Visible = True
            edcTDStartEnd.SetFocus
        Case TDENDINDEX 'Start/End
            edcTDStartEnd.MaxLength = 10
            gMoveTableCtrl pbcTD, edcTDStartEnd, tmTDCtrls(TDENDINDEX).fBoxX - cmcTDStartEnd.Width, tmTDCtrls(TDENDINDEX).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15)
            cmcTDStartEnd.Move edcTDStartEnd.Left + edcTDStartEnd.Width, edcTDStartEnd.Top
            If imTDRowNo <= 2 Then
                plcTme.Move edcTDStartEnd.Left, edcTDStartEnd.Top + edcTDStartEnd.Height
            Else
                plcTme.Move edcTDStartEnd.Left, edcTDStartEnd.Top - plcTme.Height
            End If
            edcTDStartEnd.Text = smTDSave(imTDRowNo, TDENDINDEX)
            plcTme.Visible = False
            edcTDStartEnd.Visible = True  'Set visibility
            cmcTDStartEnd.Visible = True
            edcTDStartEnd.SetFocus
        Case TDDAYINDEX To TDDAYINDEX + 6 'Day index
            gMoveTableCtrl pbcTD, pbcTDDays, tmTDCtrls(ilBoxNo).fBoxX, tmTDCtrls(ilBoxNo).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15)
            If smTDSave(imTDRowNo, ilBoxNo) = "Y" Then
                ckcTDDay.Value = vbChecked
            Else
                ckcTDDay.Value = vbUnchecked
            End If
            pbcTDDays.Visible = True  'Set visibility
            ckcTDDay.SetFocus
        Case TDPCSPOTINDEX 'Percent spot index
            gMoveTableCtrl pbcTD, edcTDPCSpot, tmTDCtrls(TDPCSPOTINDEX).fBoxX, tmTDCtrls(TDPCSPOTINDEX).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15)
            imAnyZeroPct = False
            If (smTDSave(imTDRowNo, TDPCSPOTINDEX) = "") Then
                slTotalPct = "0"
                ilBlankPct = False
                For ilLoop = 1 To imTDRowNo - 1 Step 1
                    If smTDSave(ilLoop, TDPCSPOTINDEX) = "" Then
                        ilBlankPct = True
                        Exit For
                    Else
                        If gCompNumberStr(smTDSave(ilLoop, TDPCSPOTINDEX), "0") = 0 Then
                            imAnyZeroPct = True
                        Else
                            slTotalPct = gAddStr(smTDSave(ilLoop, TDPCSPOTINDEX), slTotalPct)
                        End If
                    End If
                Next ilLoop
                If ilBlankPct Then
                    edcTDPCSpot.Text = ""
                Else
                    If imAnyZeroPct Then
                        edcTDPCSpot.Text = "0"
                    Else
                        edcTDPCSpot.Text = gSubStr("100.00", slTotalPct)
                    End If
                End If
            Else
                edcTDPCSpot.Text = smTDSave(imTDRowNo, TDPCSPOTINDEX)
            End If
            edcTDPCSpot.Visible = True  'Set visibility
            edcTDPCSpot.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTDSetFocus                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mTDSetFocus(ilBoxNo As Integer)
'
'   mTDSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBTDCtrls Or ilBoxNo > UBound(tmTDCtrls) Then
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcTD.Left - pbcArrow.Width - 15, plcTD.Top + tmTDCtrls(1).fBoxY + (imTDRowNo - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case TDSTARTINDEX 'Start/End
            edcTDStartEnd.SetFocus
        Case TDENDINDEX 'Start/End
            edcTDStartEnd.SetFocus
        Case TDDAYINDEX To TDDAYINDEX + 6 'Day index
            ckcTDDay.SetFocus
        Case TDPCSPOTINDEX 'Percent spot index
            edcTDPCSpot.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTDSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mTDSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    If ilBoxNo < imLBTDCtrls Or ilBoxNo > UBound(tmTDCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case TDSTARTINDEX 'Vehicle
            cmcTDStartEnd.Visible = False
            plcTme.Visible = False
            edcTDStartEnd.Visible = False  'Set visibility
            slStr = edcTDStartEnd.Text
            If gValidTime(slStr) Then
                If slStr <> "" Then
                    slStr = gFormatTime(slStr, "A", "1")
                End If
                gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                smTDShow(imTDRowNo, TDSTARTINDEX) = tmTDCtrls(ilBoxNo).sShow
                If smTDSave(imTDRowNo, TDSTARTINDEX) <> edcTDStartEnd.Text Then
                    imTDChg = True
                End If
                smTDSave(imTDRowNo, TDSTARTINDEX) = edcTDStartEnd.Text
            Else
                Beep
                edcTDStartEnd.Text = smTDSave(imTDRowNo, TDSTARTINDEX)
            End If
        Case TDENDINDEX 'Vehicle
            cmcTDStartEnd.Visible = False
            plcTme.Visible = False
            edcTDStartEnd.Visible = False  'Set visibility
            edcTDStartEnd.Visible = False  'Set visibility
            slStr = edcTDStartEnd.Text
            If gValidTime(slStr) Then
                If slStr <> "" Then
                    slStr = gFormatTime(slStr, "A", "1")
                End If
                gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
                smTDShow(imTDRowNo, TDENDINDEX) = tmTDCtrls(ilBoxNo).sShow
                If smTDSave(imTDRowNo, TDENDINDEX) <> edcTDStartEnd.Text Then
                    imTDChg = True
                End If
                smTDSave(imTDRowNo, TDENDINDEX) = edcTDStartEnd.Text
            Else
                Beep
                edcTDStartEnd.Text = smTDSave(imTDRowNo, TDENDINDEX)
            End If
        Case TDDAYINDEX To TDDAYINDEX + 6 'Day
            pbcTDDays.Visible = False  'Set visibility
            If ckcTDDay.Value = vbChecked Then
                slStr = "4"
                If smTDSave(imTDRowNo, ilBoxNo) <> "Y" Then
                    imTDChg = True
                End If
                smTDSave(imTDRowNo, ilBoxNo) = "Y"
            Else
                slStr = "  "
                If smTDSave(imTDRowNo, ilBoxNo) <> "N" Then
                    imTDChg = True
                End If
                smTDSave(imTDRowNo, ilBoxNo) = "N"
            End If
            gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
            smTDShow(imTDRowNo, ilBoxNo) = tmTDCtrls(ilBoxNo).sShow
        Case TDPCSPOTINDEX 'Percent Spot
            edcTDPCSpot.Visible = False  'Set visibility
            slStr = edcTDPCSpot.Text
            If Not imAnyZeroPct Then
                If (slStr = "") Or (slStr = "0") Then
                    imAnyZeroPct = True
                Else
                    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
                End If
            End If
            gSetShow pbcTD, slStr, tmTDCtrls(ilBoxNo)
            smTDShow(imTDRowNo, TDPCSPOTINDEX) = tmTDCtrls(ilBoxNo).sShow
            If smTDSave(imTDRowNo, TDPCSPOTINDEX) <> edcTDPCSpot.Text Then
                imTDChg = True
            End If
            smTDSave(imTDRowNo, TDPCSPOTINDEX) = edcTDPCSpot.Text
    End Select
    mSetTime
    mSetCommands
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Daypart
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilMsg As Integer, ilSetBox As Integer) As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilOneYes As Integer
    Dim ilOneTime As Integer
    If smSave(1) = "" Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Daypart Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = NAMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If (smLibSave(1) = "") And (smLibSave(2) = "") And (smLibSave(3) = "") Then
        ilOneTime = False
        For ilLoop = 1 To imMaxTDRows Step 1
            If smTDSave(ilLoop, TDSTARTINDEX) <> "" Then
            'If (tmRdf.iStartTime(0, ilLoop) <> 1) Or (tmRdf.iStartTime(1, ilLoop) <> 0) Then
                If smTDSave(ilLoop, TDENDINDEX) = "" Then
                'If (tmRdf.iEndTime(0, ilLoop) = 1) And (tmRdf.iEndTime(1, ilLoop) = 0) Then
                    If ilMsg = SHOWMSG Then
                        ilRes = MsgBox("End Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    End If
                    If ilSetBox Then
                        imTDRowNo = ilLoop
                        imTDBoxNo = TDENDINDEX
                    End If
                    mTestSaveFields = NO
                    Exit Function
                End If
                '8/14/06:  Allow XMidNight
                'If (ilMsg = SHOWMSG) And ilSetBox Then
                '    If (gTimeToLong(smTDSave(ilLoop, TDENDINDEX), True) < gTimeToLong(smTDSave(ilLoop, TDSTARTINDEX), False)) Or (gTimeToLong(smTDSave(ilLoop, TDENDINDEX), True) = gTimeToLong(smTDSave(ilLoop, TDSTARTINDEX), False)) Then
                '        If ilMsg = SHOWMSG Then
                '            ilRes = MsgBox("End Time must be After Start Time", vbOkOnly + vbExclamation, "Incomplete")
                '        End If
                '        If ilSetBox Then
                '            imTDRowNo = ilLoop
                '            imTDBoxNo = TDENDINDEX
                '        End If
                '        mTestSaveFields = NO
                '        Exit Function
                '    End If
                'End If
                ilOneTime = True
                ilOneYes = False
                For ilIndex = 1 To 7 Step 1
                    If smTDSave(ilLoop, TDDAYINDEX + ilIndex - 1) = "Y" Then
                        ilOneYes = True
                        Exit For
                    End If
                Next ilIndex
                If Not ilOneYes Then
                    If ilMsg = SHOWMSG Then
                        ilRes = MsgBox("One Day must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    End If
                    If ilSetBox Then
                        imTDRowNo = ilLoop
                        imTDBoxNo = TDDAYINDEX
                    End If
                    mTestSaveFields = NO
                    Exit Function
                End If
            End If
        Next ilLoop
        If Not ilOneTime Then
            If ilMsg = SHOWMSG Then
                ilRes = MsgBox("Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
            End If
            If ilSetBox Then
                imTDRowNo = ilLoop
                imTDBoxNo = TDSTARTINDEX
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If imSave(3) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Avail selection type must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = AVAILINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If (imSave(3) = 1) Or (imSave(3) = 2) Then
        If smSave(2) = "" Then
            If ilMsg = SHOWMSG Then
                ilRes = MsgBox("Avail must be specified", vbOKOnly + vbExclamation, "Incomplete")
            End If
            If ilSetBox Then
                imBoxNo = AVAILINDEX + 1
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If imSave(1) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Base Daypart must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = BASEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSave(2) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Report Daypart must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = RPTINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSave(4) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Override Hour must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = HRSINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSave(7) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Override Print Rule must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = OVERRIDEPRTINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSave(6) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Active/Dormant must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imBoxNo = ACTIVEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
Private Sub pbcActive_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcActive_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
        imSave(6) = 0
        tmCtrls(imBoxNo).iChg = True
        pbcActive_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        imSave(6) = 1
        tmCtrls(imBoxNo).iChg = True
        pbcActive_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(6) = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(6) = 1
            pbcActive_Paint
        ElseIf imSave(6) = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(6) = 0
            pbcActive_Paint
        End If
    End If
    mSetChg imBoxNo, True
End Sub
Private Sub pbcActive_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSave(6) = 0 Then
        imSave(6) = 1
    Else
        imSave(6) = 0
    End If
    tmCtrls(imBoxNo).iChg = True
    pbcActive_Paint
    mSetChg imBoxNo, True
End Sub
Private Sub pbcActive_Paint()
    pbcActive.Cls
    pbcActive.CurrentX = fgBoxInsetX
    pbcActive.CurrentY = 0 'fgBoxInsetY
    If imSave(6) = 0 Then
        pbcActive.Print "Active"
    Else
        pbcActive.Print "Dormant"
    End If
End Sub
Private Sub pbcAvType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcAvType_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
        imSave(3) = 0
        pbcAvType_Paint
    ElseIf KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
        imSave(3) = 1
        pbcAvType_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        imSave(3) = 2
        pbcAvType_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(3) = 0 Then
            imSave(3) = 1
            pbcAvType_Paint
        ElseIf imSave(3) = 1 Then
            imSave(3) = 2
            pbcAvType_Paint
        ElseIf imSave(3) = 2 Then
            imSave(3) = 0
            pbcAvType_Paint
        End If
    End If
    mSetChg imBoxNo, True
End Sub
Private Sub pbcAvType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSave(3) = 0 Then
        imSave(3) = 1
    ElseIf imSave(3) = 1 Then
        imSave(3) = 2
    Else
        imSave(3) = 0
    End If
    pbcAvType_Paint
    mSetChg imBoxNo, True
End Sub
Private Sub pbcAvType_Paint()
    pbcAvType.Cls
    pbcAvType.CurrentX = fgBoxInsetX
    pbcAvType.CurrentY = 0 'fgBoxInsetY
    If imSave(3) = 0 Then
        pbcAvType.Print "All avails"
    ElseIf imSave(3) = 1 Then
        pbcAvType.Print "Book into"
    Else
        pbcAvType.Print "Exclude"
    End If
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    pbcArrow.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcHrs_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcHrs_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = 1
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = 2
    ElseIf imBoxNo = HRSINDEX Then
        ilIndex = 4
    Else
        Exit Sub
    End If
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        imSave(ilIndex) = 0
        pbcHrs_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        imSave(ilIndex) = 1
        pbcHrs_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(ilIndex) = 0 Then
            imSave(ilIndex) = 1
            pbcHrs_Paint
        ElseIf imSave(ilIndex) = 1 Then
            imSave(ilIndex) = 0
            pbcHrs_Paint
        End If
    End If
    mSetChg imBoxNo, True
End Sub
Private Sub pbcHrs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = 1
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = 2
    ElseIf imBoxNo = HRSINDEX Then
        ilIndex = 4
    Else
        Exit Sub
    End If
    If imSave(ilIndex) = 0 Then
        imSave(ilIndex) = 1
    Else
        imSave(ilIndex) = 0
    End If
    pbcHrs_Paint
    mSetChg imBoxNo, True
End Sub
Private Sub pbcHrs_Paint()
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = 1
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = 2
    ElseIf imBoxNo = HRSINDEX Then
        ilIndex = 4
    Else
        Exit Sub
    End If
    pbcHrs.Cls
    pbcHrs.CurrentX = fgBoxInsetX
    pbcHrs.CurrentY = 0 'fgBoxInsetY
    If imSave(ilIndex) = 0 Then
        pbcHrs.Print "Yes"
    Else
        pbcHrs.Print "No"
    End If
End Sub
Private Sub pbcLib_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = LIB1INDEX To LIB3INDEX Step 1
        If (X >= tmLibCtrls(ilBox).fBoxX) And (X <= (tmLibCtrls(ilBox).fBoxX + tmLibCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmLibCtrls(ilBox).fBoxY)) And (Y <= (tmLibCtrls(ilBox).fBoxY + tmLibCtrls(ilBox).fBoxH)) Then
                mLibSetShow imLibBoxNo
                imLibBoxNo = ilBox
                mLibEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mLibSetFocus imLibBoxNo
End Sub
Private Sub pbcLib_Paint()
    Dim ilBox As Integer

    For ilBox = imLBLibCtrls To UBound(tmLibCtrls) Step 1
        'gPaintArea pbcTD, tmTDCtrls(ilBox).fBoxX, tmTDCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15), tmTDCtrls(ilBox).fBoxW - 15, tmTDCtrls(ilBox).fBoxH - 15, WHITE
        pbcLib.CurrentX = tmLibCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcLib.CurrentY = tmLibCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcLib.Print tmLibCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcOverridePrt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcOverridePrt_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    ilIndex = 7
    If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
        imSave(ilIndex) = 0
        pbcOverridePrt_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        imSave(ilIndex) = 1
        pbcOverridePrt_Paint
    ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imSave(3) > 0 Then
            imSave(ilIndex) = 2
            pbcOverridePrt_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(ilIndex) = 0 Then
            imSave(ilIndex) = 1
            pbcOverridePrt_Paint
        ElseIf imSave(ilIndex) = 1 Then
            If imSave(3) > 0 Then
                imSave(ilIndex) = 2
            Else
                imSave(ilIndex) = 0
            End If
            pbcOverridePrt_Paint
        ElseIf imSave(ilIndex) = 2 Then
            imSave(ilIndex) = 0
            pbcOverridePrt_Paint
        End If
    End If
    mSetChg imBoxNo, True
End Sub

Private Sub pbcOverridePrt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    ilIndex = 7
    If imSave(ilIndex) = 0 Then
        imSave(ilIndex) = 1
    ElseIf imSave(ilIndex) = 1 Then
        If imSave(3) > 0 Then
            imSave(ilIndex) = 2
        Else
            imSave(ilIndex) = 0
        End If
    Else
        imSave(ilIndex) = 0
    End If
    pbcOverridePrt_Paint
    mSetChg imBoxNo, True

End Sub

Private Sub pbcOverridePrt_Paint()
    Dim ilIndex As Integer
    ilIndex = 7
    pbcOverridePrt.Cls
    pbcOverridePrt.CurrentX = fgBoxInsetX
    pbcOverridePrt.CurrentY = 0 'fgBoxInsetY
    If imSave(ilIndex) = 1 Then
        pbcOverridePrt.Print "Daypart Name"
    ElseIf imSave(ilIndex) = 2 Then
        pbcOverridePrt.Print "Avail Name"
    Else
        pbcOverridePrt.Print "None"
    End If
End Sub

Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = NAMEINDEX To ACTIVEINDEX Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmCtrls(ilBox).fBoxY)) And (Y <= (tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcSpec_Paint()
    Dim ilBox As Integer
    Dim slFont As String
    Dim slStr As String

    mPaintTitle
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        'gPaintArea pbcSpec, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
        pbcSpec.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        slStr = tmCtrls(ilBox).sShow
        If (ilBox = AVAILINDEX) And (Left$(slStr, 1) = "Ø") Then
            gPaintArea pbcSpec, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY, tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            pbcSpec.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcSpec.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
            slFont = pbcSpec.FontName
            pbcSpec.FontName = lbcAvail.FontName '"MS Sans Serif"
            pbcSpec.FontBold = False
            pbcSpec.Print Left$(slStr, 1)
            pbcSpec.FontName = slFont
            pbcSpec.FontBold = True
            pbcSpec.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX + pbcSpec.TextWidth("Ø ")
            pbcSpec.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcSpec.Print right$(slStr, Len(slStr) - 1)
        Else
            If ilBox <> AVAILINDEX + 1 Then
                pbcSpec.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
                pbcSpec.Print tmCtrls(ilBox).sShow
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.HWnd Then
        Exit Sub
    End If
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    imTDRowNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    If imBoxNo = AVAILINDEX + 1 Then
        If mAvailsBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = DPGROUPINDEX Then
        If mDPGroupBranch() Then
            Exit Sub
        End If
    End If
    'Don't allow tabbing past mandatory fields
    'If (imBoxNo >= imLBCtrls) And (imBoxNo <= tmCtrls) Then
    '    If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
    '        Beep
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
    'End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetShow 1
                ilBox = 2
            End If
        Case NAMEINDEX
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            If (cmcUpdate.Enabled) And (igDPNameCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case HRSINDEX
            If imSave(3) = 0 Then
                ilBox = AVAILINDEX
            Else
                ilBox = AVAILINDEX + 1
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcSpecTab.HWnd Then
        Exit Sub
    End If
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    imTDRowNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    If imBoxNo = AVAILINDEX + 1 Then
        If mAvailsBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = DPGROUPINDEX Then
        If mDPGroupBranch() Then
            Exit Sub
        End If
    End If
    'Don't allow tabbing past mandatory fields
    'If (imBoxNo >= imLBCtrls) And (imBoxNo <= tmCtrls) Then
    '    If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
    '        Beep
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
    'End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = ACTIVEINDEX
        Case ACTIVEINDEX 'Last control within header
            mSetShow imBoxNo
            'If mTestSaveFields() = NO Then
            '    mEnableBox imBoxNo
            '    Exit Sub
            'End If
            If rbcType(0).Value Then
                pbcTDSTab.SetFocus
            Else
                pbcTDSTab.SetFocus
            End If
            Exit Sub
        Case 0
            ilBox = NAMEINDEX
        Case AVAILINDEX
            If imSave(3) = 0 Then
                ilBox = HRSINDEX
            Else
                ilBox = AVAILINDEX + 1
            End If
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilLoop As Integer
    Dim ilRow As Integer
    Dim slTotalPct As String
    Dim ilTestPct As Integer
    Dim slStr As String
    ilMaxRow = 0
    For ilLoop = 1 To imMaxTDRows Step 1
        If smTDSave(ilLoop, 1) = "" Then
            Exit For
        End If
        ilMaxRow = ilMaxRow + 1
    Next ilLoop
    If ilMaxRow = 0 Then
        ilMaxRow = 1
    Else
        slTotalPct = "0"
        ilTestPct = True
        For ilLoop = 1 To ilMaxRow Step 1
            If smTDSave(ilLoop, TDPCSPOTINDEX) = "" Then
                ilTestPct = False
                Exit For
            End If
            slTotalPct = gAddStr(smTDSave(ilLoop, TDPCSPOTINDEX), slTotalPct)
        Next ilLoop
        If (Not ilTestPct) Or (Val(slTotalPct) <> 100) Then
            ilMaxRow = ilMaxRow + 1
        End If
        If ilMaxRow > imMaxTDRows Then
            ilMaxRow = imMaxTDRows
        End If
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = 1 To TDPCSPOTINDEX Step 1
            If (X >= tmTDCtrls(ilBox).fBoxX) And (X <= (tmTDCtrls(ilBox).fBoxX + tmTDCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmTDCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmTDCtrls(ilBox).fBoxY + tmTDCtrls(ilBox).fBoxH)) Then
                    If ((imTDBoxNo = TDSTARTINDEX) Or (imTDBoxNo = TDENDINDEX)) And (edcTDStartEnd.Text <> "") Then
                        slStr = edcTDStartEnd.Text
                        If Not gValidTime(slStr) Then
                            Beep
                            edcTDStartEnd.SetFocus
                            Exit Sub
                        End If
                    End If
                    mTDSetShow imTDBoxNo
                    imTDRowNo = ilRow
                    imTDBoxNo = ilBox
                    mTDEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mTDSetFocus imTDBoxNo
End Sub
Private Sub pbcTD_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim slFont As String
    Dim slStr As String
    For ilRow = LBONE To UBound(smTDSave, 1) Step 1
        For ilBox = imLBTDCtrls To UBound(tmTDCtrls) Step 1
            'gPaintArea pbcTD, tmTDCtrls(ilBox).fBoxX, tmTDCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15), tmTDCtrls(ilBox).fBoxW - 15, tmTDCtrls(ilBox).fBoxH - 15, WHITE
            pbcTD.CurrentX = tmTDCtrls(ilBox).fBoxX + fgBoxInsetX
            If (ilBox >= TDDAYINDEX) And (ilBox <= TDDAYINDEX + 6) Then
                slStr = smTDShow(ilRow, ilBox)
                pbcTD.CurrentY = tmTDCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                slFont = pbcTD.FontName
                pbcTD.FontName = "Monotype Sorts"
                pbcTD.FontBold = False
                'If slStr = "Y" Then
                '    slStr = "4"
                'Else
                '    slStr = "  "
                'End If
                pbcTD.Print slStr
                pbcTD.FontName = slFont
                pbcTD.FontBold = True
            Else
                pbcTD.CurrentY = tmTDCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                pbcTD.Print smTDShow(ilRow, ilBox)
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcTDDays_Click()
    mTDSetShow imTDBoxNo
End Sub
Private Sub pbcTDSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcTDSTab.HWnd Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If rbcType(0).Value Then
        mLibSetShow imLibBoxNo
        imLibBoxNo = -1
        imTabDirection = -1 'Set- Right to left
        Select Case imTDBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imTDRowNo = 1
                ilBox = 1
            Case TDSTARTINDEX 'Name (first control within header)
                mTDSetShow imTDBoxNo
                If imTDRowNo <= 1 Then
                    imTDBoxNo = -1
                    pbcSpecTab.SetFocus
                    Exit Sub
                Else
                    If edcTDStartEnd.Text <> "" Then
                        slStr = edcTDStartEnd.Text
                        If Not gValidTime(slStr) Then
                            Beep
                            edcTDStartEnd.SetFocus
                            Exit Sub
                        End If
                    End If
                    ilBox = TDPCSPOTINDEX
                    imTDRowNo = imTDRowNo - 1
                    imTDBoxNo = ilBox
                    mTDEnableBox ilBox
                    Exit Sub
                End If
            Case TDENDINDEX
                If edcTDStartEnd.Text <> "" Then
                    slStr = edcTDStartEnd.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcTDStartEnd.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = imTDBoxNo - 1
            Case TDPCSPOTINDEX
                ilBox = TDDAYINDEX + 6
            Case Else
                ilBox = imTDBoxNo - 1
        End Select
        mTDSetShow imTDBoxNo
        imTDBoxNo = ilBox
        mTDEnableBox ilBox
    Else
        mTDSetShow imTDBoxNo
        imTDBoxNo = -1
        imTabDirection = -1 'Set- Right to left
        Select Case imLibBoxNo
            Case -1 'Tab from control prior to form area
                ilBox = 1
            Case LIB1INDEX 'Name (first control within header)
                mLibSetShow imLibBoxNo
                imLibBoxNo = -1
                pbcSpecTab.SetFocus
                Exit Sub
            Case Else
                ilBox = imLibBoxNo - 1
        End Select
        mLibSetShow imLibBoxNo
        imLibBoxNo = ilBox
        mLibEnableBox ilBox
    End If
End Sub
Private Sub pbcTDTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim ilRow As Integer
    Dim slTotalPct As String
    Dim ilTestPct As Integer
    Dim slStr As String
    If GetFocus() <> pbcTDTab.HWnd Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If rbcType(0).Value Then
        mLibSetShow imLibBoxNo
        imLibBoxNo = -1
        imTabDirection = 0 'Set- Left to right
        Select Case imTDBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                For imTDRowNo = imMaxTDRows To 1 Step -1
                    If smTDSave(imTDRowNo, 1) <> "" Then
                        If imTDRowNo = imMaxTDRows Then
                            ilBox = TDPCSPOTINDEX
                            Exit For
                        Else
                            imTDRowNo = imTDRowNo + 1
                            ilBox = 1
                            Exit For
                        End If
                    End If
                Next imTDRowNo
                If imTDRowNo <= 0 Then
                    imTDRowNo = 1
                    ilBox = 1
                End If
            Case TDSTARTINDEX
                ilRow = 1
                For ilLoop = 1 To imMaxTDRows Step 1
                    If smTDSave(ilLoop, 1) = "" Then
                        Exit For
                    End If
                    ilRow = ilRow + 1
                Next ilLoop
                If (edcTDStartEnd.Text = "") And (imTDRowNo = 1) And (ilRow = 1) Then
                    Beep
                    ilBox = 1
                    imTDBoxNo = ilBox
                    mTDEnableBox ilBox
                    Exit Sub
                End If
                If (edcTDStartEnd.Text = "") And (ilRow = imTDRowNo) Then
                    mTDSetShow imTDBoxNo
                    imTDBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                slStr = edcTDStartEnd.Text
                If Not gValidTime(slStr) Then
                    Beep
                    edcTDStartEnd.SetFocus
                    Exit Sub
                End If
                ilBox = imTDBoxNo + 1
            Case TDENDINDEX
                slStr = edcTDStartEnd.Text
                If Not gValidTime(slStr) Then
                    Beep
                    edcTDStartEnd.SetFocus
                    Exit Sub
                End If
                ilBox = imTDBoxNo + 1
            Case TDPCSPOTINDEX 'Last control within header
                mTDSetShow imTDBoxNo
                slTotalPct = "0"
                ilTestPct = True
                For ilLoop = 1 To imTDRowNo Step 1
                    If smTDSave(ilLoop, TDPCSPOTINDEX) = "" Then
                        ilTestPct = False
                        Exit For
                    End If
                    slTotalPct = gAddStr(smTDSave(ilLoop, TDPCSPOTINDEX), slTotalPct)
                Next ilLoop
                For ilLoop = imTDRowNo + 1 To imMaxTDRows Step 1
                    If smTDSave(ilLoop, TDPCSPOTINDEX) <> "" Then
                        ilTestPct = False
                    End If
                Next ilLoop
                If (imTDRowNo = imMaxTDRows) Or ((ilTestPct) And (Val(slTotalPct) = 100)) Then
                    'If mTestSaveFields() = NO Then
                    '    mTDEnableBox imTDBoxNo
                    '    Exit Sub
                    'End If
                    imTDBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imTDRowNo = imTDRowNo + 1
                ilBox = 1
                imTDBoxNo = ilBox
                mTDEnableBox ilBox
                Exit Sub
            Case Else
                ilBox = imTDBoxNo + 1
        End Select
        mTDSetShow imTDBoxNo
        imTDBoxNo = ilBox
        mTDEnableBox ilBox
    Else
        mTDSetShow imTDBoxNo
        imTDBoxNo = -1
        imTDRowNo = -1
        Select Case imLibBoxNo
            Case -1 'Tab from control prior to form area
                ilBox = LIB1INDEX
            Case LIB3INDEX
                mLibSetShow imLibBoxNo
                'If mTestSaveFields() = NO Then
                '    mLibEnableBox imLibBoxNo
                '    Exit Sub
                'End If
                imLibBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imLibBoxNo + 1
        End Select
        mLibSetShow imLibBoxNo
        imLibBoxNo = ilBox
        mLibEnableBox ilBox
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
                    Select Case imTDBoxNo
                        Case TDSTARTINDEX
                            imBypassFocus = True    'Don't change select text
                            edcTDStartEnd.SetFocus
                            'SendKeys slKey
                            gSendKeys edcTDStartEnd, slKey
                        Case TDENDINDEX
                            imBypassFocus = True    'Don't change select text
                            edcTDStartEnd.SetFocus
                            'SendKeys slKey
                            gSendKeys edcTDStartEnd, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            plcLib.Visible = False
            pbcLib.Visible = False
            plcTD.Visible = True
            pbcTD.Visible = True
            'pbcTDSTab.SetFocus
        Else
            plcTD.Visible = False
            pbcTD.Visible = False
            plcLib.Visible = True
            pbcLib.Visible = True
            pbcTDSTab.SetFocus
            mLibPop
        End If
    End If
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = -1
    mTDSetShow imTDBoxNo
    imTDBoxNo = -1
    mLibSetShow imLibBoxNo
    imLibBoxNo = -1
    pbcArrow.Visible = False
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case AVAILINDEX + 1
            imLbcArrowSetting = False
            gProcessLbcClick lbcAvail, edcSpecDropDown, imChgMode, imLbcArrowSetting
        Case DPGROUPINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcDPGroup, edcSpecDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Daypart"
End Sub

Private Function mDPGroupBranch() As Integer
'
'   ilRet = mDPGroupBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcSpecDropDown, lbcDPGroup, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcSpecDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mDPGroupBranch = False
        Exit Function
    End If
    sgMnfCallType = "3"
    igMNmCallSource = CALLSOURCERATECARD
    If edcSpecDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CopyInv^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "CopyInv^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CopyInv^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "CopyInv^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'CopyInv.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'CopyInv.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mDPGroupBranch = True
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
        lbcDPGroup.Clear
        smDPGroupCodeTag = ""
        mDPGroupPop
        If imTerminate Then
            mDPGroupBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcDPGroup
        sgMNmName = ""
        If gLastFound(lbcDPGroup) > 0 Then
            imChgMode = True
            lbcDPGroup.ListIndex = gLastFound(lbcDPGroup)
            edcSpecDropDown.Text = lbcDPGroup.List(lbcDPGroup.ListIndex)
            imChgMode = False
            mDPGroupBranch = False
            mSetChg imBoxNo, False
        Else
            imChgMode = True
            lbcDPGroup.ListIndex = 1
            edcSpecDropDown.Text = lbcDPGroup.List(1)
            imChgMode = False
            mSetChg imBoxNo, False
            edcSpecDropDown.SetFocus
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
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDPGroupPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate DPGroupouncer list        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mDPGroupPop()
'
'   mDPGroupPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slDPGroup As String      'DPGroup name, saved to determine if changed
    Dim ilDPGroup As Integer      'DPGroup name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "3"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilDPGroup = lbcDPGroup.ListIndex
    If ilDPGroup > 1 Then
        slDPGroup = lbcDPGroup.List(ilDPGroup)
    End If
    ilRet = gIMoveListBox(Daypart, lbcDPGroup, tmDPGroupCode(), smDPGroupCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDPGroupPopErr
        gCPErrorMsg ilRet, "mDPGroupPop (gIMoveListBox)", Daypart
        On Error GoTo 0
        lbcDPGroup.AddItem "[None]", 0
        lbcDPGroup.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilDPGroup > 1 Then
            gFindMatch slDPGroup, 2, lbcDPGroup
            If gLastFound(lbcDPGroup) > 1 Then
                lbcDPGroup.ListIndex = gLastFound(lbcDPGroup)
            Else
                lbcDPGroup.ListIndex = -1
            End If
        Else
            lbcDPGroup.ListIndex = ilDPGroup
        End If
        imChgMode = False
    End If
    Exit Sub
mDPGroupPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
Private Sub mPaintTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcSpec.ForeColor
    slFontName = pbcSpec.FontName
    flFontSize = pbcSpec.FontSize
    pbcSpec.ForeColor = BLUE
    pbcSpec.FontBold = False
    pbcSpec.FontSize = 7
    pbcSpec.FontName = "Arial"
    pbcSpec.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    pbcSpec.CurrentX = tmCtrls(DPGROUPINDEX).fBoxX + 15
    pbcSpec.CurrentY = tmCtrls(DPGROUPINDEX).fBoxY - 15
    pbcSpec.Print "Daypart Group"
    pbcSpec.FontSize = flFontSize
    pbcSpec.FontName = slFontName
    pbcSpec.FontSize = flFontSize
    pbcSpec.ForeColor = llColor
    pbcSpec.FontBold = True
End Sub

