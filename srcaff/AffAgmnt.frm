VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAgmnt 
   Caption         =   "Affiliate Agreement Flights"
   ClientHeight    =   6435
   ClientLeft      =   480
   ClientTop       =   675
   ClientWidth     =   9780
   Icon            =   "AffAgmnt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9780
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Delivery"
      ForeColor       =   &H80000008&
      Height          =   4530
      Index           =   1
      Left            =   7155
      TabIndex        =   67
      Top             =   5940
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame frcExclude 
         Caption         =   "Excludes"
         Height          =   870
         Left            =   0
         TabIndex        =   266
         Top             =   3600
         Width           =   7185
         Begin VB.CheckBox ckcExcludeCntrTypeV 
            Caption         =   "Reservation"
            Height          =   195
            Left            =   5085
            TabIndex        =   275
            Top             =   615
            Width           =   1320
         End
         Begin VB.CheckBox ckcExcludeCntrTypeS 
            Caption         =   "PSA"
            Height          =   195
            Left            =   3465
            TabIndex        =   274
            Top             =   615
            Width           =   1230
         End
         Begin VB.CheckBox ckcExcludeCntrTypeM 
            Caption         =   "Promo"
            Height          =   225
            Left            =   2235
            TabIndex        =   273
            Top             =   615
            Width           =   960
         End
         Begin VB.CheckBox ckcExcludeCntrTypeT 
            Caption         =   "Remnant"
            Height          =   225
            Left            =   5085
            TabIndex        =   272
            Top             =   375
            Width           =   1110
         End
         Begin VB.CheckBox ckcExcludeCntrTypeR 
            Caption         =   "Direct Response"
            Height          =   210
            Left            =   3465
            TabIndex        =   271
            Top             =   375
            Width           =   1500
         End
         Begin VB.CheckBox ckcExcludeCntrTypeQ 
            Caption         =   "Per Inquiry"
            Height          =   195
            Left            =   2235
            TabIndex        =   270
            Top             =   375
            Width           =   1185
         End
         Begin VB.CheckBox ckcExcludeFillSpot 
            Caption         =   "Fill"
            Height          =   225
            Left            =   2070
            TabIndex        =   267
            Top             =   135
            Width           =   735
         End
         Begin VB.Label lacExcludeCntrType 
            Caption         =   "Contract Types to Exclude"
            Height          =   225
            Left            =   120
            TabIndex        =   269
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lacExcludeSpotType 
            Caption         =   "Spot Types to Exclude"
            Height          =   285
            Left            =   120
            TabIndex        =   268
            Top             =   90
            Width           =   1875
         End
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "Univision"
         Height          =   195
         Index           =   2
         Left            =   7560
         TabIndex        =   265
         Top             =   480
         Width           =   1755
      End
      Begin VB.Frame frcAudioDelivery 
         Caption         =   "Audio Delivery"
         Height          =   1725
         Left            =   0
         TabIndex        =   248
         Top             =   1875
         Width           =   8820
         Begin VB.CheckBox ckcSendNotCarried 
            Caption         =   "Treat Pledge Status of ""Not Carried"" same as ""Aired"""
            Height          =   195
            Left            =   3345
            TabIndex        =   162
            Top             =   495
            Width           =   4140
         End
         Begin VB.CheckBox ckcSendDelays 
            Caption         =   "Send Delays to X-Digital Auto Playback"
            Height          =   210
            Left            =   3345
            TabIndex        =   253
            TabStop         =   0   'False
            Top             =   750
            Width           =   3465
         End
         Begin VB.ListBox lbcAudioDelivery 
            Height          =   1185
            Left            =   240
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   249
            Top             =   240
            Width           =   2775
         End
         Begin VB.Frame frcIdcGroup 
            Height          =   330
            Left            =   7200
            TabIndex        =   260
            Top             =   75
            Width           =   1935
            Begin VB.OptionButton optIDCGroup 
               Caption         =   "Location"
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   263
               Top             =   120
               Width           =   915
            End
            Begin VB.OptionButton optIDCGroup 
               Caption         =   "Station"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   262
               Top             =   120
               Width           =   795
            End
            Begin VB.OptionButton optIDCGroup 
               Caption         =   "None"
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   261
               Top             =   120
               Width           =   735
            End
            Begin VB.Label lbcIDCGroup 
               Caption         =   "IDC Group By"
               Height          =   270
               Left            =   240
               TabIndex        =   264
               Top             =   120
               Width           =   1005
            End
         End
         Begin VB.TextBox txtIDCReceiverID 
            Height          =   285
            Left            =   5145
            MaxLength       =   5
            TabIndex        =   257
            Top             =   1365
            Width           =   1305
         End
         Begin VB.TextBox txtXDReceiverID 
            Height          =   285
            Left            =   5145
            MaxLength       =   9
            TabIndex        =   255
            Top             =   1020
            Width           =   1320
         End
         Begin VB.OptionButton optVoiceTracked 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   5070
            TabIndex        =   251
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optVoiceTracked 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   5850
            TabIndex        =   252
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label lacIDCReceiverID 
            Caption         =   "IDC Site ID:"
            Height          =   255
            Left            =   3345
            TabIndex        =   256
            Top             =   1365
            Width           =   1650
         End
         Begin VB.Label lacXDReceiverID 
            Caption         =   "X-Digital Station ID:"
            Height          =   255
            Left            =   3345
            TabIndex        =   254
            Top             =   1065
            Width           =   1410
         End
         Begin VB.Label lacVoiceTracked 
            Caption         =   "Voice-Tracked"
            Height          =   255
            Left            =   3345
            TabIndex        =   250
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label lacAudioTitle 
            Caption         =   "Audio Delivery Service"
            Height          =   225
            Left            =   120
            TabIndex        =   259
            Top             =   1155
            Width           =   1845
         End
      End
      Begin VB.Frame frcLogDelivery 
         Caption         =   "Log Delivery"
         Height          =   1290
         Left            =   0
         TabIndex        =   242
         Top             =   570
         Width           =   7455
         Begin VB.CheckBox ckcUnivision 
            Caption         =   "Export to Univision"
            Height          =   255
            Left            =   3240
            TabIndex        =   247
            Top             =   840
            Width           =   2055
         End
         Begin VB.ListBox lbcLogDelivery 
            Height          =   960
            Left            =   240
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   243
            Top             =   240
            Width           =   2775
         End
         Begin VB.Frame frcSendLogEMail 
            Caption         =   "Frame3"
            Height          =   285
            Left            =   3240
            TabIndex        =   244
            Top             =   240
            Width           =   4005
            Begin VB.OptionButton rbcSendLogEMail 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   145
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton rbcSendLogEMail 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   2520
               TabIndex        =   246
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.Label lacSendLogEMail 
               Caption         =   "Send E-Mail Notification for Log"
               Height          =   270
               Left            =   120
               TabIndex        =   245
               Top             =   0
               Width           =   2340
            End
         End
         Begin VB.Label lacLogTitle 
            Caption         =   "Log Delivery Service"
            Height          =   225
            Left            =   360
            TabIndex        =   258
            Top             =   960
            Width           =   1845
         End
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "Jelli"
         Height          =   195
         Index           =   6
         Left            =   1725
         TabIndex        =   79
         Top             =   825
         Width           =   930
      End
      Begin VB.Frame frcVoiceTracked 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   2145
         Width           =   2955
      End
      Begin VB.Frame frdAudio 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   585
         Left            =   120
         TabIndex        =   96
         Top             =   1455
         Width           =   8490
         Begin VB.OptionButton rbcAudio 
            Caption         =   "None"
            Height          =   195
            Index           =   5
            Left            =   4260
            TabIndex        =   103
            Top             =   315
            Width           =   975
         End
         Begin VB.OptionButton rbcAudio 
            Caption         =   "Wegener-iPump"
            Height          =   195
            Index           =   4
            Left            =   1770
            TabIndex        =   101
            Top             =   315
            Width           =   1575
         End
         Begin VB.OptionButton rbcAudio 
            Caption         =   "IDC"
            Height          =   195
            Index           =   3
            Left            =   3495
            TabIndex        =   102
            Top             =   315
            Width           =   765
         End
         Begin VB.OptionButton rbcAudio 
            Caption         =   "Wegener-Compel"
            Height          =   195
            Index           =   2
            Left            =   4725
            TabIndex        =   100
            Top             =   30
            Width           =   1710
         End
         Begin VB.OptionButton rbcAudio 
            Caption         =   "X-Digital: Break"
            Height          =   195
            Index           =   1
            Left            =   3165
            TabIndex        =   99
            Top             =   30
            Width           =   1500
         End
         Begin VB.OptionButton rbcAudio 
            Caption         =   "X-Digital: ISCI"
            Height          =   195
            Index           =   0
            Left            =   1770
            TabIndex        =   98
            Top             =   30
            Width           =   1335
         End
         Begin VB.Label lacAudio 
            Caption         =   "Audio Delivery Service"
            Height          =   195
            Left            =   0
            TabIndex        =   97
            Top             =   15
            Width           =   1785
         End
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "Clear Channel"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   4995
         TabIndex        =   78
         Top             =   540
         Width           =   1485
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "CBS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   4110
         TabIndex        =   77
         Top             =   540
         Width           =   795
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "Marketron"
         Height          =   195
         Index           =   3
         Left            =   2850
         TabIndex        =   76
         Top             =   540
         Width           =   1185
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "Cumulus"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   75
         Top             =   540
         Width           =   960
      End
      Begin VB.CheckBox ckcExportTo 
         Caption         =   "CSI Electronic Affidavit"
         Height          =   195
         Index           =   0
         Left            =   6825
         TabIndex        =   74
         Top             =   270
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame frcPostType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4485
         TabIndex        =   91
         Top             =   3975
         Width           =   7215
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Program Start Times"
            Height          =   255
            Index           =   2
            Left            =   2925
            TabIndex        =   94
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Dayparts"
            Height          =   255
            Index           =   1
            Left            =   6420
            TabIndex        =   95
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Exact Times"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   93
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblPostType 
            Caption         =   "Posting Method"
            Height          =   270
            Left            =   0
            TabIndex        =   92
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame frcLogType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4470
         TabIndex        =   80
         Top             =   4215
         Width           =   6405
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Break Numbers"
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   84
            Top             =   30
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Time Range"
            Height          =   255
            Index           =   1
            Left            =   2550
            TabIndex        =   83
            Top             =   30
            Width           =   1395
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Exact Times"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   82
            Top             =   30
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblLogType 
            Caption         =   "Log Format"
            Height          =   270
            Left            =   0
            TabIndex        =   81
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame frcExport 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   120
         TabIndex        =   69
         Top             =   180
         Width           =   3810
         Begin VB.OptionButton rbcExportType 
            Caption         =   "Export"
            Height          =   255
            Index           =   1
            Left            =   2430
            TabIndex        =   72
            Top             =   0
            Width           =   825
         End
         Begin VB.OptionButton rbcExportType 
            Caption         =   "Manual"
            Height          =   255
            Index           =   0
            Left            =   1230
            TabIndex        =   71
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblExportType 
            Caption         =   "Affidavit Control"
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.Frame frcPosting 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   285
         Left            =   120
         TabIndex        =   85
         Top             =   1110
         Width           =   7785
         Begin VB.OptionButton optPost 
            Caption         =   "Exact Times by Advertiser"
            Height          =   255
            Index           =   3
            Left            =   4410
            TabIndex        =   89
            Top             =   -15
            Width           =   2445
         End
         Begin VB.OptionButton optPost 
            Caption         =   "Exact Times by Date"
            Height          =   255
            Index           =   2
            Left            =   2460
            TabIndex        =   88
            Top             =   0
            Width           =   2040
         End
         Begin VB.OptionButton optPost 
            Caption         =   "C.P. Spot Count"
            Height          =   255
            Index           =   1
            Left            =   7245
            TabIndex        =   90
            Top             =   0
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.OptionButton optPost 
            Caption         =   "C.P. Receipt Only"
            Height          =   255
            Index           =   0
            Left            =   690
            TabIndex        =   87
            Top             =   0
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.Label Label21 
            Caption         =   "Posting"
            Height          =   255
            Left            =   0
            TabIndex        =   86
            Top             =   -15
            Width           =   690
         End
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Index           =   1
         Left            =   75
         ScaleHeight     =   90
         ScaleWidth      =   120
         TabIndex        =   105
         Top             =   4260
         Width           =   120
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   90
         ScaleHeight     =   45
         ScaleWidth      =   60
         TabIndex        =   68
         Top             =   195
         Width           =   60
      End
      Begin VB.Label lacExportTo 
         Caption         =   "Log Delivery Service"
         Height          =   225
         Left            =   120
         TabIndex        =   73
         Top             =   525
         Width           =   1845
      End
   End
   Begin VB.PictureBox pbcArial 
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
      Left            =   -120
      ScaleHeight     =   165
      ScaleWidth      =   195
      TabIndex        =   240
      Top             =   6105
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   4170
      Index           =   0
      Left            =   7530
      TabIndex        =   11
      Top             =   5400
      Width           =   9195
      Begin VB.Frame frcService 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Top             =   3780
         Width           =   3975
         Begin VB.OptionButton optService 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   59
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optService 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   60
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label lacService 
            Caption         =   "Service Agreement"
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   15
            Width           =   1605
         End
      End
      Begin VB.Frame frcCompensation 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   225
         Left            =   4575
         TabIndex        =   61
         Top             =   3780
         Width           =   4590
         Begin VB.OptionButton optComp 
            Caption         =   "Barter"
            Height          =   255
            Index           =   0
            Left            =   1395
            TabIndex        =   63
            Top             =   0
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optComp 
            Caption         =   "Affiliate"
            Height          =   255
            Index           =   1
            Left            =   2460
            TabIndex        =   64
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optComp 
            Caption         =   "Network"
            Height          =   255
            Index           =   2
            Left            =   3525
            TabIndex        =   65
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Compensation"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.CommandButton cmcBrowse 
         Caption         =   "&Browse..."
         Height          =   315
         Left            =   7410
         TabIndex        =   56
         Top             =   3315
         Width           =   1125
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         Top             =   2115
         Width           =   2655
      End
      Begin VB.Frame frcContractPrinted 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   1380
         Width           =   3975
         Begin VB.OptionButton optContractPrinted 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   29
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optContractPrinted 
            Caption         =   "No"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   30
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label lacContractPrinted 
            Caption         =   "Contract Printed"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   15
            Width           =   1890
         End
      End
      Begin VB.TextBox txtHistorialDate 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   165
         Width           =   855
      End
      Begin VB.Frame frcFormerNCR 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4575
         TabIndex        =   38
         Top             =   1755
         Width           =   4020
         Begin VB.OptionButton optFormerNCR 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   2325
            TabIndex        =   40
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optFormerNCR 
            Caption         =   "No"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   41
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Chronically Overdue"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   15
            Width           =   2475
         End
      End
      Begin VB.Frame frcNCR 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   1755
         Width           =   3990
         Begin VB.OptionButton optNCR 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   37
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optNCR 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   36
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Critically Overdue"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   15
            Width           =   1905
         End
      End
      Begin VB.Frame frcSuppressNotices 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   2115
         Width           =   3990
         Begin VB.OptionButton optSuppressNotices 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   45
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optSuppressNotices 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   44
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lacSuppressNotices 
            Caption         =   "Suppress Overdue Notices"
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   15
            Width           =   2235
         End
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   -30
         ScaleHeight     =   45
         ScaleWidth      =   15
         TabIndex        =   66
         Top             =   3900
         Width           =   15
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   -30
         ScaleHeight     =   45
         ScaleWidth      =   60
         TabIndex        =   12
         Top             =   75
         Width           =   60
      End
      Begin VB.TextBox txtDays 
         Height          =   285
         Left            =   6900
         MaxLength       =   3
         TabIndex        =   32
         Top             =   1365
         Width           =   495
      End
      Begin VB.TextBox txtComments 
         Height          =   285
         Left            =   1140
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   7395
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcMarketRep 
         Height          =   315
         Left            =   1440
         TabIndex        =   51
         Top             =   2730
         Width           =   2655
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcServiceRep 
         Height          =   315
         Left            =   5880
         TabIndex        =   53
         Top             =   2730
         Width           =   2655
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcContractPDF 
         Height          =   315
         Left            =   1440
         TabIndex        =   55
         Top             =   3330
         Width           =   5805
         _ExtentX        =   12118
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_Calendar txtStartDate 
         Height          =   285
         Left            =   1140
         TabIndex        =   14
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Text            =   "01/18/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
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
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar txtOnAirDate 
         Height          =   285
         Left            =   3735
         TabIndex        =   18
         Top             =   165
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Text            =   "01/18/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
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
         CSI_ForceMondaySelectionOnly=   -1  'True
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar txtEndDate 
         Height          =   285
         Left            =   1140
         TabIndex        =   16
         Top             =   555
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Text            =   "01/18/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
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
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar txtOffAirDate 
         Height          =   285
         Left            =   3735
         TabIndex        =   20
         Top             =   555
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Text            =   "01/18/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
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
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar txtDropDate 
         Height          =   285
         Left            =   6510
         TabIndex        =   24
         Top             =   555
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Text            =   "01/18/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
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
         CSI_DefaultDateType=   0
      End
      Begin VB.Label lacPDFPath 
         Caption         =   "PDF Path:"
         Height          =   270
         Left            =   120
         TabIndex        =   221
         Top             =   3090
         Width           =   8415
      End
      Begin VB.Label lacContractPDF 
         Caption         =   "Contract PDF:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3390
         Width           =   1425
      End
      Begin VB.Label lacStationServiceRep 
         Caption         =   "Station Service Rep:"
         Height          =   315
         Left            =   4575
         TabIndex        =   49
         Top             =   2445
         Width           =   4125
      End
      Begin VB.Label lacStationMarketRep 
         Caption         =   "Station Market Rep:"
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   2445
         Width           =   4215
      End
      Begin VB.Label Label30 
         Caption         =   "Web Password:"
         Height          =   255
         Left            =   4575
         TabIndex        =   46
         Top             =   2115
         Width           =   1740
      End
      Begin VB.Label lacMktgRep 
         Caption         =   "Market Rep:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2775
         Width           =   1425
      End
      Begin VB.Label Label32 
         Caption         =   "Service Rep:"
         Height          =   255
         Left            =   4575
         TabIndex        =   52
         Top             =   2775
         Width           =   1305
      End
      Begin VB.Label lacHistorialDate 
         Caption         =   "Historical Start Date:"
         Height          =   225
         Left            =   4920
         TabIndex        =   21
         Top             =   195
         Width           =   1785
      End
      Begin VB.Label lacOnAirDate 
         Caption         =   "On Air Date:"
         Height          =   225
         Left            =   2355
         TabIndex        =   17
         Top             =   195
         Width           =   1485
      End
      Begin VB.Label lacOffAirDate 
         Caption         =   "Last Date On Air:"
         Height          =   255
         Left            =   2355
         TabIndex        =   19
         Top             =   585
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "Cancellation Notice Required:"
         Height          =   255
         Left            =   4575
         TabIndex        =   31
         Top             =   1410
         Width           =   2565
      End
      Begin VB.Label Label11 
         Caption         =   "Days"
         Height          =   255
         Left            =   7635
         TabIndex        =   33
         Top             =   1410
         Width           =   480
      End
      Begin VB.Label lacEndDate 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label lacStartDate 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   1125
      End
      Begin VB.Label lacDropdate 
         Caption         =   "Drop Date:"
         Height          =   255
         Left            =   4920
         TabIndex        =   23
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMulticast 
      Height          =   525
      Left            =   3240
      TabIndex        =   232
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   926
      _Version        =   393216
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frcEvent 
      Caption         =   "Event"
      Height          =   4380
      Left            =   7710
      TabIndex        =   230
      Top             =   4425
      Visible         =   0   'False
      Width           =   8670
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEvent 
         Height          =   3435
         Left            =   75
         TabIndex        =   231
         Top             =   675
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6059
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcSeason 
         Height          =   315
         Left            =   75
         TabIndex        =   239
         Top             =   225
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VB.Label lacGreenkey 
         Caption         =   "Green=Firm"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   1470
         TabIndex        =   238
         Top             =   4140
         Width           =   810
      End
      Begin VB.Label lacOrangeKey 
         Caption         =   "Orange=Tentative"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   2310
         TabIndex        =   237
         Top             =   4140
         Width           =   1215
      End
      Begin VB.Label lacBlueKey 
         Caption         =   "Blue=Postponed"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3555
         TabIndex        =   236
         Top             =   4140
         Width           =   1230
      End
      Begin VB.Label lacRedKey 
         Caption         =   "Red=Canceled"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4770
         TabIndex        =   235
         Top             =   4140
         Width           =   1050
      End
      Begin VB.Label lacEventKey 
         Caption         =   "Event number color:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   195
         Left            =   135
         TabIndex        =   234
         Top             =   4140
         Width           =   1320
      End
      Begin VB.Label lacMulticast 
         Caption         =   "Multicast:"
         Height          =   660
         Index           =   1
         Left            =   1830
         TabIndex        =   233
         Top             =   -15
         Width           =   795
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   825
      Top             =   5565
   End
   Begin V81Affiliate.CSI_ComboBoxList cbcETDay 
      Height          =   165
      Left            =   7800
      TabIndex        =   218
      Top             =   1125
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   291
      BackColor       =   -2147483644
      ForeColor       =   -2147483644
      BorderStyle     =   0
   End
   Begin VB.Frame frcET 
      Appearance      =   0  'Flat
      Caption         =   "EstimatedTime"
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   8055
      TabIndex        =   215
      Top             =   4185
      Visible         =   0   'False
      Width           =   3750
      Begin VB.TextBox txtET 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   30
         TabIndex        =   217
         Top             =   795
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcETTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   219
         Top             =   3465
         Width           =   60
      End
      Begin VB.PictureBox pbcETSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   216
         Top             =   165
         Width           =   60
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdET 
         Height          =   2595
         Left            =   60
         TabIndex        =   220
         Top             =   285
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   4577
         _Version        =   393216
         Rows            =   10
         Cols            =   6
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.FileListBox lbcContractPDFList 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   9270
      Pattern         =   "*.PDF"
      TabIndex        =   208
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame frcTab 
      Caption         =   "Personnel"
      Height          =   3795
      Index           =   4
      Left            =   8340
      TabIndex        =   199
      Top             =   3945
      Visible         =   0   'False
      Width           =   9150
      Begin V81Affiliate.AffContactGrid udcContactGrid 
         Height          =   390
         Left            =   -45
         TabIndex        =   200
         Top             =   180
         Width           =   9090
         _ExtentX        =   15266
         _ExtentY        =   767
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Interface"
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   3
      Left            =   8625
      TabIndex        =   156
      Top             =   3750
      Width           =   7995
      Begin VB.Frame fraShow 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   285
         Left            =   120
         TabIndex        =   157
         Top             =   165
         Width           =   3345
         Begin VB.OptionButton optCarryCmml 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   160
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optCarryCmml 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   1725
            TabIndex        =   159
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Carry Commercials"
            Height          =   270
            Left            =   0
            TabIndex        =   158
            Top             =   15
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   285
         Left            =   120
         TabIndex        =   161
         Top             =   510
         Width           =   3255
         Begin VB.OptionButton optSendTape 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   165
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSendTape 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   1725
            TabIndex        =   164
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Send Tape to Station"
            Height          =   270
            Left            =   0
            TabIndex        =   163
            Top             =   15
            Width           =   1710
         End
      End
      Begin VB.TextBox txtShipInfo 
         Height          =   285
         Left            =   1815
         MaxLength       =   40
         TabIndex        =   171
         Top             =   1635
         Width           =   3975
      End
      Begin VB.TextBox txtLabelID 
         Height          =   285
         Left            =   975
         MaxLength       =   10
         TabIndex        =   169
         Top             =   1215
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdjustDates 
         Caption         =   "Adjust Dates?"
         Height          =   375
         Left            =   6225
         TabIndex        =   182
         Top             =   3405
         Width           =   1665
      End
      Begin VB.Frame frcPrintCP 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5340
         TabIndex        =   178
         Top             =   3045
         Visible         =   0   'False
         Width           =   2580
         Begin VB.OptionButton optPrintCP 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   1410
            TabIndex        =   180
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.OptionButton optPrintCP 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   810
            TabIndex        =   179
            Top             =   0
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Print C.P."
            Height          =   255
            Left            =   0
            TabIndex        =   181
            Top             =   0
            Visible         =   0   'False
            Width           =   930
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   172
         Top             =   2100
         Width           =   6420
         Begin VB.OptionButton optRadarClearType 
            Caption         =   "Prog + Cmml"
            Height          =   255
            Index           =   1
            Left            =   3060
            TabIndex        =   175
            Top             =   0
            Width           =   1245
         End
         Begin VB.OptionButton optRadarClearType 
            Caption         =   "Cmml Only"
            Height          =   255
            Index           =   0
            Left            =   1890
            TabIndex        =   174
            Top             =   0
            Width           =   1155
         End
         Begin VB.OptionButton optRadarClearType 
            Caption         =   "Schd"
            Height          =   255
            Index           =   2
            Left            =   4395
            TabIndex        =   176
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton optRadarClearType 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   3
            Left            =   5205
            TabIndex        =   177
            Top             =   0
            Width           =   930
         End
         Begin VB.Label Label22 
            Caption         =   "RADAR Clearance Type"
            Height          =   255
            Left            =   0
            TabIndex        =   173
            Top             =   15
            Width           =   1875
         End
      End
      Begin VB.TextBox txtNoCDs 
         Height          =   285
         Left            =   1485
         MaxLength       =   2
         TabIndex        =   167
         Top             =   825
         Width           =   375
      End
      Begin VB.Label lblShipInfo 
         Caption         =   "Shipping Instructions:"
         Height          =   285
         Left            =   120
         TabIndex        =   170
         Top             =   1635
         Width           =   2415
      End
      Begin VB.Label lblLabelID 
         Caption         =   "Label ID:"
         Height          =   285
         Left            =   120
         TabIndex        =   168
         Top             =   1215
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Number of CD's:"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   855
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdFastEnd 
      Caption         =   "Fast End"
      Height          =   375
      Left            =   4420
      TabIndex        =   153
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdFastAdd 
      Caption         =   "Fast Add"
      Height          =   375
      Left            =   3200
      TabIndex        =   152
      Top             =   6045
      Width           =   1100
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   0
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6135
      Width           =   0
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   60
      Top             =   5490
   End
   Begin VB.CommandButton cmdRemap 
      Caption         =   "Time Remap"
      Height          =   375
      Left            =   5640
      TabIndex        =   135
      Top             =   6045
      Width           =   1200
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Pledge Information"
      ForeColor       =   &H80000008&
      Height          =   4380
      Index           =   2
      Left            =   0
      TabIndex        =   115
      Top             =   2070
      Visible         =   0   'False
      Width           =   8835
      Begin VB.PictureBox pbcEmbeddedOrROS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   3720
         ScaleHeight     =   150
         ScaleWidth      =   765
         TabIndex        =   127
         Top             =   2625
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtLdMult 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   228
         Top             =   330
         Width           =   315
      End
      Begin VB.ListBox lbcStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Height          =   810
         ItemData        =   "AffAgmnt.frx":08CA
         Left            =   3975
         List            =   "AffAgmnt.frx":08CC
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1410
      End
      Begin VB.ListBox lbcMulticast 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         ItemData        =   "AffAgmnt.frx":08CE
         Left            =   3105
         List            =   "AffAgmnt.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   222
         Top             =   570
         Visible         =   0   'False
         Width           =   3630
      End
      Begin VB.TextBox txtAirPlay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1995
         TabIndex        =   122
         Top             =   2490
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton cmcPledgeBy 
         Caption         =   "Auto-Fill"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7200
         TabIndex        =   213
         Top             =   3720
         Width           =   1335
      End
      Begin VB.ComboBox cbcAirPlayNo 
         Height          =   315
         ItemData        =   "AffAgmnt.frx":08D2
         Left            =   6795
         List            =   "AffAgmnt.frx":08D4
         TabIndex        =   211
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkMonthlyWebPost 
         Caption         =   "Allow Monthly Posting"
         Height          =   210
         Left            =   2850
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   3780
         Width           =   2400
      End
      Begin VB.CheckBox ckcProhibitSplitCopy 
         Caption         =   "Prohibit Split Copy with 'Air Live'"
         Height          =   210
         Left            =   120
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   3780
         Width           =   2670
      End
      Begin VB.PictureBox pbcDayFed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   4635
         ScaleHeight     =   150
         ScaleWidth      =   765
         TabIndex        =   126
         Top             =   2145
         Visible         =   0   'False
         Width           =   765
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
         Left            =   4935
         Picture         =   "AffAgmnt.frx":08D6
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtDropdown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   4020
         TabIndex        =   123
         Top             =   1830
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Left            =   3675
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   210
         Begin VB.CheckBox ckcDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Small Fonts"
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
            TabIndex        =   128
            Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
            Top             =   15
            Width           =   180
         End
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1215
         TabIndex        =   121
         Top             =   2760
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcPledgeSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   15
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   120
         Top             =   585
         Width           =   60
      End
      Begin VB.PictureBox pbcPledgeTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   -30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   130
         Top             =   3900
         Width           =   60
      End
      Begin VB.PictureBox pbcPledgeFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   -30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   118
         Top             =   -60
         Width           =   60
      End
      Begin VB.PictureBox pbcArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   15
         Picture         =   "AffAgmnt.frx":09D0
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   825
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox edcNoAirPlays 
         Height          =   255
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   151
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdClearAvails 
         Caption         =   "Clear Pledges"
         Height          =   300
         Left            =   7200
         TabIndex        =   149
         Top             =   4035
         Width           =   1335
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   0
         Index           =   2
         Left            =   1680
         ScaleHeight     =   0
         ScaleWidth      =   90
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   4410
         Width           =   90
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   75
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   180
         Width           =   15
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPledge 
         Height          =   3030
         Left            =   135
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   675
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   3
         Cols            =   26
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   26
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton cmcPledgeBy 
         Caption         =   "Dayparts"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7365
         TabIndex        =   214
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lacLdMult 
         Caption         =   "Load Multiplier:"
         Height          =   255
         Left            =   135
         TabIndex        =   229
         Top             =   375
         Width           =   1140
      End
      Begin VB.Label lacMulticast 
         Caption         =   "Multicast:"
         Height          =   660
         Index           =   0
         Left            =   1830
         TabIndex        =   223
         Top             =   -15
         Width           =   795
      End
      Begin VB.Label Label24 
         Caption         =   "Display Air Play:"
         Height          =   270
         Left            =   5655
         TabIndex        =   210
         Top             =   15
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lacNoAirPlays 
         Caption         =   "# Air Plays:"
         Height          =   270
         Left            =   135
         TabIndex        =   209
         Top             =   15
         Width           =   870
      End
      Begin VB.Label lacPrgTimes 
         Alignment       =   1  'Right Justify
         Caption         =   "Program Times:"
         Height          =   210
         Left            =   6510
         TabIndex        =   201
         Top             =   465
         Width           =   2235
      End
      Begin VB.Image imcTrash 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6615
         Picture         =   "AffAgmnt.frx":0CDA
         Top             =   3795
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lacPledgeType 
         Caption         =   "Pledge by:"
         Height          =   270
         Left            =   5205
         TabIndex        =   148
         Top             =   0
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label26 
         Caption         =   "ALL Times Shown here are Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3690
         TabIndex        =   147
         Top             =   4020
         Width           =   2580
      End
      Begin VB.Image imcPrt 
         Height          =   480
         Left            =   8310
         Picture         =   "AffAgmnt.frx":0FE4
         Top             =   45
         Width           =   480
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Not Used-Controls"
      ForeColor       =   &H80000008&
      Height          =   4170
      Index           =   5
      Left            =   9090
      TabIndex        =   106
      Top             =   3270
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Frame frcMulticast 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   4770
         TabIndex        =   224
         Top             =   3495
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   226
            Top             =   -30
            Width           =   855
         End
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   1185
            TabIndex        =   225
            Top             =   -30
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label lblMulti 
            Caption         =   "* Indicates Multicast"
            Height          =   225
            Left            =   -15
            TabIndex        =   227
            Top             =   285
            Width           =   1725
         End
      End
      Begin VB.TextBox txtACName 
         Height          =   285
         Left            =   4395
         MaxLength       =   40
         TabIndex        =   204
         Top             =   3120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtACPhone 
         Height          =   285
         Left            =   2730
         MaxLength       =   20
         TabIndex        =   203
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboAffAE 
         Height          =   315
         ItemData        =   "AffAgmnt.frx":12EE
         Left            =   1470
         List            =   "AffAgmnt.frx":12F0
         Sorted          =   -1  'True
         TabIndex        =   202
         Top             =   3630
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.TextBox txtLogPassword 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   196
         Top             =   2430
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtEmailAddr 
         Height          =   285
         Left            =   4965
         MaxLength       =   240
         TabIndex        =   195
         Top             =   2430
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdGenPassword 
         Caption         =   "Generate Password"
         Height          =   300
         Left            =   1575
         TabIndex        =   194
         Top             =   2835
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Frame frcBarCode 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   225
         Left            =   120
         TabIndex        =   190
         Top             =   1980
         Visible         =   0   'False
         Width           =   3225
         Begin VB.OptionButton optBarCode 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1725
            TabIndex        =   192
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optBarCode 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   191
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Bar Codes on C.P. "
            Height          =   255
            Left            =   0
            TabIndex        =   193
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.Frame frcContractSigned 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   285
         Left            =   120
         TabIndex        =   184
         Top             =   1080
         Visible         =   0   'False
         Width           =   7620
         Begin VB.OptionButton optSigned 
            Caption         =   "Not Returned"
            Height          =   255
            Index           =   0
            Left            =   1380
            TabIndex        =   188
            Top             =   0
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton optSigned 
            Caption         =   "Returned"
            Height          =   255
            Index           =   1
            Left            =   2790
            TabIndex        =   187
            Top             =   0
            Width           =   1305
         End
         Begin VB.OptionButton optSigned 
            Caption         =   "Rejected"
            Height          =   255
            Index           =   2
            Left            =   5100
            TabIndex        =   186
            Top             =   -15
            Width           =   1080
         End
         Begin VB.TextBox txtRetDate 
            Height          =   285
            Left            =   3975
            TabIndex        =   185
            Top             =   0
            Width           =   885
         End
         Begin VB.Label labSigned 
            Caption         =   "Contract Signed"
            Height          =   255
            Left            =   0
            TabIndex        =   189
            Top             =   0
            Visible         =   0   'False
            Width           =   1365
         End
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   3
         Left            =   45
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   114
         Top             =   3075
         Width           =   60
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   3
         Left            =   120
         ScaleHeight     =   45
         ScaleWidth      =   60
         TabIndex        =   107
         Top             =   285
         Width           =   60
      End
      Begin VB.TextBox txtCP 
         Height          =   285
         Left            =   3030
         MaxLength       =   3
         TabIndex        =   111
         Top             =   555
         Width           =   630
      End
      Begin VB.TextBox txtLog 
         Height          =   285
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   109
         Top             =   540
         Width           =   630
      End
      Begin VB.TextBox txtOther 
         Height          =   285
         Left            =   5055
         MaxLength       =   3
         TabIndex        =   113
         Top             =   555
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "Aff/E-Mail Contact:"
         Height          =   255
         Left            =   3000
         TabIndex        =   207
         Top             =   3135
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label20 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   2070
         TabIndex        =   206
         Top             =   3375
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Affiliate A/E:"
         Height          =   255
         Left            =   435
         TabIndex        =   205
         Top             =   3705
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblWebPW 
         Caption         =   "Web Password"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   2430
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblWebEmail 
         Caption         =   "Web Email"
         Height          =   255
         Left            =   4005
         TabIndex        =   197
         Top             =   2430
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "CP Format:"
         Height          =   255
         Left            =   1980
         TabIndex        =   183
         Top             =   585
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Certificate of Performance Format:"
         Height          =   255
         Left            =   -3240
         TabIndex        =   110
         Top             =   1020
         Width           =   2760
      End
      Begin VB.Label Label8 
         Caption         =   "Log Format:"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   555
         Width           =   1155
      End
      Begin VB.Label Label14 
         Caption         =   "Other Format:"
         Height          =   255
         Left            =   3945
         TabIndex        =   112
         Top             =   585
         Width           =   1200
      End
   End
   Begin VB.ListBox lbcLookup1 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffAgmnt.frx":12F2
      Left            =   0
      List            =   "AffAgmnt.frx":12F4
      TabIndex        =   139
      Top             =   5145
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lbcLookup2 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffAgmnt.frx":12F6
      Left            =   -15
      List            =   "AffAgmnt.frx":12F8
      TabIndex        =   138
      Top             =   5730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "Erase"
      Height          =   375
      Left            =   6960
      TabIndex        =   136
      Top             =   6045
      Width           =   1000
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   9270
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   255
         Left            =   4275
         TabIndex        =   241
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   120
         TabIndex        =   141
         Top             =   210
         Width           =   4080
         Begin VB.OptionButton optPSSort 
            Caption         =   "All Veh"
            Height          =   255
            Index           =   3
            Left            =   3150
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   45
            Width           =   990
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "Stations"
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   45
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "DMA"
            Height          =   255
            Index           =   1
            Left            =   1230
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   45
            Width           =   660
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "Active Veh"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   45
            Width           =   1125
         End
         Begin VB.Label Label16 
            Caption         =   "By:"
            Height          =   255
            Left            =   0
            TabIndex        =   146
            Top             =   45
            Width           =   360
         End
      End
      Begin VB.ComboBox cboSSSort 
         Height          =   315
         ItemData        =   "AffAgmnt.frx":12FA
         Left            =   4275
         List            =   "AffAgmnt.frx":12FC
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   555
         Width           =   4740
      End
      Begin VB.ComboBox cboPSSort 
         Height          =   315
         ItemData        =   "AffAgmnt.frx":12FE
         Left            =   90
         List            =   "AffAgmnt.frx":1300
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   555
         Width           =   4035
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7020
         TabIndex        =   5
         Top             =   210
         Width           =   2070
         Begin VB.OptionButton optSSSort 
            Caption         =   "DMA"
            Height          =   255
            Index           =   1
            Left            =   1350
            TabIndex        =   8
            Top             =   45
            Width           =   675
         End
         Begin VB.OptionButton optSSSort 
            Caption         =   "Stations"
            Height          =   255
            Index           =   0
            Left            =   375
            TabIndex        =   7
            Top             =   45
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.Label lblSort 
            Caption         =   "By:"
            Height          =   255
            Left            =   105
            TabIndex        =   6
            Top             =   70
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5100
         TabIndex        =   2
         Top             =   210
         Width           =   1860
         Begin VB.OptionButton optExAll 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   1335
            TabIndex        =   4
            Top             =   45
            Width           =   780
         End
         Begin VB.OptionButton optExAll 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   735
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   45
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.Label Label10 
            Caption         =   "Affiliated"
            Height          =   225
            Left            =   120
            TabIndex        =   140
            Top             =   70
            Width           =   750
         End
      End
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save as New"
      Height          =   375
      Left            =   1800
      TabIndex        =   134
      Top             =   6045
      Width           =   1300
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   6645
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6435
      FormDesignWidth =   9780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8055
      TabIndex        =   137
      Top             =   6045
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as Changed"
      Height          =   375
      Left            =   240
      TabIndex        =   133
      Top             =   6045
      Width           =   1475
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   75
      TabIndex        =   10
      Top             =   1005
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   8281
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Main"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Delivery"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Pledge"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Interface"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "P&ersonnel"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacAttCode 
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   8880
      TabIndex        =   212
      Top             =   1065
      Width           =   780
   End
End
Attribute VB_Name = "frmAgmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************
'*  frmAgrmnt - enters station/vehicle agreement information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private lmAttCode As Long
Private imShttCode As Integer
Private imVefCode As Integer
Private imDatLoaded As Integer
Private imFirstTime As Integer
Private imIntegralSet As Integer
Private imFieldChgd As Integer
Private smSvOnAirDate As String
Private smSvOffAirDate As String
Private smSvDropDate As String
Private smSvMulticast As String
Private imRepopAgnmt As String
Private smLastPostedDate As String
Private imOkToChange As Integer
Private smTrueOffAirDate As String
Private smCurDate As String
Private imAffShttCode() As Integer
Private imAffVefCode() As Integer
Private imInChg As Integer
Private smShttACName As String
Private smShttACPhone As String
Private imBSMode As Integer
Private imIgnoreTimeTypeChg As Integer
Private imDateChgd As Integer
Private imVefCombo As Integer
Private smPassword As String
Private bFormWasAlreadyResized As Boolean
Private bmShowDates As Boolean
'3/23/15: Add Send Delays to XDS
Private bmSupportXDSDelay As Boolean
Private smContractPDFSubFolder As String
Private bmBypassTestDP As Boolean
Private lmPdSTime As Long
Private lmPdETime As Long
'8/12/16: Separtated Pledge and Estimate so that est can be saved in the past
Private bmPledgeDataChgd As Boolean
Private bmETDataChgd As Boolean
Private smIDCReceiverID As String
Private smMktronActiveDate As String
Private smCompensation As String

Private bmDefaultEstDay As Boolean

Private imLastPledgeColSorted As Integer
Private imLastPledgeSort As Integer

Private tmStatusTypes(0 To 14) As STATUSTYPES

Private tmOverlapInfo() As AGMNTOVERLAPINFO
Private imSource As Integer '0=New; 1=Change

Private imTabIndex As Integer
Private imIgnoreTabs As Integer
Private imIgnoreTabClick As Integer
Private imAllowInsert As Integer     'Allow insert key to be pressed
Private imRowSelNo As Integer
Private tmAvail As SSFAVAIL
Private smStationZone As String
Private tmAdjustDates() As ADJUSTDATES
Private IsAgmntDirty As Boolean
Private smWebExports As String
Private sToFileHeader As String
Private smInitialattWebEmail As String
Private smInitialattWebPW As String
Private smInitialattLogType As Integer
Private smInitialattPostType As Integer
Private smMonthlyWebPost As String
Private imColPos(0 To 25) As Integer 'Save column position because of merge

Private smPreviousNCR As String   '7-6-09  attncr
'Dan 5457
Private smPreviousXDReceiver As String
Private smPreviousXDContact As XDIGITALSTATIONINFO
Private bmIsXDSiteStation As Boolean
'Dan 5589
Private bmXDIdChanged As Boolean
'Dan 7375
Private smPreviousXDAudioDelivery As String
'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imCloseListBox As Integer
Private lmETEnableRow As Long         'Current or last row focus was on
Private lmETEnableCol As Long         'Current or last column focus was on
Private tmETAvailInfo() As ETAVAILINFO
Private tmAvailDat() As DAT
Private imETColPos(0 To 5) As Integer 'Save column position because of merge
Private tmMCDat() As DAT

Private tmPetInfo() As PETINFO

Private rst_Webl As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_Aet As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset
Private rst_Pft As ADODB.Recordset
Private rst_ept As ADODB.Recordset
Private rst_crf As ADODB.Recordset
'5882 no longer needed
'Private rst_ief As ADODB.Recordset
Private rst_Pet As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset

Private imChgDropDate As Integer
Private imChgOnAirDate As Integer
Private smAttWebInterface As String

Private smPledgeByEvent As String
Private imLastEventColSorted As Integer
Private imLastEventSort As Integer

Private bmSaving As Boolean

Private smDefaultEmbeddedOrROS As String
'7902
Private bmInMultiListChange As Boolean
'9452
Private bmSendNotCarriedChange As Boolean
'8418
Private tmVendorList() As VendorByVersion
Private Type VendorByVersion
    sName As String
    iVersion As Integer
    iCode As Integer
    sType As String
End Type

Const MONFDINDEX = 0
Const TUEFDINDEX = 1
Const WEDFDINDEX = 2
Const THUFDINDEX = 3
Const FRIFDINDEX = 4
Const SATFDINDEX = 5
Const SUNFDINDEX = 6
Const STARTTIMEFDINDEX = 7
Const ENDTIMEFDINDEX = 8
Const STATUSINDEX = 9
Const AIRPLAYINDEX = 10
Const MONPDINDEX = 11
Const TUEPDINDEX = 12
Const WEDPDINDEX = 13
Const THUPDINDEX = 14
Const FRIPDINDEX = 15
Const SATPDINDEX = 16
Const SUNPDINDEX = 17
Const DAYFEDINDEX = 18
Const STARTTIMEPDINDEX = 19
Const ENDTIMEPDINDEX = 20
Const ESTIMATETIMEINDEX = 21
Const EMBEDDEDORROSINDEX = 22
Const ESTIMATEDFIRSTINDEX = 23
Const CODEINDEX = 24
Const SORTINDEX = 25

Const ETFDDAYINDEX = 0
Const ETFDTIMEINDEX = 1
Const ETDAYINDEX = 2
Const ETTIMEINDEX = 3
Const ETAVAILINFOINDEX = 4
Const ETEPTCODEINDEX = 5

Const TABMAIN = 1
Const TABDELIVERY = 2
Const TABPLEDGE = 3
Const TABINTERFACE = 4
Const TABPERSONNEL = 5

Const MCCALLLETTERSINDEX = 0
Const MCMARKETINDEX = 1
'8/6/19: adding owner
Const MCOWNERINDEX = 2
Const MCWITHINDEX = 3
Const MCDATERANGEINDEX = 4
Const MCSHTTCODEINDEX = 5
Const MCATTCODEINDEX = 6
Const MCSELECTEDINDEX = 7

Const EVTEVENTNOINDEX = 0
Const EVTFEEDSOURCEINDEX = 1
Const EVTLANGUAGEINDEX = 2
Const EVTVISITTEAMINDEX = 3
Const EVTHOMETEAMINDEX = 4
Const EVTAIRDATEINDEX = 5
Const EVTAIRTIMEINDEX = 6
Const EVTCARRYINDEX = 7
Const EVTNOTCARRIEDINDEX = 8
Const EVTUNDECIDEDINDEX = 9
Const EVTPETINFOINDEX = 10
Const EVTSORTINDEX = 11




Private Function mAdjOverlapAgmnts(llOnAir As Long, llOffAir As Long, llDropDate As Long) As Integer
    
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim llEndDate As Long
    Dim llTestEndDate As Long
    Dim slDrop As String
    Dim ilAgreeType As Integer
    Dim ilExported As Integer
    ReDim tlCpttArray(0 To 0) As CPTTARRAY
    Dim rst_Cptt As ADODB.Recordset
    Dim ilUpper As Integer
    Dim slEndDate As String
    Dim slStartDate As String
    Dim ilRet As Integer
    Dim llAttCode As Long
    Dim temp_rst As ADODB.Recordset
    Dim slStr As String

    
    'manual = 0, web site = 1, Univision = 2
    'If rbcExportType(0).Value = True Then
    '    ilAgreeType = 0 'Manual
    'ElseIf rbcExportType(1).Value = True Then
    '    ilAgreeType = 1
    'ElseIf rbcExportType(3).Value = True Then
    '    ilAgreeType = 1 'Cumulus
    'Else
    '    ilAgreeType = 2 'Univision
    'End If
    '7701
    If rbcExportType(0).Value = True Then
        ilAgreeType = 0 'Manual
    ElseIf rbcExportType(1).Value = True Then
        'Dan 7701 exportTo is always checked when not manual, so I removed this test.
'        If ckcExportTo(0).Value = vbChecked Then
'            ilAgreeType = 1 'Network Web
'        Else
'            If mRetrieveMultiListString(lbcLogDelivery) <> "" Then
'                ilAgreeType = 1
'            End If
'
''            With lbcLogDelivery
''                If .Text <> NONESELECTEDLOG Then
''                    ilAgreeType = 1
''                End If
''            End With
'        End If
    Else
        ilAgreeType = -1
    End If
'    If rbcExportType(0).Value = True Then
'        ilAgreeType = 0 'Manual
'    ElseIf rbcExportType(1).Value = True Then
'        If ckcExportTo(0).Value = vbChecked Then
'            ilAgreeType = 1 'Network Web
'        ElseIf ckcExportTo(1).Value = vbChecked Then
'            ilAgreeType = 1 'Cumulus
'        ElseIf ckcExportTo(2).Value = vbChecked Then
'            ilAgreeType = 2 'Univision
'        ElseIf ckcExportTo(3).Value = vbChecked Then
'            ilAgreeType = 3 'Marketron
'        '6592
'        ElseIf ckcExportTo(4).Value = vbChecked Then
'            ilAgreeType = 4 'CBS
'        ElseIf ckcExportTo(6).Value = vbChecked Then
'            ilAgreeType = 6 'Jelli
'        End If
'    Else
'        ilAgreeType = -1
'    End If
    
    If UBound(tmOverlapInfo) <= LBound(tmOverlapInfo) Then
        mAdjOverlapAgmnts = True
        Exit Function
    End If
    
    On Error GoTo ErrHand
    
    If llDropDate < llOffAir Then
        llEndDate = llDropDate
    Else
        llEndDate = llOffAir
    End If
    For ilLoop = LBound(tmOverlapInfo) To UBound(tmOverlapInfo) - 1 Step 1
        'Determine if Agreement should be Deleted or Terminated
        If tmOverlapInfo(ilLoop).lDropDate < tmOverlapInfo(ilLoop).lOffAirDate Then
            llTestEndDate = tmOverlapInfo(ilLoop).lDropDate
        Else
            llTestEndDate = tmOverlapInfo(ilLoop).lOffAirDate
        End If
        
        If (llOnAir <= tmOverlapInfo(ilLoop).lOnAirDate) And (tmOverlapInfo(ilLoop).lOnAirDate <= llEndDate) Then
            'Delete old agreement as it is starts after new agreement
            SQLQuery = "SELECT cpttStartDate, cpttCode FROM Cptt WHERE (cpttAtfCode = " & tmOverlapInfo(ilLoop).lAttCode & ")"
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            ilUpper = 0
            While Not rst_Cptt.EOF
                tlCpttArray(ilUpper).lCpttCode = rst_Cptt!cpttCode
                tlCpttArray(ilUpper).sCpttStartDate = rst_Cptt!CpttStartDate
                ilUpper = ilUpper + 1
                ReDim Preserve tlCpttArray(0 To ilUpper)
                rst_Cptt.MoveNext
            Wend
                
            For ilIdx = 0 To ilUpper - 1 Step 1
                '9/25/06: Removed retaining of CPTT as user no longer allowed to set date prior to last posted date
                'ilExported = gCheckIfSpotsHaveBeenExported(imVefCode, tlCpttArray(ilIdx).sCpttStartDate, ilAgreeType)
                'If (igChangedNewErased = 1 Or igChangedNewErased = 2) And (ilExported = False) Then
                    SQLQuery = "DELETE FROM Cptt WHERE (cpttCode = " & tlCpttArray(ilIdx).lCpttCode & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    llAttCode = tmOverlapInfo(ilLoop).lAttCode
                    slStartDate = tlCpttArray(ilIdx).sCpttStartDate
                    slEndDate = DateAdd("d", 6, tlCpttArray(ilIdx).sCpttStartDate)
                    
                    'Doug (9/25/06): Remove Web Spots if using web.  For Martketron, i'm disallowing overlaps if aet exist within date change
                    'Delete the spots from the web first
                    slStr = "Select attExportType, attExportToWeb, attWebInterface from att where attCode = " & llAttCode
                    Set temp_rst = gSQLSelectCall(slStr)
                    If temp_rst.EOF = False Then
                        '7701 changed test
'                        If (temp_rst!attExportType = 1) And ((temp_rst!attExportToWeb = "Y") Or (gIfNullInteger(temp_rst!vatWvtIdCodeLog) = Vendors.Cumulus)) And gHasWebAccess Then
                        If temp_rst!attExportType = 1 And gHasWebAccess Then
                            ilRet = gWebDeleteSpots(llAttCode, Format$(slStartDate, sgSQLDateForm), Format$(slEndDate, sgSQLDateForm))
                        End If
                    End If
                    
                    SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(tlCpttArray(ilIdx).sCpttStartDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    'If (ilAgreeType = 1) Or (ilAgreeType = 2) Then
                    If (ilAgreeType > 0) Then
                        ilRet = gAlertAdd("R", "S", imVefCode, tlCpttArray(ilIdx).sCpttStartDate)
                    End If
                'End If
            Next ilIdx
            
            'Delete old agreement as it is starts after new agreement
            ' JD 12-18-2006 Added new function to properly remove an agreement.
            If Not gDeleteAgreement(tmOverlapInfo(ilLoop).lAttCode, "AffAgreementLog.Txt") Then
                gLogMsg "FAIL: mAdjOverlapAgmnts - Unable to delete att code " & tmOverlapInfo(ilLoop).lAttCode, "AffErrorLog.Txt", False
            End If
'            cnn.BeginTrans
'            SQLQuery = "DELETE FROM dat WHERE (datAtfCode = " & tmOverlapInfo(ilLoop).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Att WHERE (AttCode = " & tmOverlapInfo(ilLoop).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            cnn.CommitTrans
        ElseIf (tmOverlapInfo(ilLoop).lOnAirDate < llOnAir) And (llTestEndDate >= llOnAir) Then
            'Terminate the agreement on llOnAir minus 1 as it starts prior to new agreement
            slDrop = Format$(llOnAir, "m/d/yy")
            'cnn.BeginTrans

            SQLQuery = "SELECT cpttStartDate, cpttCode FROM cptt WHERE (cpttAtfCode = " & tmOverlapInfo(ilLoop).lAttCode & " And "
            SQLQuery = SQLQuery & "cpttStartDate >= '" & Format$(slDrop, sgSQLDateForm) & "'" & ")"
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            ilUpper = 0
            While Not rst_Cptt.EOF
                tlCpttArray(ilUpper).lCpttCode = rst_Cptt!cpttCode
                tlCpttArray(ilUpper).sCpttStartDate = rst_Cptt!CpttStartDate
                ilUpper = ilUpper + 1
                ReDim Preserve tlCpttArray(0 To ilUpper)
                rst_Cptt.MoveNext
            Wend
                
            For ilIdx = 0 To ilUpper - 1 Step 1
                '9/25/06: Removed retaining of CPTT as user no longer allowed to set date prior to last posted date
                'ilExported = gCheckIfSpotsHaveBeenExported(imVefCode, tlCpttArray(ilIdx).sCpttStartDate, ilAgreeType)
                ''igChangedNewErased values  1 = changed, 2 = new, 3 = erased
                ''If they are changing an agreement and it's already been exported then don't delete the CPTTs
                'If (igChangedNewErased = 1 Or igChangedNewErased = 2) And (ilExported = False) Then
                    SQLQuery = "DELETE FROM Cptt WHERE (cpttCode = " & tlCpttArray(ilIdx).lCpttCode & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    llAttCode = tmOverlapInfo(ilLoop).lAttCode
                    slStartDate = tlCpttArray(ilIdx).sCpttStartDate
                    slEndDate = DateAdd("d", 6, tlCpttArray(ilIdx).sCpttStartDate)
                    
                    'Doug (9/25/06): Remove Web Spots if using web.  For Martketron, i'm disallowing overlaps if aet exist within date change
                    'Delete the spots from the web first
                    slStr = "Select attExportType, attExportToWeb, attWebInterface from att where attCode = " & llAttCode
                    Set temp_rst = gSQLSelectCall(slStr)
                    If temp_rst.EOF = False Then
                        '7701 test changed
'                        If (temp_rst!attExportType = 1) And ((temp_rst!attExportToWeb = "Y") Or (gIfNullInteger(temp_rst!vatWvtIdCodeLog) = Vendors.Cumulus)) And gHasWebAccess Then
                        If temp_rst!attExportType = 1 And gHasWebAccess Then
                            ilRet = gWebDeleteSpots(llAttCode, Format$(slStartDate, sgSQLDateForm), Format$(slEndDate, sgSQLDateForm))
                        End If
                    End If
                    
                    SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    'cnn.CommitTrans

                    'If (ilAgreeType = 1) Or (ilAgreeType = 2) Then
                    If (ilAgreeType > 0) Then
                        ilRet = gAlertAdd("R", "S", imVefCode, tlCpttArray(ilIdx).sCpttStartDate)
                    End If
                'End If
            Next ilIdx
            slDrop = DateAdd("d", -1, slDrop)
            SQLQuery = "UPDATE att SET "
            If bmShowDates Then
                SQLQuery = SQLQuery & "attDropDate = '" & Format$(slDrop, sgSQLDateForm) & "', "
            Else
                SQLQuery = SQLQuery & "attOffAir = '" & Format$(slDrop, sgSQLDateForm) & "', "
            End If
            SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
            SQLQuery = SQLQuery & " WHERE attCode = " & tmOverlapInfo(ilLoop).lAttCode & ""
            'cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                mMousePointer vbDefault
                gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
                mAdjOverlapAgmnts = False
                Exit Function
            End If
            'cnn.CommitTrans
        End If
    Next ilLoop
    mAdjOverlapAgmnts = True
    Exit Function
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "", "Agreement-mAdjOverAgments"
    mAdjOverlapAgmnts = False
    Exit Function
End Function


Private Sub mPledgeEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llCol As Long
    Dim ilBypass As Integer
    
    ''If optTimeType(0).Value = False And optTimeType(1).Value = False Then
    '2/5/11:  Allow input without avails defined
    'If UBound(tgDat) <= LBound(tgDat) Then
    '    gMsgBox "Either Dayparts or Avails must be selected before entering pledge information", vbOKOnly
    '    Exit Sub
    'End If
    If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
        ilBypass = False
        If (grdPledge.Col = ESTIMATETIMEINDEX) And (grdPledge.CellBackColor <> LIGHTYELLOW) Then
            ilBypass = True
        End If
        If Not ilBypass Then
            gMsgBox "This agreement may not be changed.  Spots have been posted against it." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "The Last Week that Spots were Posted was on " & smLastPostedDate
            Exit Sub
        End If
    End If
    If (grdPledge.Row >= grdPledge.FixedRows) And (grdPledge.Row < grdPledge.Rows) And (grdPledge.Col >= MONFDINDEX) And (grdPledge.Col < grdPledge.Cols - 1) Then
        lmEnableRow = grdPledge.Row
        lmEnableCol = grdPledge.Col
        imShowGridBox = True
        pbcArrow.Move grdPledge.Left - pbcArrow.Width, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + (grdPledge.RowHeight(grdPledge.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdPledge.Col
            Case MONFDINDEX To SUNFDINDEX, MONPDINDEX To SUNPDINDEX 'Days
                'D.S. 11/10/05 Added the 2 for loops and the grdPledge.Col = lmEnableCol
                For llCol = MONFDINDEX To SUNFDINDEX Step 1
                    grdPledge.Col = llCol
                    grdPledge.CellFontName = "Monotype Sorts"
                Next llCol
                For llCol = MONPDINDEX To SUNPDINDEX Step 1
                    grdPledge.Col = llCol
                    grdPledge.CellFontName = "Monotype Sorts"
                Next llCol
                grdPledge.Col = lmEnableCol
                pbcDay.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 15
                If grdPledge.Text = "4" Then
                    ckcDay.Value = vbChecked
                Else
                    ckcDay.Value = vbUnchecked
                End If
                pbcDay.Visible = True
                If ckcDay.Enabled Then
                    ckcDay.SetFocus
                End If
            Case DAYFEDINDEX
                pbcDayFed.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) + 30, grdPledge.RowHeight(grdPledge.Row) - 15
                If grdPledge.Text = "Missing" Then
                    grdPledge.CellForeColor = vbBlack
                    grdPledge.Text = ""
                End If
                If grdPledge.Text = "" Then
                    grdPledge.Text = "A"
                End If
                If pbcDayFed.Height > grdPledge.RowHeight(grdPledge.Row) - 15 Then
                    pbcDayFed.FontName = "Arial"
                    pbcDayFed.Height = grdPledge.RowHeight(grdPledge.Row) - 15
                End If
                pbcDayFed.Visible = True
                If pbcDayFed.Enabled Then
                    pbcDayFed.SetFocus
                End If
            Case STARTTIMEFDINDEX, ENDTIMEFDINDEX, STARTTIMEPDINDEX, ENDTIMEPDINDEX  'Date
                txtTime.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 15
                If grdPledge.Text <> "Missing" Then
                    txtTime.Text = grdPledge.Text
                Else
                    txtTime.Text = ""
                End If
                If txtTime.Height > grdPledge.RowHeight(grdPledge.Row) - 15 Then
                    txtTime.FontName = "Arial"
                    txtTime.Height = grdPledge.RowHeight(grdPledge.Row) - 15
                End If
                txtTime.Visible = True
                If txtTime.Enabled Then
                    txtTime.SetFocus
                End If
            Case STATUSINDEX
                If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
                    gMsgBox "This agreement may not be changed.  Spots have been posted against it." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "The Last Week that Spots were Posted was on " & smLastPostedDate
                    Exit Sub
                End If
                'txtDropdown.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 15
                txtDropdown.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) + 2 * grdPledge.ColWidth(grdPledge.Col + 1), grdPledge.RowHeight(grdPledge.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcStatus, 4
                slStr = grdPledge.Text
                ilIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcStatus.ListIndex = ilIndex
                Else
                    lbcStatus.ListIndex = 0
                End If
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                If txtDropdown.Height > grdPledge.RowHeight(grdPledge.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPledge.RowHeight(grdPledge.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                If txtDropdown.Enabled Then
                    txtDropdown.SetFocus
                End If
            Case AIRPLAYINDEX
                txtAirPlay.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col), grdPledge.RowHeight(grdPledge.Row) - 15
                slStr = grdPledge.Text
                txtAirPlay.Text = slStr
                txtAirPlay.Visible = True
                If txtAirPlay.Enabled Then
                    txtAirPlay.SetFocus
                End If
            Case ESTIMATETIMEINDEX
                grdPledge.CellForeColor = vbBlack
                lmPdSTime = gTimeToLong(Trim$(grdPledge.TextMatrix(grdPledge.Row, STARTTIMEPDINDEX)), False)
                lmPdETime = gTimeToLong(Trim$(grdPledge.TextMatrix(grdPledge.Row, ENDTIMEPDINDEX)), True)
                If grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = "" Then
                    grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = "-1"
                End If
                mPopEstTimes
                frcET.Visible = True
                'If Trim$(grdET.TextMatrix(grdET.FixedRows, ETDAYINDEX)) = "" Then
                    pbcETSTab.SetFocus
                'End If
            Case EMBEDDEDORROSINDEX
                pbcEmbeddedOrROS.Move grdPledge.Left + imColPos(grdPledge.Col) + grdPledge.ColWidth(grdPledge.Col) - pbcEmbeddedOrROS.Width, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, pbcEmbeddedOrROS.Width, grdPledge.RowHeight(grdPledge.Row) - 15
                If grdPledge.Text = "Missing" Then
                    grdPledge.CellForeColor = vbBlack
                    grdPledge.Text = ""
                End If
                If Trim$(grdPledge.Text) = "" Then
                    grdPledge.Text = smDefaultEmbeddedOrROS
                End If
                If pbcEmbeddedOrROS.Height > grdPledge.RowHeight(grdPledge.Row) - 15 Then
                    pbcEmbeddedOrROS.FontName = "Arial"
                    pbcEmbeddedOrROS.Height = grdPledge.RowHeight(grdPledge.Row) - 15
                End If
                pbcEmbeddedOrROS.Visible = True
                If pbcEmbeddedOrROS.Enabled Then
                    pbcEmbeddedOrROS.SetFocus
                End If
        End Select
        imcTrash.Visible = True
    End If
End Sub

Private Sub mPledgeSetShow()
    Dim llRow As Long
    Dim slStr As String

    If (lmEnableRow >= grdPledge.FixedRows) And (lmEnableRow < grdPledge.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        mSetPledgeFromStatus lmEnableRow, lmEnableCol
    End If
    If lmEnableCol = ESTIMATETIMEINDEX Then
        slStr = ""
        For llRow = grdET.FixedRows To grdET.Rows - 1 Step 1
            If Trim$(grdET.TextMatrix(llRow, ETTIMEINDEX)) <> "" Then
                If slStr = "" Then
                    slStr = grdET.TextMatrix(llRow, ETDAYINDEX) & ":" & grdET.TextMatrix(llRow, ETTIMEINDEX)
                Else
                    slStr = slStr & ", " & grdET.TextMatrix(llRow, ETDAYINDEX) & ":" & grdET.TextMatrix(llRow, ETTIMEINDEX)
                End If
            End If
        Next llRow
        grdPledge.TextMatrix(lmEnableRow, lmEnableCol) = slStr
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imShowGridBox = False
    pbcArrow.Visible = False
    pbcDay.Visible = False
    pbcDayFed.Visible = False
    pbcEmbeddedOrROS.Visible = False
    txtTime.Visible = False
    txtDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcStatus.Visible = False
    frcET.Visible = False
    txtAirPlay.Visible = False
    imcTrash.Visible = False
End Sub

Private Function mTestZone() As Integer

    Dim ilVef As Integer
    Dim ilFoundZone As Integer
    Dim ilZone As Integer
    Dim ilRet As Integer
    Dim iZoneDefined As Integer
    
    If smStationZone = "" Then
        ilFoundZone = False
        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilVef).iCode = imVefCode Then
                For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                    If (Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) <> "") And (Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) <> "~~~") Then
                        ilFoundZone = True
                        Exit For
                    End If
                Next ilZone
                Exit For
            End If
        Next ilVef
        If ilFoundZone Then
            ilRet = gMsgBox("Warning: Station Missing Zone and Vehicle has Zone" & Chr$(13) & Chr$(10) & "Proceed With Save?", vbYesNo)
            If ilRet = vbNo Then
                mTestZone = False
                Exit Function
            End If
             
        End If
    Else
        ilFoundZone = False
        iZoneDefined = False
        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilVef).iCode = imVefCode Then
                For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                    If (Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) <> "") And (Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) <> "~~~") Then
                        iZoneDefined = True
                    End If
                    If StrComp(smStationZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                        ilFoundZone = True
                        Exit For
                    End If
                Next ilZone
                Exit For
            End If
        Next ilVef
        If (Not ilFoundZone) And (iZoneDefined) Then
            ilRet = gMsgBox("Warning: Station Zone and Vehicle Zone Don't Match" & Chr$(13) & Chr$(10) & "Proceed With Save?", vbYesNo)
            If ilRet = vbNo Then
                mTestZone = False
                Exit Function
            End If
        End If
    End If
    mTestZone = True
End Function

Private Function mMoveDaypart()
    Dim iLoop As Integer
    Dim iPack As Integer
    Dim sTime As String
    Dim iTRow As Integer
    Dim iDay As Integer
    Dim iIndex As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilIndex As Integer
    Dim llRowIndex As Long
    Dim ilFirstFdDay As Integer
    Dim ilFirstPdDay As Integer
    Dim ilNext As Integer
    
    grdPledge.Redraw = False
'    'Test if fields defined
'    grdDayparts.MoveFirst
'    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
'        sTime = grdPledge.TextMatrix(llRow, 7)
'        If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
'            grdPledge.Redraw = True
'            gSendKeys "%I", True
'            Beep
'            mMousePointer vbDefault
'            mMoveDaypart = False
'            Exit Function
'        End If
'        sTime = grdPledge.TextMatrix(llRow, 8)
'        If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then    'Time not valid.
'            grdPledge.Redraw = True
'            gSendKeys "%I", True
'            Beep
'            mMousePointer vbDefault
'            mMoveDaypart = False
'            Exit Function
'        End If
'        If grdDayparts.Columns(9).ListIndex >= 0 Then
'            iIndex = grdDayparts.Columns(9).ItemData(grdDayparts.Columns(9).ListIndex)
'            If tmStatusTypes(iIndex).iPledged = 1 Then
'                sTime = grdDayparts.Columns(17).Text
'                If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
'                    grdDayparts.Redraw = True
'                    gSendKeys "%I", True
'                    Beep
'                    grdDayparts.Columns(17).CellStyleSet ("oops"), grdDayparts.Row
'                    mMousePointer vbDefault
'                    mMoveDaypart = False
'                    Exit Function
'                End If
'                sTime = grdDayparts.Columns(18).Text
'                If gIsTime(sTime) = False Then   'Time not valid.
'                    grdDayparts.Redraw = True
'                    gSendKeys "%I", True
'                    Beep
'                    grdDayparts.Columns(18).CellStyleSet ("oops"), grdDayparts.Row
'                    mMousePointer vbDefault
'                    mMoveDaypart = False
'                    Exit Function
'                End If
'            End If
'        End If
'        grdDayparts.MoveNext
'    Next iLoop
'    'Move Values into array
'    grdDayparts.MoveFirst
'    For iLoop = 0 To UBound(tgDat) - 1 Step 1
    ReDim tgDat(0 To grdPledge.Rows - grdPledge.FixedRows - 1) As DAT
    iLoop = 0
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            ilFirstFdDay = -1
            For iDay = MONFDINDEX To SUNFDINDEX Step 1
                If Trim$(grdPledge.TextMatrix(llRow, iDay)) <> "" Then
                    tgDat(iLoop).iFdDay(iDay) = 1
                    If ilFirstFdDay = -1 Then
                        ilFirstFdDay = iDay
                    End If
                Else
                    tgDat(iLoop).iFdDay(iDay) = 0
                End If
            Next iDay
            sTime = grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX)
            If sTime <> "" Then
                sTime = gConvertTime(sTime)
    '            If Second(sTime) = 0 Then
    '                grdDayparts.Columns(7).Text = Format$(sTime, sgShowTimeWOSecForm)
    '            Else
    '                grdDayparts.Columns(7).Text = Format$(sTime, sgShowTimeWSecForm)
    '            End If
            End If
            tgDat(iLoop).sFdSTime = sTime
            sTime = grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX)
            sTime = gConvertTime(sTime)
    '        If Second(sTime) = 0 Then
    '            grdDayparts.Columns(8).Text = Format$(sTime, sgShowTimeWOSecForm)
    '        Else
    '            grdDayparts.Columns(8).Text = Format$(sTime, sgShowTimeWSecForm)
    '        End If
            tgDat(iLoop).sFdETime = sTime
            slStr = grdPledge.TextMatrix(llRow, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tmStatusTypes(iIndex).iPledged = 0 Then  'live
                    tgDat(iLoop).iFdStatus = tmStatusTypes(iIndex).iStatus
                    For iDay = MONFDINDEX To SUNFDINDEX Step 1
                        If Trim$(grdPledge.TextMatrix(llRow, iDay)) <> "" Then
                            tgDat(iLoop).iPdDay(iDay) = 1
                        Else
                            tgDat(iLoop).iPdDay(iDay) = 0
                        End If
                    Next iDay
                    tgDat(iLoop).sPdDayFed = ""
                    tgDat(iLoop).sPdSTime = tgDat(iLoop).sFdSTime
                    tgDat(iLoop).sPdETime = tgDat(iLoop).sFdETime
                ElseIf tmStatusTypes(iIndex).iPledged = 1 Then  'delayed
                    tgDat(iLoop).iFdStatus = tmStatusTypes(iIndex).iStatus
                    ilFirstPdDay = -1

                    For iDay = MONPDINDEX To SUNPDINDEX Step 1
                        If Trim$(grdPledge.TextMatrix(llRow, iDay)) <> "" Then
                            tgDat(iLoop).iPdDay(iDay - MONPDINDEX) = 1
                            If ilFirstPdDay = -1 Then
                                ilFirstPdDay = iDay
                            End If
                        Else
                            tgDat(iLoop).iPdDay(iDay - MONPDINDEX) = 0
                        End If
                    Next iDay
                    '11/26/14: Handle Monday feed to Sunday pledge
                    'If ilFirstPdDay - MONPDINDEX < ilFirstFdDay - MONFDINDEX Then
                    If (ilFirstPdDay - MONPDINDEX < ilFirstFdDay - MONFDINDEX) Or ((ilFirstPdDay - MONPDINDEX > ilFirstFdDay - MONFDINDEX) And (Trim$(grdPledge.TextMatrix(llRow, DAYFEDINDEX)) <> "")) Then
                        tgDat(iLoop).sPdDayFed = grdPledge.TextMatrix(llRow, DAYFEDINDEX)
                    ElseIf ilFirstPdDay - MONPDINDEX > ilFirstFdDay - MONFDINDEX Then
                        tgDat(iLoop).sPdDayFed = "A"
                    Else
                        tgDat(iLoop).sPdDayFed = ""
                    End If
                    sTime = grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX)
                    sTime = gConvertTime(sTime)
    '                If Second(sTime) = 0 Then
    '                    grdDayparts.Columns(17).Text = Format$(sTime, sgShowTimeWOSecForm)
    '                Else
    '                    grdDayparts.Columns(17).Text = Format$(sTime, sgShowTimeWSecForm)
    '                End If
                    tgDat(iLoop).sPdSTime = sTime
                    sTime = grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX)
                    If Len(Trim$(sTime)) = 0 Then
                        sTime = grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX)
                    End If
                    sTime = gConvertTime(sTime)
    '                If Second(sTime) = 0 Then
    '                    grdDayparts.Columns(18).Text = Format$(sTime, sgShowTimeWOSecForm)
    '                Else
    '                    grdDayparts.Columns(18).Text = Format$(sTime, sgShowTimeWSecForm)
    '                End If
                    tgDat(iLoop).sPdETime = sTime
                    
                ElseIf tmStatusTypes(iIndex).iPledged = 2 Then
                    tgDat(iLoop).iFdStatus = tmStatusTypes(iIndex).iStatus
                    For iDay = 0 To 6 Step 1
                        tgDat(iLoop).iPdDay(iDay) = 0
                    Next iDay
                    tgDat(iLoop).sPdDayFed = ""
                    tgDat(iLoop).sPdSTime = ""
                    tgDat(iLoop).sPdETime = ""
                ElseIf tmStatusTypes(iIndex).iPledged = 3 Then
                    tgDat(iLoop).iFdStatus = tmStatusTypes(iIndex).iStatus
                    For iDay = 0 To 6 Step 1
                        tgDat(iLoop).iPdDay(iDay) = 0
                    Next iDay
                    tgDat(iLoop).sPdDayFed = ""
                    tgDat(iLoop).sPdSTime = ""
                    tgDat(iLoop).sPdETime = ""
                End If
            End If
            tgDat(iLoop).iAirPlayNo = Val(grdPledge.TextMatrix(llRow, AIRPLAYINDEX))
            tgDat(iLoop).sEstimatedTime = "N"
            If Trim$(grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX)) <> "" Then
                ilNext = grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX)
                Do While ilNext <> -1
                    If (Trim$(tmETAvailInfo(ilNext).sETTime) <> "") Or (Trim$(tmETAvailInfo(ilNext).sETDay) <> "") Then
                        tgDat(iLoop).sEstimatedTime = "Y"
                        Exit Do
                    End If
                    ilNext = tmETAvailInfo(ilNext).iNextET
                Loop
            End If
            If tgDat(iLoop).sEstimatedTime = "Y" Then
                tgDat(iLoop).iFirstET = Val(grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX))
            Else
                tgDat(iLoop).iFirstET = -1
            End If
            '7/15/14
            If Trim$(grdPledge.TextMatrix(llRow, EMBEDDEDORROSINDEX)) = "" Then
                tgDat(iLoop).sEmbeddedOrROS = smDefaultEmbeddedOrROS
            Else
                tgDat(iLoop).sEmbeddedOrROS = grdPledge.TextMatrix(llRow, EMBEDDEDORROSINDEX)
            End If
            iLoop = iLoop + 1
        End If
    Next llRow
    If iLoop > 0 Then
        imDatLoaded = True
    End If
    ReDim Preserve tgDat(0 To iLoop) As DAT
'    grdDayparts.FirstRow = grdDayparts.AddItemBookmark(iTRow)
'    grdDayparts.Redraw = True
'    iLoop = UBound(tgDat) - 1
'    Do While (iLoop >= 0) And (UBound(tgDat) > 0)
'        If (Trim$(tgDat(iLoop).sFdSTime) = "") Then
'            For iPack = iLoop To UBound(tgDat) - 1 Step 1
'                tgDat(iPack) = tgDat(iPack + 1)
'            Next iPack
'            ReDim Preserve tgDat(0 To UBound(tgDat) - 1) As DAT
'        Else
'            iLoop = iLoop - 1
'        End If
'    Loop
    grdPledge.Redraw = True
    mMoveDaypart = True
End Function





Private Sub ClearControls()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llVpf As Long
    
    '8418
    mLoadAllowedVendorServices imShttCode
    smContractPDFSubFolder = ""
    If sgContractPDFPath = "" Then
        lacContractPDF.Enabled = False
        cbcContractPDF.Enabled False
        cmcBrowse.Enabled = False
        lacPDFPath.Caption = "PDF Path: Not defined in Affliat.ini"
    Else
        lacPDFPath.Caption = "PDF Path: " '& sgContractPDFPath
    End If
    lacAttCode.Caption = ""
    smSvOnAirDate = ""
    smSvOffAirDate = ""
    smSvDropDate = ""
    smMktronActiveDate = "1/1/1970"
    smSvMulticast = ""
    smLastPostedDate = "1/1/1970"
    lmAttCode = 0
    txtStartDate.Text = ""
    txtEndDate.Text = ""
    txtOnAirDate.Text = ""
    txtOffAirDate.Text = ""
    txtLdMult.Text = "1"        'Air play used instead of load factor
    lacLdMult.Visible = False
    txtLdMult.Visible = False
    txtNoCDs.Text = "1"
    cbcAirPlayNo.Clear
    cbcAirPlayNo.AddItem "[All]"
    cbcAirPlayNo.AddItem "1"
    cbcAirPlayNo.ListIndex = 0
    txtDays.Text = ""
    optPrintCP(0).Value = True
    optCarryCmml(0).Value = True
    optSuppressNotices(1).Value = True
    optNCR(1).Value = True                  '7-6-09 ncr agreeement flag
    optFormerNCR(1).Value = True            '7-6-09  former NCR offender
    smPreviousNCR = ""
    optService(1).Value = True  '10/28/14: Service agreement
    'optSendTape(0).Value = True
    optSendTape(1).Value = True
    optBarCode(1).Value = True
    edcNoAirPlays.Text = "1"
    'optTimeType(0).Value = False
    'optTimeType(1).Value = False
    sgAirPlay1TimeType = ""
    sgCDStartTime = ""
    ckcProhibitSplitCopy.Value = vbUnchecked
    optComp(0).Value = True
    optPost(0).Value = True
    optSigned(0).Value = True
    txtRetDate.Text = ""
    txtACName.Text = ""
    txtACPhone.Text = ""
    txtLog.Text = ""
    txtCP.Text = ""
    txtOther.Text = ""
    txtComments.Text = ""
    txtDropDate.Text = ""
    txtLogPassword.Text = ""
    txtEmailAddr.Text = ""
    rbcExportType(0).Value = False
    rbcExportType(1).Value = False
    'rbcExportType(2).Value = False
    'rbcExportType(3).Value = False
    ckcExportTo(0).Value = vbUnchecked
    ckcExportTo(1).Value = vbUnchecked
    ckcExportTo(2).Value = vbUnchecked
    ckcExportTo(3).Value = vbUnchecked
    ckcExportTo(4).Value = vbUnchecked
    ckcExportTo(5).Value = vbUnchecked
    ckcExportTo(0).Enabled = False
    ckcExportTo(1).Enabled = False
    ckcExportTo(2).Enabled = False
    ckcExportTo(3).Enabled = False
    '6592
    ckcExportTo(4).Value = vbUnchecked
    ckcExportTo(4).Enabled = False
    
    ckcExportTo(6).Value = vbUnchecked
    ckcExportTo(6).Enabled = False
    rbcAudio(5).Value = True
    ckcSendDelays.Value = vbUnchecked
    ckcSendNotCarried.Value = vbUnchecked
    'If Not gUsingWeb Then
    '    'rbcExportType(1).Enabled = False
    '    'rbcExportType(3).Enabled = False
    '    ckcExportTo(0).Enabled = False
    '    ckcExportTo(1).Enabled = False
    'End If
    'If Not gUsingUnivision Then
    '    'rbcExportType(2).Enabled = False
    '    ckcExportTo(2).Enabled = False
    'End If
    If (Not gUsingWeb) And (Not gUsingUnivision) Then
        rbcExportType(0).Value = True
        frcPosting.Visible = True
    Else
        frcPosting.Visible = False
    End If
    '7701
    frcLogDelivery.Visible = False
    frcAudioDelivery.Visible = True
    rbcLogType(0).Value = False
    rbcLogType(1).Value = False
    rbcLogType(2).Value = False
    rbcPostType(0).Value = False
    rbcPostType(1).Value = False
    rbcPostType(2).Value = False
    rbcSendLogEMail(0).Value = True
    rbcSendLogEMail(1).Value = False
    rbcMulticast(0).Value = False
    rbcMulticast(1).Value = True
    
    frcLogType.Visible = False
    frcPostType.Visible = False
    '7701 removed
    'frcSendLogEMail.Visible = False
    lblWebPW.Visible = False
    lblWebEmail.Visible = False
    txtLogPassword.Visible = False
    cmdGenPassword.Visible = False
    txtEmailAddr.Visible = False
    optRadarClearType(2).Value = True
    'Replaced by Market Rep, attMktRepUstCode
    cboAffAE.ListIndex = -1
    cboAffAE.Text = ""
    optVoiceTracked(1).Value = True
    txtXDReceiverID.Text = ""
    txtIDCReceiverID.Text = ""
    smIDCReceiverID = ""
    cbcMarketRep.SetListIndex = -1
    cbcServiceRep.SetListIndex = -1
    cbcContractPDF.SetListIndex = -1
    '7701 removed
'    lacIDCReceiverID.Visible = False
'    txtIDCReceiverID.Visible = False
'    frcIdcGroup.Visible = False
'    frcVoiceTracked.Visible = False
'    lacXDReceiverID.Visible = False
'    txtXDReceiverID.Visible = False
    '6466
    optIDCGroup(0).Value = True
    mIDCShowGroup False
    mGetShttInfo
    sgVehProgStartTime = ""
    sgVehProgEndTime = ""
    lacPrgTimes.Caption = ""
    ilRet = gDetermineAgreementTimes(imShttCode, imVefCode, "1/1/1970", "12/31/2069", "12/31/2069", "", sgVehProgStartTime, sgVehProgEndTime)
    If sgVehProgStartTime <> "" Then
        sgVehProgStartTime = gCompactTime(sgVehProgStartTime)
        lacPrgTimes.Caption = "Program Times: " & sgVehProgStartTime
        If sgVehProgEndTime <> "" Then
            sgVehProgEndTime = gCompactTime(sgVehProgEndTime)
            lacPrgTimes.Caption = lacPrgTimes.Caption & "-" & sgVehProgEndTime
        End If
    End If
    imFieldChgd = False
    If sgUstWin(2) = "I" Then
        cmdSave.Enabled = False
        cmdNew.Enabled = True
    End If
    
    smDefaultEmbeddedOrROS = "R"
    llVpf = gBinarySearchVpf(CLng(imVefCode))
    If llVpf <> -1 Then
        smDefaultEmbeddedOrROS = tgVpfOptions(llVpf).sEmbeddedOrROS
    End If
    If Trim$(smDefaultEmbeddedOrROS) = "" Then
        smDefaultEmbeddedOrROS = "R"
    End If
    ReDim tgDat(0 To 0) As DAT
    mClearPledgeGrid
    ReDim tgAirPlaySpec(0 To 0) As AIRPLAYSPEC
    ReDim tgBreakoutSpec(0 To 0) As BREAKOUTSPEC
    ReDim tgDPSelection(0 To 0) As DPSELECTION
    ReDim tmETAvailInfo(0 To 0) As ETAVAILINFO
    ReDim tmPetInfo(0 To 0) As PETINFO
    imDatLoaded = False
    '8/12/16: Separtated Pledge and Estimate so that est can be saved in the past
    bmPledgeDataChgd = False
    bmETDataChgd = False
    mClearEventGrid
    cbcSeason.Clear
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 4 'Clear
    udcContactGrid.Action 3 'populate
    If (imTabIndex = TABPLEDGE) And (optExAll(1).Value = True) Then
        bmBypassTestDP = True
        ''TabStrip1.SetFocus
        'gSendKeys "%M", False
        'gSendKeys "%M", False
        TabStrip1.Tabs(TABMAIN).Selected = True
    End If
    '7701
    'default email to 'yes'
    rbcSendLogEMail(0).Value = True
    frcSendLogEMail.Visible = True
    mSetMultilist lbcLogDelivery, 0
    mSetMultilist lbcAudioDelivery, 0
    mEnableDeliveryOptions True
    mEnableDeliveryOptions False
    
    '4/3/19
    ckcExcludeFillSpot.Value = vbUnchecked
    ckcExcludeCntrTypeQ.Value = vbUnchecked
    ckcExcludeCntrTypeR.Value = vbUnchecked
    ckcExcludeCntrTypeT.Value = vbUnchecked
    ckcExcludeCntrTypeM.Value = vbUnchecked
    ckcExcludeCntrTypeS.Value = vbUnchecked
    ckcExcludeCntrTypeV.Value = vbUnchecked

End Sub
Private Sub BindControls()
    Dim llLoop As Long
    '7701
    Dim rstVat As ADODB.Recordset
    Dim slattExportToMarketron As String
    Dim slattExportToCBS As String
    Dim slattExportToClearCh As String
   ' Dim slattExportToJelli As String
    '9452
    bmSendNotCarriedChange = False
    
    slattExportToMarketron = "N"
    slattExportToClearCh = "N"
    slattExportToCBS = "N"
    On Error GoTo ErrHand
    
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
        StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        lacAttCode.Caption = rst!attCode
        lacAttCode.Visible = True
    Else
        lacAttCode.Visible = False
    End If
    'frcMulticast.Visible = False
    'lacMulticast.Visible = False
    'lbcMulticast.Visible = False
    ''If attMulticast is defined then set it.  Otherwise default is set to False
    'If rst!attMulticast = "Y" Then
    '    rbcMulticast(0).Value = True
    'Else
    '    rbcMulticast(1).Value = True
    'End If
    
    cbcAirPlayNo.Clear
    cbcAirPlayNo.AddItem "[All]"
    cbcAirPlayNo.AddItem "1"
    cbcAirPlayNo.ListIndex = 0
    
    'If rst!attPledgeType = "D" Then
    '    optTimeType(0).Value = True
    'ElseIf rst!attPledgeType = "A" Then
    '    optTimeType(1).Value = True
    ''ElseIf rst!attPledgeType = "C" Then
    ''    optTimeType(2).Value = True
    'Else
    '    optTimeType(0).Value = False
    '    optTimeType(1).Value = False
    '    optTimeType(2).Value = False
    'End If
    sgAirPlay1TimeType = rst!attPledgeType
    
    'If rst!attWebInterface = "C" Then
    '    If gUsingWeb Then
    '        rbcExportType(3).Value = True
    '    End If
    'Else
    '    If rst!attExportType = 0 Then
    '        rbcExportType(0).Value = True
    '    ElseIf rst!attExportType = 1 And gUsingWeb Then
    '        rbcExportType(1).Value = True
    '    ElseIf rst!attExportType = 2 And gUsingUnivision Then
    '        rbcExportType(2).Value = True
    '    End If
    'End If
    If rst!attExportType = 0 Then
        rbcExportType(0).Value = True
        ckcExportTo(0).Value = vbUnchecked
        '7701
'        ckcExportTo(1).Value = vbUnchecked
'        ckcExportTo(1).Enabled = False
    Else
        rbcExportType(1).Value = True
        If gUsingWeb Then
        '7701
'            ckcExportTo(1).Enabled = True
'            If rst!attWebInterface = "C" Then
'                ckcExportTo(1).Value = vbChecked
'            '    ckcExportTo(0).Value = vbChecked
'            End If
'            '6592 CBS
'            If rst!attExportToCBS = "Y" Then
'                ckcExportTo(4).Value = vbChecked
'            End If
            'not 7701
            'If rst!attExportToWeb = "Y" Then
                ckcExportTo(0).Value = vbChecked
            'End If
        Else
            '7701
           ' ckcExportTo(1).Enabled = False
            ckcExportTo(0).Value = vbUnchecked
           ' ckcExportTo(1).Value = vbUnchecked
        End If
        '7701
'        If rst!attExportToUnivision = "Y" Then
'            ckcExportTo(2).Value = vbChecked
'        End If
        '8000
        If rst!attExportToUnivision = "Y" Then
            ckcUnivision.Value = vbChecked
        Else
            ckcUnivision.Value = vbUnchecked
        End If
'        If rst!attExportToMarketron = "Y" Then
'            ckcExportTo(3).Value = vbChecked
'        End If
'        If rst!attExportToJelli = "Y" Then
'            ckcExportTo(6).Value = vbChecked
'        End If
    End If
    '7701
'    rbcAudio(5).Value = True
'    If rst!attAudioDelivery = "X" Then
'        rbcAudio(0).Value = True
'    ElseIf rst!attAudioDelivery = "B" Then
'        rbcAudio(1).Value = True
'    ElseIf rst!attAudioDelivery = "W" Then
'        rbcAudio(2).Value = True
'    ElseIf rst!attAudioDelivery = "I" Then
'        rbcAudio(3).Value = True
'    ElseIf rst!attAudioDelivery = "P" Then
'        rbcAudio(4).Value = True
'    End If

    '7/17/19: Move below after setting vendors
    'If rst!attSendDelayToXDS = "Y" Then
    '    ckcSendDelays.Value = vbChecked
    'Else
    '    ckcSendDelays.Value = vbUnchecked
    'End If
    'If rst!attXDSSendNotCarry = "Y" Then
    '    ckcSendNotCarried.Value = vbChecked
    'Else
    '    ckcSendNotCarried.Value = vbUnchecked
    'End If
    If frcLogType.Visible = True Then
        If rst!attLogType = 0 Then
            rbcLogType(0).Value = True
        ElseIf rst!attLogType = 1 Then
            rbcLogType(1).Value = True
        ElseIf rst!attLogType = 2 Then
            rbcLogType(2).Value = True
        End If
    End If
    
    If frcPostType.Visible = True Then
        If rst!attPostType = 0 Then
            rbcPostType(0).Value = True
        ElseIf rst!attPostType = 1 Then
            rbcPostType(1).Value = True
        ElseIf rst!attPostType = 2 Then
            rbcPostType(2).Value = True
        End If
    End If

    If rst!attSendLogEmail = 0 Then
        rbcSendLogEMail(0).Value = True
    ElseIf rst!attSendLogEmail = 1 Then
        rbcSendLogEMail(1).Value = True
    End If

    'txtLogPassword.Text = Trim$(rst!attWebPW)
    'txtEmailAddr.Text = Trim$(rst!attWebEmail)
    'txtProgContProv.Text = Trim(rst!attPgContProv)
    'txtCommAudProv.Text = Trim$(rst!attCommAudProv)

    ' Save the current values so we can detect if any of them get changed.
    smInitialattWebEmail = txtEmailAddr.Text
    smInitialattWebPW = txtLogPassword.Text
    If frcLogType.Visible = True Then
        smInitialattLogType = rst!attLogType
    End If
    If frcPostType.Visible = True Then
        smInitialattPostType = rst!attPostType
    End If

    txtLabelID.Text = Trim$(rst!attLabelID)
    txtShipInfo.Text = Trim$(rst!attLabelShipInfo)
    
    'If it's undefined set it to NO
    If rst!attMonthlyWebPost = "Y" Then
        smMonthlyWebPost = "Y"
        chkMonthlyWebPost.Value = vbChecked
    Else
        smMonthlyWebPost = "N"
        chkMonthlyWebPost.Value = vbUnchecked
    End If
    
    lmAttCode = rst!attCode  'rst(0).Value
    imShttCode = rst!attshfCode
    imVefCode = rst!attvefCode
    '8418
    mLoadAllowedVendorServices imShttCode
    SQLQuery = "SELECT * FROM VAT_Vendor_Agreement WHERE vatattcode = " & rst!attCode
    Set rstVat = gSQLSelectCall(SQLQuery)
    Do While Not rstVat.EOF
        If Not mSetMultilist(lbcLogDelivery, rstVat!vatwvtvendorid) Then
            mSetMultilist lbcAudioDelivery, rstVat!vatwvtvendorid
        End If
        Select Case rstVat!vatwvtvendorid
            Case Vendors.cBs
                slattExportToCBS = "Y"
            Case Vendors.iHeart
                slattExportToClearCh = "Y"
            Case Vendors.NetworkConnect
                slattExportToMarketron = "Y"
        End Select
        rstVat.MoveNext
    Loop
    
    '7/17/19: Moved from above
    If rst!attSendDelayToXDS = "Y" Then
        ckcSendDelays.Value = vbChecked
    Else
        ckcSendDelays.Value = vbUnchecked
    End If
    If rst!attXDSSendNotCarry = "Y" Then
        ckcSendNotCarried.Value = vbChecked
    Else
        ckcSendNotCarried.Value = vbUnchecked
    End If

    mSetEventTitles

    mGetShttInfo
        
    'frcMulticast.Visible = False
    lacMulticast(0).Visible = False
    lacMulticast(1).Visible = False
    'lbcMulticast.Visible = False
    grdMulticast.Visible = False
    smSvMulticast = "N"
    If gIsMulticast(imShttCode) Then
        'frcMulticast.Visible = True
        mSetMulticast
        'If attMulticast is defined then set it.  Otherwise default is set to False
        If rst!attMulticast = "Y" Then
            rbcMulticast(0).Value = True
            smSvMulticast = "Y"
        Else
            rbcMulticast(1).Value = True
        End If
    'Else
    '    frcMulticast.Visible = True
    End If
    
    cbcMarketRep.SetListIndex = -1
    For llLoop = 1 To cbcMarketRep.ListCount - 1 Step 1
        If cbcMarketRep.GetItemData(CInt(llLoop)) = rst!attMktRepUstCode Then
            cbcMarketRep.SetListIndex = llLoop
            Exit For
        End If
    Next llLoop
    
    cbcServiceRep.SetListIndex = -1
    For llLoop = 1 To cbcServiceRep.ListCount - 1 Step 1
        If cbcServiceRep.GetItemData(CInt(llLoop)) = rst!attServRepUstCode Then
            cbcServiceRep.SetListIndex = llLoop
            Exit For
        End If
    Next llLoop
    
    
    'Replaced by Market Rep, attMktRepUstCode
    cboAffAE.ListIndex = -1
    cboAffAE.Text = ""
    'If rst!attArttCode > 0 Then
    '    For llLoop = 0 To cboAffAE.ListCount - 1 Step 1
    '        If rst!attArttCode = cboAffAE.ItemData(llLoop) Then
    '            cboAffAE.ListIndex = llLoop
    '            cboAffAE.Text = cboAffAE.List(llLoop)
    '            Exit For
    '        End If
    '    Next llLoop
    'End If
    
    If (DateValue(gAdjYear(rst!attAgreeStart)) = DateValue("1/1/1970")) Or (DateValue(gAdjYear(rst!attAgreeStart)) = DateValue("1/1/70")) Then    'Placeholder value to prevent using Nulls/outer joins
        txtStartDate.Text = ""
    Else
        txtStartDate.Text = Format$(Trim$(rst!attAgreeStart), sgShowDateForm)
    End If
    If (DateValue(gAdjYear(rst!attAgreeEnd)) = DateValue("12/31/2069")) Or (DateValue(gAdjYear(rst!attAgreeEnd)) = DateValue("12/31/69")) Then
        txtEndDate.Text = ""
    Else
        txtEndDate.Text = Format$(rst!attAgreeEnd, sgShowDateForm)
    End If
    If (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/1970")) Then  'Or (rst!attOnAir = "1/1/70") Then
        txtOnAirDate.Text = ""
        imChgOnAirDate = False
    Else
        txtOnAirDate.Text = Format$(rst!attOnAir, sgShowDateForm)
        imChgOnAirDate = False
    End If
    smSvOnAirDate = rst!attOnAir
    If (DateValue(gAdjYear(rst!attOffAir)) = DateValue("12/31/2069")) Then   'Or (rst!attOffAir = "12/31/69") Then
        txtOffAirDate.Text = ""
    Else
        txtOffAirDate.Text = Format$(rst!attOffAir, sgShowDateForm)
    End If
    smSvOffAirDate = rst!attOffAir
    txtLdMult.Text = Trim$(rst!attLoad)
    If Val(Trim$(rst!attLoad)) <= 1 Then
        lacLdMult.Visible = False
        txtLdMult.Visible = False
    Else
        lacLdMult.Visible = True
        txtLdMult.Visible = True
    End If
    txtNoCDs.Text = Trim$(rst!attNoCDs)
    txtDays.Text = Trim$(rst!attNotice)
    If rst!attPrintCP <> -1 Then
        optPrintCP(rst!attPrintCP).Value = True
    Else
        optPrintCP(0).Value = True
    End If
    If IsNull(rst!attSuppressNotice) Then
        optSuppressNotices(1).Value = True
    Else
        If rst!attSuppressNotice <> "Y" Then
            optSuppressNotices(1).Value = True
        Else
            optSuppressNotices(0).Value = True
        End If
    End If
    
    '7-6-09
    If IsNull(rst!attNCR) Then          'NCR field never updated
        optNCR(1).Value = True         'not NCR
        optFormerNCR(1).Value = True    'not repeat offender
        smPreviousNCR = "N"             'retain previous value of attncr
    Else
        If rst!attNCR = "N" Or Trim$(rst!attNCR) = "" Then        'not NCR
            optNCR(1).Value = True      'set control to show NO
            'optFormerNCR(1).Value = True            'not former offender
            smPreviousNCR = "N"          'retain previous value of attncr
            If rst!attFormerNCR = "N" Or Trim$(rst!attFormerNCR) = "" Or IsNull(rst!attFormerNCR) Then 'not former NCR
                optFormerNCR(1).Value = True   'set flag not former offender
            Else
                optFormerNCR(0).Value = True
            End If
        Else                            'already an NCR agreement
            smPreviousNCR = "Y"         'retain previous value of attncr
            optNCR(0).Value = True
            If rst!attFormerNCR = "N" Or Trim$(rst!attFormerNCR) = "" Or IsNull(rst!attFormerNCR) Then 'not former NCR
                optFormerNCR(1).Value = True   'set flag not former offender
            'if already Yes, leave alone
            Else
                optFormerNCR(0).Value = True
            End If
        End If
    End If
    
    '10/28/14: Service agreement
    If sgUsingServiceAgreement = "Y" Then
        If IsNull(rst!attServiceAgreement) Then
            optService(1).Value = True
        Else
            If rst!attServiceAgreement <> "Y" Then
                optService(1).Value = True
            Else
                optService(0).Value = True
            End If
        End If
    Else
        optService(1).Value = True
    End If
    If rst!attCarryCmml <> -1 Then
        optCarryCmml(rst!attCarryCmml).Value = True
    Else
        optCarryCmml(0).Value = True
    End If
    If rst!attSendTape <> -1 Then
        optSendTape(rst!attSendTape).Value = True
    Else
        optSendTape(0).Value = True
    End If
    If rst!attBarCode <> -1 Then
        optBarCode(rst!attBarCode).Value = True
    Else
        optBarCode(1).Value = True
    End If
    'If rst!attTimeType <> -1 Then
    '    optTimeType(rst!attTimeType).Value = True
    'Else
    '    optTimeType(0).Value = True
    'End If
    'optTimeType(0).Value = False
    'optTimeType(1).Value = False
    'optTimeType(2).Value = False
    If rst!attForbidSplitLive = "Y" Then
        ckcProhibitSplitCopy.Value = vbChecked
    Else
        ckcProhibitSplitCopy.Value = vbUnchecked
    End If
    If smCompensation = "Y" Then
        If rst!attComp <> -1 Then
            optComp(rst!attComp).Value = True
        Else
            optComp(0).Value = True
        End If
    Else
        optComp(0).Value = True
    End If
    If rst!attPostingType <> -1 Then
        optPost(rst!attPostingType).Value = True
    Else
        optPost(0).Value = True
    End If
    
'    If optPost(2).Value = True Then
'        frcExport.Visible = True
'    End If
    If rbcExportType(0).Value = True Then
        frcPosting.Visible = True
    Else
        frcPosting.Visible = False
    End If
    
    'If ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) And gUsingWeb Then
    If ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) And gUsingWeb Then
        '5/15/11: Remove PostType and LogType
        'frcPostType.Visible = True
        ''frcExport.Visible = True
        '5/15/11: Remove PostType and LogType
        'frcLogType.Visible = True
        frcSendLogEMail.Visible = True
        
        'D.S. 11/21/08 No longer used
        'lblWebPW.Visible = True
        'lblWebEmail.Visible = True
        'txtLogPassword.Visible = True
        'cmdGenPassword.Visible = True
        'txtEmailAddr.Visible = True
        
    Else
        frcPostType.Visible = False
        frcLogType.Visible = False
        frcSendLogEMail.Visible = False
        
        'D.S. 11/21/08 No longer used
        'lblWebPW.Visible = False
        'lblWebEmail.Visible = False
        'txtLogPassword.Visible = False
        'cmdGenPassword.Visible = False
        'txtEmailAddr.Visible = False
    End If
    
    If IsNull(rst!attVoiceTracked) Then
        optVoiceTracked(1).Value = True
    Else
        If rst!attVoiceTracked <> "Y" Then
            optVoiceTracked(1).Value = True
        Else
            optVoiceTracked(0).Value = True
        End If
    End If
    txtXDReceiverID.Text = rst!attXDReceiverId
    If Val(txtXDReceiverID.Text) = 0 Then
        txtXDReceiverID.Text = ""
    End If
    'Dan M 5457
    smPreviousXDReceiver = txtXDReceiverID.Text
    '7701
    smPreviousXDAudioDelivery = mRetrieveMultiListString(lbcAudioDelivery)
    'smPreviousXDAudioDelivery = lbcAudioDelivery.Text
    'Dan M 7375
   ' smPreviousXDAudioDelivery = rbcAudio(0).Value & rbcAudio(1).Value
    txtIDCReceiverID.Text = Trim$(rst!attIDCReceiverID)
    If Len(txtIDCReceiverID) = 0 Then
        smIDCReceiverID = ""
    Else
        smIDCReceiverID = txtIDCReceiverID.Text
        mIDCShowGroup True
    End If
    '6466  can never be null
    Select Case rst!attidcgrouptype
        Case "S"
            optIDCGroup(1).Value = True
        Case "L"
            optIDCGroup(2).Value = True
        Case Else
            optIDCGroup(0).Value = True
    End Select
    
    '4/3/19
    If rst!attExcludeFillSpot = "Y" Then
        ckcExcludeFillSpot.Value = vbChecked
    Else
        ckcExcludeFillSpot.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeQ = "Y" Then
        ckcExcludeCntrTypeQ.Value = vbChecked
    Else
        ckcExcludeCntrTypeQ.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeR = "Y" Then
        ckcExcludeCntrTypeR.Value = vbChecked
    Else
        ckcExcludeCntrTypeR.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeT = "Y" Then
        ckcExcludeCntrTypeT.Value = vbChecked
    Else
        ckcExcludeCntrTypeT.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeM = "Y" Then
        ckcExcludeCntrTypeM.Value = vbChecked
    Else
        ckcExcludeCntrTypeM.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeS = "Y" Then
        ckcExcludeCntrTypeS.Value = vbChecked
    Else
        ckcExcludeCntrTypeS.Value = vbUnchecked
    End If
    If rst!attExcludeCntrTypeV = "Y" Then
        ckcExcludeCntrTypeV.Value = vbChecked
    Else
        ckcExcludeCntrTypeV.Value = vbUnchecked
    End If

    
    'Dan M more 5457
    mXDigitalContact (False)
'    If optPost(2).Value = False Then
'        frcExport.Visible = False
'        frcLogType.Visible = False
'        frcPostType.Visible = False
'        lblWebPW.Visible = False
'        lblWebEmail.Visible = False
'        txtLogPassword.Visible = False
'        cmdGenPassword.Visible = False
'        txtEmailAddr.Visible = False
'    End If
    
    If rst!attSigned > 0 Then
        optSigned(rst!attSigned).Value = True
    Else
        optSigned(0).Value = True
    End If
    If optSigned(1).Value = True Then
        If (DateValue(gAdjYear(rst!attSignDate)) = DateValue("1/1/1970")) Then    ' Or (rst!attSignDate = "1/1/70") Then
            txtRetDate.Text = ""
        Else
            txtRetDate.Text = Format$(rst!attSignDate, sgShowDateForm)
        End If
    Else
        txtRetDate.Text = ""
    End If
    txtLog.Text = rst!attGenLog
    txtCP.Text = rst!attGenCP
    txtACName.Text = rst!attACName
    txtACPhone.Text = rst!attACPhone
    If (DateValue(gAdjYear(rst!attDropDate)) = DateValue("12/31/2069")) Then  ' Or (rst!attDropDate = "12/31/69") Then
        txtDropDate.Text = ""
        imChgDropDate = False
    Else
        txtDropDate.Text = Format$(rst!attDropDate, sgShowDateForm)
        imChgDropDate = False
    End If
    smSvDropDate = Format$(rst!attDropDate, "m/d/yyyy")
    
    '4/15/09: Set Off Air date to Drop date if Drop date is sooner then the off air date and remove the drop date.
    'This is required as the Drop date have been made invisble
    If (txtDropDate.Text <> "") And (Not bmShowDates) Then
        If txtOffAirDate.Text = "" Then
            txtOffAirDate.Text = txtDropDate.Text
            smSvOffAirDate = txtOffAirDate.Text
            txtDropDate.Text = ""
        Else
            If DateValue(gAdjYear(txtDropDate.Text)) < DateValue(gAdjYear(txtOffAirDate.Text)) Then
                txtOffAirDate.Text = txtDropDate.Text
                smSvOffAirDate = txtOffAirDate.Text
                txtDropDate.Text = ""
            End If
        End If
    End If
    '7701 use my new string variables in place of rst
    'Start check to see if any posting has been done for this agreement
    smLastPostedDate = gGetLastPostedDate(rst!attCode, rst!attExportType, rst!attExportToWeb, rst!attExportToUnivision, slattExportToMarketron, slattExportToCBS, slattExportToClearCh)
    If DateValue(smSvOffAirDate) >= DateValue(smSvDropDate) Then
        smTrueOffAirDate = smSvOffAirDate
    Else
        smTrueOffAirDate = smSvDropDate
    End If
    
    If IsNull(rst!attMktronActiveDate) Then
        '7701
        If slattExportToMarketron = "Y" Then
            smMktronActiveDate = "8/19/2010"
        Else
            smMktronActiveDate = "1/1/1970"
        End If
    Else
        smMktronActiveDate = Format$(rst!attMktronActiveDate, sgShowDateForm)
    End If
    
    txtComments.Text = Trim$(rst!attComments)
    txtOther.Text = rst!attGenOther
    If IsNull(rst!attStartTime) Then
        sgCDStartTime = ""
    Else
        sgCDStartTime = Format$(rst!attStartTime, "hh:mmA/P")
    End If
    If rst!attRadarClearType = "C" Then
        optRadarClearType(0).Value = True
    ElseIf rst!attRadarClearType = "P" Then
        optRadarClearType(1).Value = True
    ElseIf rst!attRadarClearType = "E" Then
        optRadarClearType(3).Value = True
    Else
        optRadarClearType(2).Value = True
    End If
    sgVehProgStartTime = "12am"
    sgVehProgEndTime = "12am"
    If IsNull(rst!attVehProgStartTime) Then
        lacPrgTimes.Caption = ""
    Else
        sgVehProgStartTime = gCompactTime(Format$(rst!attVehProgStartTime, "hh:mm:ssA/P"))
        lacPrgTimes.Caption = "Program Times: " & sgVehProgStartTime
        If Not IsNull(rst!attVehProgEndTime) Then
            sgVehProgEndTime = gCompactTime(Format$(rst!attVehProgEndTime, "hh:mm:ssA/P"))
            lacPrgTimes.Caption = lacPrgTimes.Caption & "-" & sgVehProgEndTime
        End If
    End If
    edcNoAirPlays.Text = rst!attNoAirPlays
        
    imFieldChgd = False
    imDateChgd = False
    If smLastPostedDate <> "" Then
        If DateValue(smLastPostedDate) >= DateValue(smSvOnAirDate) And DateValue(smLastPostedDate) <= DateValue(smTrueOffAirDate) Then
            imOkToChange = False
        Else
            imOkToChange = True
        End If
        If imOkToChange Then
            Call mEnableControls
        Else
            Call mDisableControls
        End If
    Else
        imOkToChange = True
        Call mEnableControls
    End If
    
    If sgUstWin(2) = "I" Then
        cmdSave.Enabled = True
        cmdNew.Enabled = False
    End If
    
    ReDim tmETAvailInfo(0 To 0) As ETAVAILINFO
    
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 3 'populate

    ReDim tgAirPlaySpec(0 To 0) As AIRPLAYSPEC
    ReDim tgBreakoutSpec(0 To 0) As BREAKOUTSPEC
    ReDim tgDPSelection(0 To 0) As DPSELECTION
    ReDim tmPetInfo(0 To 0) As PETINFO
    '8/12/16: Separtated Pledge and Estimate so that est can be saved in the past
    bmPledgeDataChgd = False
    bmETDataChgd = False
    bmSaving = False
    Exit Sub

ErrHand:
    mMousePointer vbDefault
    gHandleError "", "Agreeement-BindControls"
    imInChg = False
End Sub
Private Sub mXDigitalContact(blCompare As Boolean)
'ttp 5457
    Dim slCurrentXDContact As XDIGITALSTATIONINFO
    
    If imShttCode > 0 Then
        'first time through, store values
        If Not blCompare Then
            'is site set to allow xdigital?
            If bmIsXDSiteStation Then
                gXDStationContact imShttCode, smPreviousXDContact
            ' this should be blank anyway, as will never get set.  But just to make sure.
            ' this stops the processing the 2nd time through if site not set to allow xdigital
            Else
                smPreviousXDContact.sEmail = ""
            End If
        'second time, if have previous values, compare and save as needed
        ElseIf Len(Trim$(smPreviousXDContact.sEmail)) > 0 Then
            gXDStationContact imShttCode, slCurrentXDContact
            With slCurrentXDContact
                If .sContactName <> smPreviousXDContact.sContactName Or .sEmail <> smPreviousXDContact.sEmail Or .sPhone <> smPreviousXDContact.sPhone Then
                    'need to set previous here only in case they make further changes after save but on same contact screen
                    smPreviousXDContact.sContactName = .sContactName
                    smPreviousXDContact.sEmail = .sEmail
                    smPreviousXDContact.sPhone = .sPhone
                    SQLQuery = "UPDATE shtt set shttSentToXDSStatus = 'M' WHERE shttCode = " & imShttCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mXDigitalContact"
                        Exit Sub
                    End If
                End If
            End With
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "mXDigitalContact"
End Sub
Private Sub mDisableControls()
    'txtOnAirDate.Enabled = False
    txtStartDate.SetEnabled False
    '1/9/13:Allow any End Date
    'txtEndDate.SetEnabled False
    'txtLdMult.Enabled = False
    'txtNoCDs.Enabled = False
    txtDays.Enabled = False
    'txtACName.Enabled = False
    'txtACPhone.Enabled = False
    'txtComments.Enabled = False
    'optSuppressNotices(0).Enabled = False
    'optSuppressNotices(1).Enabled = False
    'optPrintCP(0).Enabled = False
    'optPrintCP(1).Enabled = False
    cmdAdjustDates.Enabled = False
    txtLog.Enabled = False
    txtCP.Enabled = False
    'txtOther.Enabled = False
    rbcMulticast(0).Enabled = False
    rbcMulticast(1).Enabled = False
    'txtLabelID.Enabled = False
    'txtShipInfo.Enabled = False
    'optTimeType(0).Enabled = False
    'optTimeType(1).Enabled = False
    cmcPledgeBy(0).Enabled = False
    cmcPledgeBy(1).Enabled = False
    'optCarryCmml(0).Enabled = False
    'optCarryCmml(1).Enabled = False
    'optSendTape(0).Enabled = False
    'optSendTape(1).Enabled = False
    rbcExportType(0).Enabled = False
    rbcExportType(1).Enabled = False
    'rbcExportType(3).Enabled = False
    rbcLogType(0).Enabled = False
    rbcLogType(1).Enabled = False
    rbcPostType(0).Enabled = False
    rbcPostType(1).Enabled = False
    rbcPostType(2).Enabled = False
    'rbcSendLogEMail(0).Enabled = False
    'rbcSendLogEMail(1).Enabled = False
    'txtLogPassword.Enabled = False
    'txtEmailAddr.Enabled = False
    'cmdGenPassword.Enabled = False
    optSigned(0).Enabled = False
    optSigned(1).Enabled = False
    txtRetDate.Enabled = False
    'txtLogPassword.Enabled = False
    optComp(1).Enabled = False
    optComp(2).Enabled = False
    'optBarCode(0).Enabled = False
    'optBarCode(1).Enabled = False
    cmdErase.Enabled = False
    cmdClearAvails.Enabled = False
    imcTrash.Enabled = False
    '7701
    mEnableDeliveryOptions True
    mEnableDeliveryOptions False

End Sub

Private Sub mEnableControls()
    'txtOnAirDate.Enabled = True
    If bmShowDates Then
        txtStartDate.SetEnabled True
        txtEndDate.SetEnabled True
        txtDropDate.SetEnabled True
    End If
    txtLdMult.Enabled = True
    txtNoCDs.Enabled = True
    txtDays.Enabled = True
    'txtACName.Enabled = True
    'txtACPhone.Enabled = True
    'txtComments.Enabled = True
    'optSuppressNotices(0).Enabled = True
    'optSuppressNotices(1).Enabled = True
    'optPrintCP(0).Enabled = True
    'optPrintCP(1).Enabled = True
    cmdAdjustDates.Enabled = True
    txtLog.Enabled = True
    txtCP.Enabled = True
    'txtOther.Enabled = True
    rbcMulticast(0).Enabled = True
    rbcMulticast(1).Enabled = True
    'txtLabelID.Enabled = True
    'txtShipInfo.Enabled = True
    'optTimeType(0).Enabled = True
    'optTimeType(1).Enabled = True
    cmcPledgeBy(0).Enabled = True
    cmcPledgeBy(1).Enabled = True
    optCarryCmml(0).Enabled = True
    optCarryCmml(1).Enabled = True
    optSendTape(0).Enabled = True
    optSendTape(1).Enabled = True
    rbcExportType(0).Enabled = True
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        rbcExportType(1).Enabled = False
    Else
        rbcExportType(1).Enabled = True
    End If
    'rbcExportType(3).Enabled = True
    rbcLogType(0).Enabled = True
    rbcLogType(1).Enabled = True
    rbcPostType(0).Enabled = True
    rbcPostType(1).Enabled = True
    rbcPostType(2).Enabled = True
    'rbcSendLogEMail(0).Enabled = True
    'rbcSendLogEMail(1).Enabled = True
    'txtLogPassword.Enabled = True
    'txtEmailAddr.Enabled = True
    'cmdGenPassword.Enabled = True
    optSigned(0).Enabled = True
    optSigned(1).Enabled = True
    txtRetDate.Enabled = True
    'txtLogPassword.Enabled = True
    optComp(1).Enabled = True
    optComp(2).Enabled = True
    'optBarCode(0).Enabled = True
    'optBarCode(1).Enabled = True
    cmdErase.Enabled = True
    cmdClearAvails.Enabled = True
    imcTrash.Enabled = True
    '7701
    mEnableDeliveryOptions True
    mEnableDeliveryOptions False
End Sub


Private Sub cbcAirPlayNo_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cbcContractPDF_DblClick()
    Dim slStr As String
    If cbcContractPDF.ListIndex >= 0 Then
        slStr = cbcContractPDF.GetName(cbcContractPDF.ListIndex)
        ShellExecute 0&, vbNullString, sgContractPDFPath & smContractPDFSubFolder & slStr, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub cbcContractPDF_GotFocus()
    imIgnoreTabs = False
    cbcContractPDF.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcContractPDF_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcContractPDF_ReSetLoc()
    'Display list abobe edit box
    cbcContractPDF.Top = lacContractPDF.Top - 30 + txtStartDate.Height - cbcContractPDF.Height
End Sub

Private Sub cbcETDay_OnChange()
    If grdET.TextMatrix(lmETEnableRow, lmETEnableCol) <> cbcETDay.Text Then
        grdET.TextMatrix(lmETEnableRow, lmETEnableCol) = cbcETDay.Text
        imFieldChgd = True
        '8/12/16: Separtated Pledge and Estimate so that est can be saved in the past
        'bmPledgeDataChgd = True
        bmETDataChgd = True
    End If
End Sub

Private Sub cbcMarketRep_GotFocus()
    imIgnoreTabs = False
    cbcMarketRep.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcMarketRep_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcSeason_GotFocus()
    tmcDelay.Enabled = False
    cbcSeason.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcSeason_OnChange()
    imFieldChgd = True
    tmcDelay.Enabled = False
    Sleep 10
    tmcDelay.Enabled = True
End Sub

Private Sub cbcServiceRep_GotFocus()
    imIgnoreTabs = False
    cbcServiceRep.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcServiceRep_OnChange()
    imFieldChgd = True
End Sub

Private Sub cboAffAE_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    mMousePointer vbHourglass
    sName = LTrim$(cboAffAE.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboAffAE.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboAffAE.ListIndex = lRow
        cboAffAE.SelStart = iLen
        cboAffAE.SelLength = Len(cboAffAE.Text)
    End If
    imFieldChgd = True
    mMousePointer vbDefault
    imInChg = False
    Exit Sub
End Sub

Private Sub cboAffAE_Click()
    cboAffAE_Change
End Sub

Private Sub cboAffAE_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboAffAE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboAffAE.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub






Private Sub cboPSSort_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim tmp_rst As ADODB.Recordset
    Dim llTemp As Long
    Dim slStr As String
    Dim slName As String
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim ilVff As Integer
    
        
    If imInChg Then
        Exit Sub
    End If
       
    'Clear the radio buttons (dayparts live, cd/tape and avails) each time a new affiliate
    'is brought up.
    'optTimeType(0).Value = False
    'optTimeType(1).Value = False

    
    
    imInChg = True
    mMousePointer vbHourglass
    sName = LTrim$(cboPSSort.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then       'Sorting by stations/markets
        lRow = SendMessageByString(cboPSSort.hwnd, CB_FINDSTRING, -1, sName)
    Else
        lRow = SendMessageByString(cboPSSort.hwnd, CB_FINDSTRING, -1, sName)
    End If
    If lRow >= 0 Then
        'If optPSSort(2).Value = False Then       'Sorting by stations/markets
        '    cboPSSort.Bookmark = lRow ' + 1
        'Else
        '    cboPSSort.Bookmark = lRow
        'End If
        'cboPSSort.Text = cboPSSort.Columns(0).Text
        cboPSSort.ListIndex = lRow
        cboPSSort.SelStart = iLen
        cboPSSort.SelLength = Len(cboPSSort.Text)
        TabStrip1.Tabs(3).Caption = "&Pledge"
        ClearControls
        cboSSSort.Clear
        DoEvents
        
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then       'Sorting by stations/markets
            imVefCode = 0
            mSetEventTitles
            imShttCode = CInt(cboPSSort.ItemData(lRow))
            'TabStrip1.Tabs(3).Caption = "Pledge &Information "
            'For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            '    If tgStationInfo(iLoop).iCode = imShttCode Then
            '        TabStrip1.Tabs(3).Caption = "Pledge &Information " & Trim$(tgStationInfo(iLoop).sZone)
            '        Exit For
            '    End If
            'Next iLoop
            If optExAll(1).Value = False Then   'Affiliates, then populate cboSSSort, other it is not required
                optExAll_Click 0
            Else
                optExAll_Click 1
            End If
        Else
            imShttCode = 0
            imVefCode = CInt(cboPSSort.ItemData(lRow))
            mSetEventTitles
            If optExAll(1).Value = False Then
                optExAll_Click 0
            Else
                optExAll_Click 1
            End If
        End If
        
        SQLQuery = "SELECT * FROM att"
        SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF = False Then
            BindControls
            IsAgmntDirty = True
        Else
            IsAgmntDirty = False
            ClearControls
            'Replaced by Market Rep, attMktRepUstCode
            'mSetInitAffAE
        End If
                
    
    Else
        imShttCode = 0
        imVefCode = 0
        mSetEventTitles
        If optExAll(1).Value = False Then
            optExAll_Click 0
        Else
            optExAll_Click 1
        End If
    End If
    mGetPledgeBy
    mGetEventStartDate
    mPopSeason
    mShowTabs
    ilRet = mPopMulticast()
    If sgUstWin(2) = "I" Then
        cmdRemap.Enabled = True
    End If
    mPopContractPDF True
    smPassword = gGetStationPW(imShttCode)
    txtPassword.Text = Trim$(smPassword)
    mMousePointer vbDefault
    
    imInChg = False
End Sub

Private Sub cboPSSort_Click()
    cboPSSort_Change
    Exit Sub
End Sub

Private Sub cboPSSort_GotFocus()
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cboPSSort_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboPSSort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboPSSort.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboSSSort_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim ilRet As Integer
    Dim llTemp As Long
    Dim slStr As String
    Dim slName As String
    Dim ilUpper As Integer
    Dim ilVff As Integer
    Dim tmp_rst As ADODB.Recordset
    
    If imInChg Then
        Exit Sub
    End If
    
    'Clear the radio buttons (dayparts live, cd/tape and avails) each time a new affiliate
    'is brought up.
    'optTimeType(0).Value = False
    'optTimeType(1).Value = False

    imInChg = True
    mMousePointer vbHourglass
    sName = LTrim$(cboSSSort.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    If (optPSSort(2).Value = True) Or (optPSSort(3).Value = True) Then       'Sorting by stations/markets
        lRow = SendMessageByString(cboSSSort.hwnd, CB_FINDSTRING, -1, sName)
    Else
        lRow = SendMessageByString(cboSSSort.hwnd, CB_FINDSTRING, -1, sName)
    End If
    If lRow >= 0 Then
        On Error GoTo ErrHand
        cboSSSort.ListIndex = lRow
        cboSSSort.SelStart = iLen
        cboSSSort.SelLength = Len(cboSSSort.Text)
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then                       'Sorting by Stations
            If optExAll(0).Value Then
                lmAttCode = CLng(cboSSSort.ItemData(cboSSSort.ListIndex))
                imVefCode = 0
            Else
                imVefCode = CInt(cboSSSort.ItemData(cboSSSort.ListIndex))
            End If
            mSetEventTitles
        Else
            If optExAll(0).Value Then
                lmAttCode = CLng(cboSSSort.ItemData(cboSSSort.ListIndex))
                imShttCode = 0
            Else
                imShttCode = CInt(cboSSSort.ItemData(cboSSSort.ListIndex))
            End If
        End If
        imIgnoreTimeTypeChg = True
        SQLQuery = "SELECT * FROM att"
        SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF = False Then
            '12/19/13- Added Clear so that unchecked items are unchecked
            ClearControls
            BindControls
            IsAgmntDirty = True
        Else
            IsAgmntDirty = False
            ClearControls
            'Replaced by Market Rep, attMktRepUstCode
            'mSetInitAffAE
        End If
        imDateChgd = False
        smShttACName = ""
        smShttACPhone = ""
        
        ' SQLQuery = "SELECT shttACName, shttACPhone FROM shtt WHERE (shttCode = '" & imShttCode & "')"
        SQLQuery = "SELECT arttFirstName, arttLastName, arttPhone FROM artt WHERE (arttShttCode = '" & imShttCode & "' And arttAffContact = '1')"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            smShttACName = Trim$(rst!arttFirstName) & " " & Trim$(rst!arttLastName)
            smShttACPhone = Trim$(rst!arttPhone)
        End If
        TabStrip1.Tabs(3).Caption = "&Pledge "
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).iCode = imShttCode Then
                smStationZone = Trim$(tgStationInfo(iLoop).sZone)
                'D.S. 08/11/16 Added last posted date per Dick
                If smLastPostedDate <> "1/1/1970" Then
                    TabStrip1.Tabs(3).Caption = "&Pledge " & Trim$(tgStationInfo(iLoop).sZone) & ",   LPW " & smLastPostedDate
                Else
                    TabStrip1.Tabs(3).Caption = "&Pledge " & Trim$(tgStationInfo(iLoop).sZone)
                End If
                Exit For
            End If
        Next iLoop
        If Trim$(txtACName.Text) = "" Then
            txtACName.Text = smShttACName 'rst(0).Value
            txtACPhone.Text = smShttACPhone 'rst(0).Value
            imFieldChgd = False
        End If
        
        'ReDim tgDat(0 To 0) As DAT
        imDatLoaded = False
        ''grdDayparts.RemoveAll
        'gGrid_Clear grdPledge, True
        mClearPledgeGrid
        mGetPledgeBy
        mGetEventStartDate
        mPopSeason
        'If (imTabIndex = TABPLEDGE) And (Not imDatLoaded) Then
        If (Not imDatLoaded) And (IsAgmntDirty = True) Then
            mLoadPledge True, -1
            'set in Bind or Clear
            'If UBound(tgDat) > LBound(tgDat) Then
            '    If tgDat(0).iDACode = 0 Then
            '        optTimeType(0).Value = True
            '    ElseIf tgDat(0).iDACode = 1 Then
            '        optTimeType(1).Value = True
            '    ElseIf tgDat(0).iDACode = 2 Then
            '        optTimeType(2).Value = True
            '    End If
            'Else
            '    optTimeType(0).Value = False
            '    optTimeType(1).Value = False
            '    optTimeType(2).Value = False
            'End If
        End If
        imIgnoreTimeTypeChg = False
        imFieldChgd = False
    End If
    If optExAll(1).Value Then
        If TabStrip1.SelectedItem.Index = TABPLEDGE Then
            ''TabStrip1.SelectedItem.Index = TABMAIN
            imTabIndex = -1
            ''TabStrip1.SetFocus
            'gSendKeys "%M", True
            TabStrip1.Tabs(TABMAIN).Selected = True
        End If
    End If
    mShowTabs
    ilRet = mPopMulticast()
    mPopContractPDF True
    smPassword = gGetStationPW(imShttCode)
    txtPassword.Text = Trim$(smPassword)
    mMousePointer vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-cboSSSort"
    imInChg = False
End Sub

Private Sub cboSSSort_Click()
    cboSSSort_Change
End Sub

Private Sub cboSSSort_GotFocus()
    'cboSSSort.DroppedDown = False
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
    If cboSSSort.ListIndex < 0 Then
        If cboSSSort.ListCount > 0 Then
            cboSSSort.ListIndex = 0
            cboSSSort.SelStart = 0
            cboSSSort.SelLength = Len(cboSSSort.Text)
        End If
    End If
End Sub

Private Sub cboSSSort_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboSSSort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSSSort.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboSSSort_LostFocus()
    Dim sCallLetters As String
    Dim sVehicle As String
    Dim iLen As Integer
    Dim iFound As Integer
    Dim iLoop As Integer
    
    mMousePointer vbHourglass
    'If optPSSort(2).Value = True Then       'Sorting by stations/markets
    '    sCallLetters = Trim$(UCase$(cboSSSort.Text))
    '    iLen = Len(sCallLetters)
    '    If (iLen > 0) Then
    '        iFound = False
    '        cboSSSort.MoveFirst
    '        For iLoop = 0 To cboSSSort.Rows - 1
    '            If optSSSort(0).Value = True Then
    '                If StrComp(sCallLetters, Trim$(cboSSSort.Columns(0).Text), 1) = 0 Then
    '                    iFound = True
    '                    If imShttCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                        cboSSSort_Click
    '                    End If
    '                    Exit For
    '                End If
     '           Else
    '                If StrComp(sCallLetters, Trim$(cboSSSort.Columns(0).Text), 1) = 0 Then
    '                    iFound = True
    '                    If imShttCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                        cboSSSort_Click
    '                    End If
    '                    Exit For
    '                End If
    '            End If
    '            cboSSSort.MoveNext
    '        Next iLoop
    '        If Not iFound Then
    '            cboSSSort.MoveFirst
    '            For iLoop = 0 To cboSSSort.Rows - 1
    '                If optSSSort(0).Value = True Then
    '                    If StrComp(sCallLetters, Left$(cboSSSort.Columns(0).Text, iLen), 1) = 0 Then
    '                        If imShttCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                            cboSSSort_Click
    '                        End If
    '                        Exit For
    '                    End If
    '                Else
    '                    If StrComp(sCallLetters, Left$(cboSSSort.Columns(0).Text, iLen), 1) = 0 Then
    '                        If imShttCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                            cboSSSort_Click
    '                        End If
    '                        Exit For
    '                    End If
    '                End If
    '                cboSSSort.MoveNext
    '            Next iLoop
    '        End If
    '    End If
    'Else
    '    sVehicle = Trim$(UCase$(cboSSSort.Text))
    '    iLen = Len(sVehicle)
    '    If (iLen > 0) Then
    '        iFound = False
    '        cboSSSort.MoveFirst
    '        For iLoop = 0 To cboSSSort.Rows - 1
    '            If StrComp(sVehicle, Trim$(cboSSSort.Columns(0).Text), 1) = 0 Then
    '                iFound = True
    '                If imVefCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                    cboSSSort_Click
    '                End If
    '                Exit For
    '            End If
    '            cboSSSort.MoveNext
    '        Next iLoop
    '        If Not iFound Then
    '            cboSSSort.MoveFirst
    '            For iLoop = 0 To cboSSSort.Rows - 1
    '                If StrComp(sVehicle, Left$(cboSSSort.Columns(0).Text, iLen), 1) = 0 Then
     '                   If imVefCode <> CInt(cboSSSort.Columns(1).Text) Then
    '                        cboSSSort_Click
    '                    End If
    '                    Exit For
    '                End If
    '                cboSSSort.MoveNext
    '            Next iLoop
    '        End If
    '    End If
    'End If
    mMousePointer vbDefault
End Sub

Private Sub chkActive_Click()
    cboSSSort.Clear
    optExAll_Click 1
End Sub

Private Sub chkMonthlyWebPost_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub ckcDay_Click()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRow As Long
    
    
    If ckcDay.Value = vbChecked Then
        grdPledge.Text = 4
    Else
        grdPledge.Text = ""
    End If
    imFieldChgd = True
    bmPledgeDataChgd = True
    slStr = grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)
    llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        ilIndex = lbcStatus.ItemData(llRow)
        If (tmStatusTypes(ilIndex).iPledged = 0) Then
            If grdPledge.Col <= SUNFDINDEX Then
                grdPledge.TextMatrix(grdPledge.Row, grdPledge.Col + MONPDINDEX - MONFDINDEX) = grdPledge.Text
            Else
                grdPledge.TextMatrix(grdPledge.Row, grdPledge.Col) = grdPledge.Text
            End If
        End If
    End If
End Sub

Private Sub ckcExcludeCntrTypeM_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExcludeCntrTypeQ_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExcludeCntrTypeR_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExcludeCntrTypeS_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExcludeCntrTypeT_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExcludeFillSpot_Click()
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imFieldChgd = True
End Sub

Private Sub ckcExportTo_Click(Index As Integer)
    If ckcExportTo(Index).Value Then
        Select Case Index
            Case 0
                ckcExportTo(2).Value = vbUnchecked      'Univision
                ckcExportTo(3).Value = vbUnchecked      'Marketron
            Case 1
                '6524
'                ckcExportTo(0).Value = vbChecked        'Web
'                ckcExportTo(2).Value = vbUnchecked      'Univision
'                ckcExportTo(3).Value = vbUnchecked      'Marketron
'                '6592
'                ckcExportTo(4).Value = vbUnchecked      'cbs
                ckcExportTo(4).Value = vbUnchecked      'cbs
                ckcExportTo(0).Value = vbChecked        'Web
            Case 2
                'ckcExportTo(0).Value = vbUnchecked      'Web
                ckcExportTo(0).Value = vbChecked
                ckcExportTo(1).Value = vbUnchecked      'Cumulus
                ckcExportTo(3).Value = vbUnchecked      'Marketron
                '6592
                ckcExportTo(4).Value = vbUnchecked      'cbs
            Case 3
                'ckcExportTo(0).Value = vbUnchecked      'Web
                ckcExportTo(0).Value = vbChecked
                ckcExportTo(1).Value = vbUnchecked      'Cumulus
                ckcExportTo(2).Value = vbUnchecked      'Univision
                '6592
                ckcExportTo(4).Value = vbUnchecked      'cbs
            '6592
            Case 4
                ckcExportTo(1).Value = vbUnchecked      'Cumulus
                ckcExportTo(0).Value = vbChecked        'Web
        End Select
        If ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) And gUsingWeb Then
            '5/15/11: Remove PostType and LogType
            'frcLogType.Visible = True
            'frcPostType.Visible = True
            frcSendLogEMail.Visible = True
        Else
            frcLogType.Visible = False
            frcPostType.Visible = False
            frcSendLogEMail.Visible = False
            'lblWebPW.Visible = False
            'lblWebEmail.Visible = False
            'txtLogPassword.Visible = False
            'cmdGenPassword.Visible = False
            'txtEmailAddr.Visible = False
        End If
    Else
        If ckcExportTo(0).Value = vbUnchecked Then  'Web
            ckcExportTo(1).Value = vbUnchecked      'Cumulus
            '6592
            ckcExportTo(4).Value = vbUnchecked      'cbs
        End If
    End If
End Sub

Private Sub ckcProhibitSplitCopy_Click()
    imFieldChgd = True
End Sub

Private Sub ckcProhibitSplitCopy_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub ckcSendNotCarried_Click()
    '9452
        bmSendNotCarriedChange = True
End Sub

Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 1
    sgGetPath = sgContractPDFPath
    frmGetPath.Show vbModal
    If igGetPath = 0 Then
        smContractPDFSubFolder = Mid(sgGetPath, Len(sgContractPDFPath) + 1)
        If smContractPDFSubFolder <> "" Then
            smContractPDFSubFolder = gSetPathEndSlash(smContractPDFSubFolder, False)
        End If
        mPopContractPDF False
    End If
    
    ChDir slCurDir
    
    Exit Sub
End Sub

Private Sub cmcBrowse_GotFocus()
    imIgnoreTabs = False
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmcDropDown_Click()
    lbcStatus.Visible = Not lbcStatus.Visible
End Sub

Private Sub cmcPledgeBy_Click(Index As Integer)
    mTimeType Index
End Sub

Private Sub cmcPledgeBy_GotFocus(Index As Integer)
    imIgnoreTabs = False
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdAdjustDates_Click()
    mAdjustDates
End Sub

Private Sub cmdCancel_Click()
    Unload frmAgmnt
End Sub


Private Sub cmdCancel_GotFocus()
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

'*****************************************************************************************
'   mLoadPledge
'
'   Purpose: Populate the avails grid on the Agreement's Pledge screen.  Loop on tgDat array
'            adding avails into the grid.
'
'   Parameter: ilFromDat - If true get the avail info from the Dat table and pump it into
'                          the tgDat array.  If False call gGetAvails which populates the
'                          tgDat array.
'
'******************************************************************************************
Private Sub mLoadPledge(ilFromDat As Integer, ilTimeTypeIndex As Integer)
    Dim iUpper As Integer
    Dim iLoop As Integer
    Dim sStatus As String
    Dim iIndex As Integer
    Dim sSDate As String
    Dim iDay As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slTemp As String
    Dim llPdTime As Long
    Dim llDatPdTime As Long
    Dim llLibStartTime As Long
    Dim ilAdjustCDStartTime As Integer
    Dim llTimeOffset As Long
    Dim ilShowActTime As Integer
    Dim llVpf As Long
    'Dim ilVff As Integer
    Dim VehCombo_rst As ADODB.Recordset
    Dim blPledgeDataChgd As Boolean
    ReDim iFdDay(0 To 6) As Integer
    ReDim iPdDay(0 To 6) As Integer
    
    
    On Error GoTo ErrHand

    ilShowActTime = False
    blPledgeDataChgd = bmPledgeDataChgd
'    If (optTimeType(1).Value) And (Not imIgnoreTimeTypeChg) And (optExAll(1).Value <> True) Then
    'If (optTimeType(1).Value) And (Not imIgnoreTimeTypeChg) And (optExAll(1).Value <> True) And (UBound(tgDat) > LBound(tgDat)) Then
'    If (Not imIgnoreTimeTypeChg) And (UBound(tgDat) > LBound(tgDat)) Then
    If ((Not imIgnoreTimeTypeChg) And (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "")) Then
        'If optTimeType(0).Value = False Then
        '    ilRet = gMsgBox("Warning: You Are About To Clear Existing Daypart or Avails Information and Load New Avails!", vbOKCancel)
        '    If ilRet = vbCancel Then
        '        frmAgmnt!optTimeType(0).Value = False
        '        frmAgmnt!optTimeType(1).Value = False
        '        Exit Sub
        '    End If
        'End If
    End If

    If sgUstWin(2) = "I" Then
        cmdRemap.Enabled = True
    End If
'    If imShttCode <= 0 Then
'        gMsgBox "Station must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
'        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
'            If cboPSSort.Enabled Then
'                cboPSSort.SetFocus
'            End If
'        Else
'            If cboSSSort.Enabled Then
'                cboSSSort.SetFocus
'            End If
'        End If
'        Exit Sub
'    End If
'    If imVefCode <= 0 Then
'        gMsgBox "Vehicle must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
'        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
'            If cboPSSort.Enabled Then
'                cboPSSort.SetFocus
'            End If
'        Else
'            If cboSSSort.Enabled Then
'                cboSSSort.SetFocus
'            End If
'        End If
'        Exit Sub
'    End If
'
'    If txtOnAirDate.Text = "" Then
'        'If ((optTimeType(0).Value) Or (optTimeType(1).Value)) And (Not imIgnoreTimeTypeChg) Then
'        If (Not imIgnoreTimeTypeChg) Then
'            Beep
'            gMsgBox "On Air Date must be specified." & Chr$(13) & Chr$(10) & "Please enter.", vbOKOnly
'        '    optTimeType(0).Value = False
'        '    optTimeType(1).Value = False
'
'        End If
'        Exit Sub
'    Else
'        If gIsDate(txtOnAirDate.Text) = False Then
'            If Not imIgnoreTimeTypeChg Then
'                Beep
'                gMsgBox "On Air Date must be specified in the form mm/dd/yy.", vbOKOnly
'            End If
'            Exit Sub
'        End If
'    End If
    If Not mLoadPledgeOk() Then
        bmPledgeDataChgd = blPledgeDataChgd
        Exit Sub
    End If
    cmdRemap.Enabled = False
    cmdFastAdd.Enabled = False
    cmdFastEnd.Enabled = False
    igReload = True
    
    imVefCombo = 0
    SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
    Set VehCombo_rst = gSQLSelectCall(SQLQuery)
    If Not VehCombo_rst.EOF Then
        imVefCombo = VehCombo_rst!vefCombineVefCode
    End If
    
    smDefaultEmbeddedOrROS = "R"
    llVpf = gBinarySearchVpf(CLng(imVefCode))
    If llVpf <> -1 Then
        smDefaultEmbeddedOrROS = tgVpfOptions(llVpf).sEmbeddedOrROS
    End If
    If Trim$(smDefaultEmbeddedOrROS) = "" Then
        smDefaultEmbeddedOrROS = "R"
    End If
    
    'ilVff = gBinarySearchVff(imVefCode)
    'If ilVff <> -1 Then
    '    smPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
    '    If smPledgeByEvent = "" Then
    '        smPledgeByEvent = "N"
    '    End If
    'Else
    '    smPledgeByEvent = "N"
    'End If
    
    If imDatLoaded = False Then
        mMousePointer vbHourglass
        If smPledgeByEvent <> "Y" Then
            mClearEventGrid
            ReDim tgDat(0 To 0) As DAT
            tgDat(0).sFdSTime = ""
            iUpper = 0
            If ilFromDat Then
                'We already have the dat table loaded. No need to go and get the avails
                SQLQuery = "SELECT * "
                SQLQuery = SQLQuery + " FROM dat"
                SQLQuery = SQLQuery + " WHERE (datAtfCode = " & lmAttCode & ""
                SQLQuery = SQLQuery + " AND datShfCode = " & imShttCode & ""
                SQLQuery = SQLQuery + " AND datVefCode = " & imVefCode & ")"
                SQLQuery = SQLQuery & " ORDER BY datFdStTime"
                'If optTimeType(0).Value = True Then
                '    SQLQuery = SQLQuery + " AND dat.datDACode = 1)"
                'Else
                '    SQLQuery = SQLQuery + " AND dat.datDACode = 1)"
                'End If
        
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    While Not rst.EOF
                        tgDat(iUpper).iStatus = 1
                        tgDat(iUpper).lCode = rst!datCode    '(0).Value
                        tgDat(iUpper).lAtfCode = rst!datAtfCode  '(1).Value
                        tgDat(iUpper).iShfCode = rst!datShfCode  '(2).Value
                        tgDat(iUpper).iVefCode = rst!datVefCode  '(3).Value
                        'tgDat(iUpper).iDACode = rst!datDACode    '(4).Value
                        tgDat(iUpper).iFdDay(0) = rst!datFdMon   '(5).Value
                        tgDat(iUpper).iFdDay(1) = rst!datFdTue   '(6).Value
                        tgDat(iUpper).iFdDay(2) = rst!datFdWed   '(7).Value
                        tgDat(iUpper).iFdDay(3) = rst!datFdThu   '(8).Value
                        tgDat(iUpper).iFdDay(4) = rst!datFdFri   '(9).Value
                        tgDat(iUpper).iFdDay(5) = rst!datFdSat   '(10).Value
                        tgDat(iUpper).iFdDay(6) = rst!datFdSun   '(11).Value
                        If Second(rst!datFdStTime) = 0 Then
                            tgDat(iUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
                        Else
                            tgDat(iUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
                        End If
                        If Second(rst!datFdEdTime) = 0 Then
                            tgDat(iUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWOSecForm)
                        Else
                            tgDat(iUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWSecForm)
                        End If
                        tgDat(iUpper).iFdStatus = rst!datFdStatus    '(14).Value
                        tgDat(iUpper).iPdDay(0) = rst!datPdMon   '(15).Value
                        tgDat(iUpper).iPdDay(1) = rst!datPdTue   '(16).Value
                        tgDat(iUpper).iPdDay(2) = rst!datPdWed   '(17).Value
                        tgDat(iUpper).iPdDay(3) = rst!datPdThu   '(18).Value
                        tgDat(iUpper).iPdDay(4) = rst!datPdFri   '(19).Value
                        tgDat(iUpper).iPdDay(5) = rst!datPdSat   '(20).Value
                        tgDat(iUpper).iPdDay(6) = rst!datPdSun   '(21).Value
                        tgDat(iUpper).sPdDayFed = rst!datPdDayFed
                        If (tgDat(iUpper).iFdStatus <= 1) Or (tgDat(iUpper).iFdStatus = 9) Or (tgDat(iUpper).iFdStatus = 10) Then
                            If Second(rst!datPdStTime) = 0 Then
                                tgDat(iUpper).sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWOSecForm)
                            Else
                                tgDat(iUpper).sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWSecForm)
                            End If
                            'If tgDat(iUpper).iFdStatus = 1 Then
                                If Second(rst!datPdEdTime) = 0 Then
                                    tgDat(iUpper).sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWOSecForm)
                                Else
                                    tgDat(iUpper).sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWSecForm)
                                End If
                            'Else
                            '    tgDat(iUpper).sPdETime = ""
                            'End If
                        Else
                            tgDat(iUpper).sPdSTime = ""
                            tgDat(iUpper).sPdETime = ""
                        End If
                        tgDat(iUpper).iAirPlayNo = rst!datAirPlayNo
                        If tgDat(iUpper).iAirPlayNo > Val(edcNoAirPlays.Text) Then
                            edcNoAirPlays.Text = tgDat(iUpper).iAirPlayNo
                        End If
                        tgDat(iUpper).sEstimatedTime = rst!datEstimatedTime
                        '7/15/14
                        If (rst!datEmbeddedOrROS = Null) Then
                            tgDat(iUpper).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                        ElseIf Trim$(rst!datEmbeddedOrROS) = "" Then
                            tgDat(iUpper).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                        Else
                            tgDat(iUpper).sEmbeddedOrROS = rst!datEmbeddedOrROS
                        End If
                        ilRet = mGetET(tgDat(iUpper))
                        iUpper = iUpper + 1
                        ReDim Preserve tgDat(0 To iUpper) As DAT
                        rst.MoveNext
                    Wend
                    
                End If
                
            Else
                'If optTimeType(1).Value Then
                '    sSDate = Format$(txtOnAirDate.Text, sgShowDateForm)
                '    gGetAvails lmAttCode, imShttCode, imVefCode, sSDate
                'Else
                '    frmAffDP.Show vbModal
                'End If
                
                ilAdjustCDStartTime = False
                sSDate = Format$(txtOnAirDate.Text, sgShowDateForm)
                '2/8/05: For CD/Tape stop adjusting the Sold (Feed) time by Zone
                'gGetAvails lmAttCode, imShttCode, imVefCode, sSDate
                If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                Else
                    'If optTimeType(2).Value Then
                    '    gGetAvails lmAttCode, imShttCode, imVefCode, imVefCombo, sSDate, False
                    'Else
                        gGetAvails lmAttCode, imShttCode, imVefCode, imVefCombo, sSDate, True
                    'End If
                End If
                lgPledgeAttCode = lmAttCode
                igDayPartShttCode = imShttCode
                igDayPartVefCode = imVefCode
                igDayPartVefCombo = imVefCombo
                If edcNoAirPlays.Text <> "" Then
                    igNoAirPlays = Val(edcNoAirPlays.Text)
                    If cbcAirPlayNo.ListIndex > 0 Then
                        igDefaultAirPlayNo = Val(cbcAirPlayNo.List(cbcAirPlayNo.ListIndex))
                    Else
                        igDefaultAirPlayNo = 1
                    End If
                Else
                    igNoAirPlays = 1
                    igDefaultAirPlayNo = 1
                End If
                'If optTimeType(0).Value Then
                If ilTimeTypeIndex = 0 Then
                    'Daypart radio button on the Agrrement Pledge screen
                    igLiveDayPart = True
                    igCDTapeDayPart = False
                    igPledgeExist = False
                    If (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "") Then
                        igPledgeExist = True
                    End If
                    mGetDayParts
                    'If they pick Dayparts then the status will always be delayed
                    'per Dick on 8/27/03
                    sStatus = "2-Air Delay B'cast"  '"2-Air In Daypart"
                    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                        frmAgmntPledgeSpec.Show vbModal
                    Else
                        frmAffDP.Show vbModal
                    End If
                    If igReturnPledgeStatus = 1 Then
                        sStatus = tgStatusTypes(9).sName
                    ElseIf igReturnPledgeStatus = 2 Then
                        sStatus = tgStatusTypes(10).sName
                    End If
                Else
                    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                        igLiveDayPart = False
                        igCDTapeDayPart = False
                        igPledgeExist = False
                        If (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "") Then
                            igPledgeExist = True
                        End If
                        mGetDayParts
                        frmAgmntPledgeSpec.Show vbModal
                    End If
                End If
                
            End If
            imDatLoaded = True
            If igReload = False Then
                bmPledgeDataChgd = blPledgeDataChgd
                Exit Sub
            End If
            mPopPledge
        Else
            mPopulateEvents
            imDatLoaded = True
        End If
        mMousePointer vbDefault
    End If
    imFieldChgd = True
    bmPledgeDataChgd = blPledgeDataChgd
    Exit Sub
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mLoadPledge"
    Exit Sub
End Sub


Private Sub cmdDelete_Click()
'    Dim iLoop As Integer
'    Dim iRow As Integer
'    Dim iRows As Integer
'
'    If grdDayparts.Rows <= 0 Then
'        Exit Sub
'    End If
'    iRow = grdDayparts.Row + grdDayparts.AddItemRowIndex(grdDayparts.FirstRow)
'    If (iRow < 0) Or (iRow > grdDayparts.Rows - 1) Then
'        Exit Sub
'    End If
'    iRows = grdDayparts.Rows
'    grdDayparts.SetFocus
'    grdDayparts.DeleteSelected
'
'    If iRows <> grdDayparts.Rows Then
'        For iLoop = iRow To UBound(tgDat) - 2 Step 1
'            tgDat(iLoop) = tgDat(iLoop + 1)
'        Next iLoop
'        If UBound(tgDat) > 0 Then
'            ReDim Preserve tgDat(0 To UBound(tgDat) - 1) As DAT
'        Else
'            ReDim Preserve tgDat(0 To 0) As DAT
'        End If
'    End If
End Sub

Private Sub cmdClearAvails_Click()

    Dim ilRet As Integer
    
    'If UBound(tgDat) > LBound(tgDat) Then
    If (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "") Then
        'ilRet = gMsgBox("Warning: You Are About To Clear All Daypart or Avails Information.!", vbOKCancel)
        ilRet = gMsgBox("Warning: You Are About To Clear All Pledge Information.!", vbOKCancel)
        If ilRet = vbCancel Then
            Exit Sub
        End If
        sgAirPlay1TimeType = ""         '12-10-11
    End If
    'optTimeType(0).Value = False
    'optTimeType(1).Value = False
    sgCDStartTime = ""
'    grdDayparts.RemoveAll
    'gGrid_Clear grdPledge, True
    'ReDim tgDat(0 To 0) As DAT
    mClearPledgeGrid
    '6/26/18:   Set pledge changed flag
    bmPledgeDataChgd = True

End Sub

Private Sub cmdClearAvails_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdErase_Click()
    Dim iRet As Integer
    
    On Error GoTo ErrHand

    If IsAgmntDirty = True Then
        gSetMousePointer grdPledge, grdMulticast, vbHourglass
        SQLQuery = "SELECT Count(astCode) FROM ast WHERE (astAtfCode = " & lmAttCode & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value > 0 Then
            gSetMousePointer grdPledge, grdMulticast, vbDefault
            gMsgBox "Agreement can not be Erased as Spots have been Posted against the Agreement.", vbOKOnly
            Exit Sub
        End If
        SQLQuery = "SELECT Count(cpttCode) FROM cptt WHERE (cpttAtfCode = " & lmAttCode & " AND ((cpttPostingStatus <> 0) OR (cpttStatus = 1))" & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value > 0 Then
            gSetMousePointer grdPledge, grdMulticast, vbDefault
            gMsgBox "Agreement can not be Erased as CP's have been Posted against the Agreement.", vbOKOnly
            Exit Sub
        End If
        gSetMousePointer grdPledge, grdMulticast, vbDefault
        iRet = gMsgBox("Remove the agreement?", vbYesNo)
        If iRet = vbNo Then
            Exit Sub
        End If
        gSetMousePointer grdPledge, grdMulticast, vbHourglass
        'D.S. 10/25/04
        'igChangedNewErased values  1 = changed, 2 = new, 3 = erased
        igChangedNewErased = 3
        
        ' JD 12-18-2006 Added new function to properly remove an agreement.
        If Not gDeleteAgreement(lmAttCode, "AffErrorLog.Txt") Then
            gLogMsg "FAIL: cmdErase_Click - Unable to delete att code " & lmAttCode, "AffAgreementLog.Txt", False
        End If
'        cnn.BeginTrans
'        'D.S. 10/25/04
'        SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & lmAttCode & ")"
'        cnn.Execute SQLQuery, rdExecDirect
'        SQLQuery = "DELETE FROM Cptt WHERE (cpttAtfCode = " & lmAttCode & ")"
'        cnn.Execute SQLQuery, rdExecDirect
'        SQLQuery = "DELETE FROM dat WHERE (datAtfCode = " & lmAttCode & ")"
'        cnn.Execute SQLQuery, rdExecDirect
'        SQLQuery = "DELETE FROM Att WHERE (AttCode = " & lmAttCode & ")"
'        cnn.Execute SQLQuery, rdExecDirect
'        'Doug (9/25/06)- Remove spot records from Web Posting - Comment only
'        '                The first test above eliminates any possible exports as checking for any ast.
'
'        cnn.CommitTrans
        'Repopulate
        DoEvents
        If optSSSort(0).Value Then
            optSSSort_Click 0
        Else
            optSSSort_Click 1
        End If
        lmAttCode = 0
    End If
    
    gSetMousePointer grdPledge, grdMulticast, vbDefault
    Exit Sub
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", " Agreement-cmdErase"
End Sub

Private Sub cmdErase_GotFocus()
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdFastAdd_Click()
    
    frmFastAdd.Show vbModal
    
    If optExAll(1).Value = False Then   'Affiliates, then populate cboSSSort, other it is not required
        optExAll_Click 0
    Else
        optExAll_Click 1
    End If
    
End Sub

Private Sub cmdFastAdd_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdFastEnd_Click()
    frmFastEnd.Show vbModal
    
    If optExAll(1).Value = False Then   'Affiliates, then populate cboSSSort, other it is not required
        optExAll_Click 0
    Else
        optExAll_Click 1
    End If
    
End Sub

Private Sub cmdFastEnd_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdGenPassword_Click()
    
    Dim slPassword As String
    
    'D.S. 01/12/09
    'slPassword = gGeneratePassword(4, 1)
    'txtLogPassword.text = slPassword
    slPassword = ""

End Sub

Private Sub cmdNew_Click()
    Dim iRet As Integer
    Dim sRange As String
    
    If bmSaving Then
        Exit Sub
    End If
    bmSaving = True
    If imShttCode <= 0 Then
        gMsgBox "Station must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        bmSaving = False
        Exit Sub
    End If
    If imVefCode <= 0 Then
        gMsgBox "Vehicle must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        bmSaving = False
        Exit Sub
    End If
    'mMousePointer vbHourglass
    mMousePointer vbHourglass
    imSource = 0
    IsAgmntDirty = False
    lmAttCode = 0
    'D.S. 10/25/04
    'igChangedNewErased - 1 = changed, 2 = new, 3 = erased
    igChangedNewErased = 2
    iRet = mSave(False)
    If iRet Then
        'Repopulate
        DoEvents
        If optSSSort(0).Value Then
            optSSSort_Click 0
        Else
            optSSSort_Click 1
        End If
    Else
        If gIsMulticast(imShttCode) Then
            SQLQuery = "SELECT * FROM att"
            SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                mSetMulticast
                If rst!attMulticast = "Y" Then
                    rbcMulticast(0).Value = True
                    smSvMulticast = "Y"
                Else
                    rbcMulticast(1).Value = True
                End If
                iRet = mPopMulticast()
            End If
        End If
    End If
    imDateChgd = False
    bmSaving = False
    'mMousePointer vbDefault
    mMousePointer vbDefault
End Sub

Private Sub cmdNew_GotFocus()
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdRemap_Click()
    frmAvRemap.Show vbModal
End Sub

Private Sub cmdRemap_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub cmdSave_Click()
    Dim iRet As Integer
    Dim iPos As Integer
    Dim sChar As String
    Dim sRange As String
    Dim sStr As String
    Dim sOldStr As String
    Dim iRow As Integer
    Dim iLoop As Integer
    Dim iTRow As Integer

    If bmSaving Then
        Exit Sub
    End If
    bmSaving = True
    If imShttCode <= 0 Then
        gMsgBox "Station must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        bmSaving = False
        Exit Sub
    End If
    
    If imVefCode <= 0 Then
        gMsgBox "Vehicle must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        bmSaving = False
        Exit Sub
    End If
   
    'mMousePointer vbHourglass
    mMousePointer vbHourglass
    imSource = 1
    'D.S. 10/25/04
    'igChangedNewErased - 1 = changed, 2 = new, 3 = erased
    igChangedNewErased = 1
    iRet = mSave(False)
    If iRet And (imDateChgd Or (imRepopAgnmt = "Y")) Then
        'iTRow = cboSSSort.AddItemRowIndex(cboSSSort.FirstRow)
        'iRow = cboSSSort.Row
        'Repopulate
        DoEvents
        If optSSSort(0).Value Then
            optSSSort_Click 0
        Else
            optSSSort_Click 1
        End If
        'cboSSSort.FirstRow = cboSSSort.AddItemBookmark(iTRow)
        'For iLoop = 0 To iRow Step 1
        '    If iLoop = 0 Then
        '        cboSSSort.MoveFirst
        '    Else
        '        cboSSSort.MoveNext
        '    End If
        'Next iLoop
    Else
        If gIsMulticast(imShttCode) Then
            SQLQuery = "SELECT * FROM att"
            SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                mSetMulticast
                If rst!attMulticast = "Y" Then
                    rbcMulticast(0).Value = True
                    smSvMulticast = "Y"
                Else
                    rbcMulticast(1).Value = True
                End If
                iRet = mPopMulticast()
            End If
        End If
    End If
    imDateChgd = False
    bmSaving = False
    'mMousePointer vbDefault
    mMousePointer vbDefault
End Sub


Private Sub cmdSave_GotFocus()
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub edcNoAirPlays_GotFocus()
    mETSetShow
    mPledgeSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNoAirPlays_LostFocus()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    
    If Val(edcNoAirPlays.Text) < 1 Then ' JD 01-18-23 TTP 6280
        gMsgBox "The # of air plays must be a valid number greater than zero."
        edcNoAirPlays.Text = "1"
        edcNoAirPlays.SetFocus
        Exit Sub
    End If
    
    If cbcAirPlayNo.ListCount - 1 <> Val(edcNoAirPlays.Text) Then
        ilIndex = cbcAirPlayNo.ListIndex
        cbcAirPlayNo.Clear
        cbcAirPlayNo.AddItem "[All]"
        For ilLoop = 1 To Val(edcNoAirPlays.Text) Step 1
            cbcAirPlayNo.AddItem ilLoop
        Next ilLoop
        If ilIndex <= cbcAirPlayNo.ListCount - 1 Then
            cbcAirPlayNo.ListIndex = ilIndex
        Else
            'Missing: Ask which Air Play should be removed
        End If
    End If
End Sub


Private Sub grdEvent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim llPet As Long
    If Y < grdEvent.RowHeight(0) + grdEvent.RowHeight(1) Then
        llCol = grdEvent.MouseCol
        mEventSortCol CInt(llCol)
    Else
        If Trim$(grdEvent.TextMatrix(grdEvent.MouseRow, EVTEVENTNOINDEX)) <> "" Then
            llRow = grdEvent.MouseRow
            llCol = grdEvent.MouseCol
            grdEvent.Row = llRow
            llPet = Val(grdEvent.TextMatrix(llRow, EVTPETINFOINDEX))
            If llCol = EVTCARRYINDEX Then
                grdEvent.Col = EVTCARRYINDEX
                If grdEvent.CellBackColor <> LIGHTGREEN Then
                    bmPledgeDataChgd = True
                    tmPetInfo(llPet).sDeclaredStatus = "Y"
                    tmPetInfo(llPet).sChanged = "Y"
                End If
                grdEvent.CellBackColor = LIGHTGREEN
                grdEvent.Col = EVTNOTCARRIEDINDEX
                grdEvent.CellBackColor = vbWhite
                grdEvent.Col = EVTUNDECIDEDINDEX
                grdEvent.CellBackColor = vbWhite
            ElseIf llCol = EVTNOTCARRIEDINDEX Then
                grdEvent.Col = EVTCARRYINDEX
                grdEvent.CellBackColor = vbWhite
                grdEvent.Col = EVTNOTCARRIEDINDEX
                If grdEvent.CellBackColor <> vbRed Then
                    bmPledgeDataChgd = True
                    tmPetInfo(llPet).sDeclaredStatus = "N"
                    tmPetInfo(llPet).sChanged = "Y"
                End If
                grdEvent.CellBackColor = vbRed
                grdEvent.Col = EVTUNDECIDEDINDEX
                grdEvent.CellBackColor = vbWhite
            ElseIf llCol = EVTUNDECIDEDINDEX Then
                grdEvent.Col = EVTCARRYINDEX
                grdEvent.CellBackColor = vbWhite
                grdEvent.Col = EVTNOTCARRIEDINDEX
                grdEvent.CellBackColor = vbWhite
                grdEvent.Col = EVTUNDECIDEDINDEX
                If grdEvent.CellBackColor <> ORANGE Then
                    bmPledgeDataChgd = True
                    tmPetInfo(llPet).sDeclaredStatus = "U"
                    tmPetInfo(llPet).sChanged = "Y"
                End If
                grdEvent.CellBackColor = ORANGE
                
            End If
        End If
    End If
End Sub

Private Sub grdMulticast_Click()
    Dim llRow As Long
    Dim llCol As Long
    Dim slDates As String
    Dim llSvAttCode As Long
    Dim llAttCode As Long
    Dim ilSvShttCode As Integer
    Dim ilShttCode As Integer
    Dim llMouseRow As Long
    
    If sgUstWin(2) <> "I" Then
        grdMulticast.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        grdMulticast.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    llMouseRow = grdMulticast.MouseRow
    gSetMousePointer grdPledge, grdMulticast, vbHourglass
    If (llMouseRow >= grdMulticast.FixedRows) And (llMouseRow < grdMulticast.Rows) Then
        If grdMulticast.TextMatrix(llMouseRow, MCCALLLETTERSINDEX) <> "" Then
            If grdMulticast.TextMatrix(llMouseRow, MCSELECTEDINDEX) <> "1" Then
                grdMulticast.Row = llMouseRow
                grdMulticast.Col = grdMulticast.MouseCol
                llAttCode = Val(grdMulticast.TextMatrix(grdMulticast.Row, MCATTCODEINDEX))
                ilShttCode = Val(grdMulticast.TextMatrix(grdMulticast.Row, MCSHTTCODEINDEX))
                If grdMulticast.CellBackColor = vbWhite Then
                    grdMulticast.TextMatrix(llMouseRow, MCSELECTEDINDEX) = "1"
                    slDates = grdMulticast.TextMatrix(grdMulticast.Row, MCDATERANGEINDEX)
                    If InStr(1, slDates, "TFN", vbBinaryCompare) > 0 Then
                        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
                            If llRow <> grdMulticast.Row Then
                                If grdMulticast.TextMatrix(llMouseRow, MCSELECTEDINDEX) = "1" Then
                                    slDates = grdMulticast.TextMatrix(llMouseRow, MCDATERANGEINDEX)
                                    If InStr(1, slDates, "TFN", vbBinaryCompare) > 0 Then
                                        grdMulticast.TextMatrix(llMouseRow, MCSELECTEDINDEX) = "0"
                                        mMCPaintRowColor llRow
                                    End If
                                End If
                            End If
                        Next llRow
                        'Repopulate Grid
                        llSvAttCode = lmAttCode
                        ilSvShttCode = imShttCode
                        lmAttCode = llAttCode
                        imShttCode = ilShttCode
                        imDatLoaded = False
                        mLoadPledge True, 1
                        lmAttCode = llSvAttCode
                        imShttCode = ilSvShttCode
                        bmPledgeDataChgd = False
                    End If
                End If
            Else
                If smSvMulticast <> "Y" Then
                    grdMulticast.TextMatrix(llMouseRow, MCSELECTEDINDEX) = "0"
                End If
            End If
            imFieldChgd = True
            mMCPaintRowColor llMouseRow
        End If
    End If
    grdMulticast.Redraw = True
    gSetMousePointer grdPledge, grdMulticast, vbDefault
End Sub

Private Sub grdMulticast_EnterCell()
    imIgnoreTabs = False
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub lbcAudioDelivery_Click()
    mEnableDeliveryOptions True
End Sub

Private Sub lbcAudioDelivery_ItemCheck(Item As Integer)
    '7902
    If imInChg Then
        Exit Sub
    End If
    'disables space bar
    If Not bmInMultiListChange Then
        With lbcAudioDelivery
            bmInMultiListChange = True
            If .SelCount < 2 Then
                .Selected(Item) = Not .Selected(Item)
            Else
                .Selected(Item) = False
            End If
'            If .SelCount > 2 Then
'                MsgBox "Can only select up to two delivery vendors "
'                .Selected(Item) = False
'            End If
            bmInMultiListChange = False
        End With
    '8017
    Else
       If lbcLogDelivery.SelCount > 2 Then
            lbcLogDelivery.Selected(Item) = False
       End If
    End If
End Sub
Private Sub lbcAudioDelivery_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bmInMultiListChange = True
    With lbcAudioDelivery
'        If .SelCount > 2 Then
'            MsgBox "Can only select up to two delivery vendors "
'        Else
'             .Selected(.ListIndex) = Not .Selected(.ListIndex)
'       End If
        '8017
        If .ListIndex > -1 Then
            If .SelCount > 2 Then
                 MsgBox "Can only select up to two delivery vendors "
                ' .Selected(.ListIndex) = False
            Else
                  .Selected(.ListIndex) = Not .Selected(.ListIndex)
            End If
'            If .SelCount > 2 Then
'                .Selected(.ListIndex) = False
'            End If
        End If
    End With
    bmInMultiListChange = False
End Sub

Private Sub lbcLogDelivery_ItemCheck(Item As Integer)
    '7902
    If imInChg Then
        Exit Sub
    End If
    'disables space bar
    If Not bmInMultiListChange Then
        With lbcLogDelivery
            bmInMultiListChange = True
            If .SelCount < 2 Then
                .Selected(Item) = Not .Selected(Item)
            Else
                .Selected(Item) = False
            End If
'            If .SelCount > 2 Then
'                MsgBox "Can only select up to two delivery vendors "
'                .Selected(Item) = False
'            End If
            bmInMultiListChange = False
        End With
    '8017
    Else
       If lbcLogDelivery.SelCount > 2 Then
            lbcLogDelivery.Selected(Item) = False
       End If
    End If
End Sub
Private Sub lbcLogDelivery_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bmInMultiListChange = True
    With lbcLogDelivery
        '8017
        If .ListIndex > -1 Then
            If .SelCount > 2 Then
                 MsgBox "Can only select up to two delivery vendors "
                ' .Selected(.ListIndex) = False
            Else
                  .Selected(.ListIndex) = Not .Selected(.ListIndex)
            End If
'            If .SelCount > 2 Then
'                .Selected(.ListIndex) = False
'            End If
        End If
    End With
    bmInMultiListChange = False
End Sub

Private Sub optIDCGroup_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optIDCGroup_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optService_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optService_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub pbcEmbeddedOrROS_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        grdPledge.Text = "E"
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcEmbeddedOrROS_Paint
    ElseIf KeyAscii = Asc("R") Or (KeyAscii = Asc("r")) Then
        grdPledge.Text = "R"
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcEmbeddedOrROS_Paint
    End If
    If KeyAscii = Asc(" ") Then
        slStr = grdPledge.Text
        If slStr = "E" Then
            grdPledge.Text = "R"
        ElseIf slStr = "R" Then
            grdPledge.Text = "E"
        Else
            grdPledge.Text = "E"
        End If
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcEmbeddedOrROS_Paint
    End If
End Sub

Private Sub pbcEmbeddedOrROS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    slStr = grdPledge.Text
    If slStr = "E" Then
        grdPledge.Text = "R"
    ElseIf slStr = "R" Then
        grdPledge.Text = "E"
    Else
        grdPledge.Text = "E"
    End If
    imFieldChgd = True
    bmPledgeDataChgd = True
    pbcEmbeddedOrROS_Paint
End Sub

Private Sub pbcEmbeddedOrROS_Paint()
    pbcEmbeddedOrROS.Cls
    pbcEmbeddedOrROS.CurrentX = 15
    pbcEmbeddedOrROS.CurrentY = 0 'fgBoxInsetY
    If grdPledge.Text = "E" Then
        pbcEmbeddedOrROS.Print "Embedded"
    ElseIf grdPledge.Text = "R" Then
        pbcEmbeddedOrROS.Print "ROS"
    Else
        pbcEmbeddedOrROS.Print ""
    End If
End Sub

Private Sub rbcAudio_Click(Index As Integer)
    '7701 removed
'    lacIDCReceiverID.Visible = False
'    txtIDCReceiverID.Visible = False
'    frcIdcGroup.Visible = False
'    frcVoiceTracked.Visible = False
'    lacXDReceiverID.Visible = False
'    txtXDReceiverID.Visible = False
'    If rbcAudio(Index).Value Then
'        If (Index = 0) Or (Index = 1) Then
'            frcVoiceTracked.Visible = True
'            lacXDReceiverID.Visible = True
'            txtXDReceiverID.Visible = True
'            ckcSendDelays.Enabled = True
'        Else
'            ckcSendDelays.Enabled = False
'            ckcSendDelays.Value = vbUnchecked
'        End If
'        If Index = 3 Then
'            lacIDCReceiverID.Visible = True
'            txtIDCReceiverID.Visible = True
'            frcIdcGroup.Visible = True
'        End If
'    End If
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mPresetAgreement
End Sub

Private Sub txtAirPlay_Change()
    grdPledge.TextMatrix(lmEnableRow, lmEnableCol) = txtAirPlay.Text
End Sub

Private Sub txtAirPlay_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtAirPlay_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtIDCReceiverID_Change()
    imFieldChgd = True
    mIDCShowGroup Len(txtIDCReceiverID) > 0
End Sub

Private Sub txtIDCReceiverID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtIDCReceiverID_LostFocus()
    '6466 if not empty (so going to show grouping), see if needed
    mIDCShowGroup Len(txtIDCReceiverID) > 0
End Sub

Private Sub txtStartDate_LostFocus()
    If txtOnAirDate.Text = "" Then
        If txtStartDate.Text <> "" Then
            If gIsDate(txtStartDate.Text) = True Then
                txtOnAirDate.Text = txtStartDate.Text
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    smCurDate = Format(gNow(), sgShowDateForm)
    If imFirstTime Then
        bgAgreementVisible = True
        ReDim imAffShttCode(0 To 0) As Integer
        ReDim tmETAvailInfo(0 To 0) As ETAVAILINFO
        frcTab(0).BorderStyle = 0
        frcTab(1).BorderStyle = 0
        frcTab(2).BorderStyle = 0
        frcTab(3).BorderStyle = 0
        frcTab(4).BorderStyle = 0
        frcET.BorderStyle = 0
        frcET.Caption = ""
        frcEvent.BorderStyle = 0
        frcEvent.Caption = ""
        'Hide column 20
        grdPledge.ColWidth(CODEINDEX) = 0
        grdPledge.ColWidth(SORTINDEX) = 0
        grdPledge.ColWidth(MONFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(TUEFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(WEDFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(THUFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(FRIFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(SATFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(SUNFDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(STARTTIMEFDINDEX) = grdPledge.Width / 13 '12
        grdPledge.ColWidth(ENDTIMEFDINDEX) = grdPledge.Width / 13   '12
        'Skip status
        grdPledge.ColWidth(AIRPLAYINDEX) = grdPledge.Width / 40
        grdPledge.ColWidth(MONPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(TUEPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(WEDPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(THUPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(FRIPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(SATPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(SUNPDINDEX) = grdPledge.Width / 30   '29
        grdPledge.ColWidth(DAYFEDINDEX) = grdPledge.Width / 27  '25
        grdPledge.ColWidth(STARTTIMEPDINDEX) = grdPledge.Width / 13 '12
        grdPledge.ColWidth(ENDTIMEPDINDEX) = grdPledge.Width / 13   '12
        grdPledge.ColWidth(ESTIMATETIMEINDEX) = grdPledge.Width / 15
        grdPledge.ColWidth(EMBEDDEDORROSINDEX) = grdPledge.Width / 30
        grdPledge.ColWidth(ESTIMATEDFIRSTINDEX) = 0
        grdPledge.ColWidth(STATUSINDEX) = grdPledge.Width - GRIDSCROLLWIDTH  '(5 * grdStation.Columns(6).Width) / 6
        For ilCol = MONFDINDEX To EMBEDDEDORROSINDEX Step 1
            If ilCol <> STATUSINDEX Then
                grdPledge.ColWidth(STATUSINDEX) = grdPledge.ColWidth(STATUSINDEX) - grdPledge.ColWidth(ilCol)
            End If
        Next ilCol
        grdPledge.ColWidth(STATUSINDEX) = grdPledge.ColWidth(STATUSINDEX) - 15
        'Set Titles
        gGrid_AlignAllColsLeft grdPledge
        For ilCol = 0 To grdPledge.Cols - 1 Step 1
            imColPos(ilCol) = grdPledge.ColPos(ilCol)
        Next ilCol
        For ilCol = MONFDINDEX To ENDTIMEFDINDEX Step 1
            grdPledge.TextMatrix(0, ilCol) = "Feed"
        Next ilCol
        For ilCol = STATUSINDEX To ENDTIMEPDINDEX Step 1
            grdPledge.TextMatrix(0, ilCol) = "Pledge"
        Next ilCol
        grdPledge.TextMatrix(1, MONFDINDEX) = "M"
        grdPledge.TextMatrix(1, TUEFDINDEX) = "T"
        grdPledge.TextMatrix(1, WEDFDINDEX) = "W"
        grdPledge.TextMatrix(1, THUFDINDEX) = "Th"
        grdPledge.TextMatrix(1, FRIFDINDEX) = "F"
        grdPledge.TextMatrix(1, SATFDINDEX) = "Sa"
        grdPledge.TextMatrix(1, SUNFDINDEX) = "Su"
        grdPledge.TextMatrix(1, STARTTIMEFDINDEX) = "Start Time"
        grdPledge.TextMatrix(1, ENDTIMEFDINDEX) = "End Time"
        grdPledge.TextMatrix(1, STATUSINDEX) = "Status"
        grdPledge.TextMatrix(1, AIRPLAYINDEX) = "#"
        grdPledge.TextMatrix(1, MONPDINDEX) = "M"
        grdPledge.TextMatrix(1, TUEPDINDEX) = "T"
        grdPledge.TextMatrix(1, WEDPDINDEX) = "W"
        grdPledge.TextMatrix(1, THUPDINDEX) = "Th"
        grdPledge.TextMatrix(1, FRIPDINDEX) = "F"
        grdPledge.TextMatrix(1, SATPDINDEX) = "Sa"
        grdPledge.TextMatrix(1, SUNPDINDEX) = "Su"
        grdPledge.TextMatrix(1, DAYFEDINDEX) = "B/A"
        grdPledge.TextMatrix(1, STARTTIMEPDINDEX) = "Start Time"
        grdPledge.TextMatrix(1, ENDTIMEPDINDEX) = "End Time"
        grdPledge.TextMatrix(0, ESTIMATETIMEINDEX) = "Estimated"
        grdPledge.TextMatrix(1, ESTIMATETIMEINDEX) = "Time"
        grdPledge.TextMatrix(0, EMBEDDEDORROSINDEX) = "D"
        grdPledge.Row = 0
        grdPledge.MergeCells = 2    'flexMergeRestrictColumns
        grdPledge.MergeRow(0) = True
        grdPledge.Row = 0
        grdPledge.Col = MONFDINDEX
        grdPledge.CellAlignment = flexAlignRightTop
        grdPledge.Row = 0
        grdPledge.Col = STATUSINDEX
        grdPledge.CellAlignment = flexAlignRightTop
        gGrid_IntegralHeight grdPledge
        gGrid_Clear grdPledge, True
        For ilRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
            grdPledge.Row = ilRow
            For ilCol = MONFDINDEX To SUNFDINDEX Step 1
                grdPledge.Col = ilCol
                grdPledge.CellFontName = "Monotype Sorts"
            Next ilCol
            For ilCol = MONPDINDEX To SUNPDINDEX Step 1
                grdPledge.Col = ilCol
                grdPledge.CellFontName = "Monotype Sorts"
            Next ilCol
        Next ilRow
        grdPledge.Row = grdPledge.FixedRows - 1
        grdPledge.Col = STARTTIMEFDINDEX
        grdPledge.CellBackColor = LIGHTBLUE
        grdPledge.Col = AIRPLAYINDEX
        grdPledge.CellBackColor = LIGHTBLUE
        grdPledge.Col = STARTTIMEPDINDEX
        grdPledge.CellBackColor = LIGHTBLUE
        grdPledge.Row = grdPledge.FixedRows
        ckcDay.Height = 180
        ckcDay.Width = 180
        
        grdET.ColWidth(ETAVAILINFOINDEX) = 0
        grdET.ColWidth(ETEPTCODEINDEX) = 0
        grdET.ColWidth(ETFDDAYINDEX) = grdPledge.ColWidth(DAYFEDINDEX)
        grdET.ColWidth(ETDAYINDEX) = grdET.ColWidth(ETFDDAYINDEX)
        grdET.ColWidth(ETFDTIMEINDEX) = (5 * grdPledge.ColWidth(ESTIMATETIMEINDEX)) / 4
        grdET.ColWidth(ETTIMEINDEX) = grdET.ColWidth(ETFDTIMEINDEX)
        grdET.Width = 15    'GRIDSCROLLWIDTH + 15
        For ilCol = ETFDDAYINDEX To ETTIMEINDEX Step 1
            grdET.Width = grdET.Width + grdET.ColWidth(ilCol)
        Next ilCol
        gGrid_AlignAllColsLeft grdET
        For ilCol = 0 To grdET.Cols - 1 Step 1
            imETColPos(ilCol) = grdET.ColPos(ilCol)
        Next ilCol
        grdET.TextMatrix(0, ETFDDAYINDEX) = "Feed"
        grdET.TextMatrix(1, ETFDDAYINDEX) = "Day"
        grdET.TextMatrix(0, ETFDTIMEINDEX) = "Feed"
        grdET.TextMatrix(1, ETFDTIMEINDEX) = "Time"
        grdET.TextMatrix(0, ETDAYINDEX) = "Estimated"
        grdET.TextMatrix(1, ETDAYINDEX) = "Day"
        grdET.TextMatrix(0, ETTIMEINDEX) = "Estimated"
        grdET.TextMatrix(1, ETTIMEINDEX) = "Time"
        grdET.Row = 0
        grdET.MergeCells = 2    'flexMergeRestrictColumns
        grdET.MergeRow(0) = True
        grdET.Row = 0
        grdET.Col = ETFDDAYINDEX
        grdET.CellAlignment = flexAlignRightTop
        grdET.Row = 0
        grdET.Col = ETDAYINDEX
        grdET.CellAlignment = flexAlignRightTop
        
        grdET.Top = 0
        grdET.Left = 15
        grdET.Height = grdET.RowHeight(0) * 9 + 15
        gGrid_IntegralHeight grdET
        For ilRow = grdET.FixedRows To grdET.Rows - 1 Step 1
            grdET.Row = ilRow
            For ilCol = ETFDDAYINDEX To ETFDTIMEINDEX Step 1
                grdET.Col = ilCol
                grdET.CellBackColor = LIGHTYELLOW
            Next ilCol
        Next ilRow
        grdET.Height = grdET.Height + 30
        frcET.Left = grdPledge.Left + imColPos(ESTIMATETIMEINDEX) - grdET.Width
        frcET.Width = grdET.Width
        frcET.Height = grdET.Height 'grdET.RowHeight(0) * grdET.Rows + 15
        frcET.Top = frcTab(2).Top + grdPledge.Top + (grdPledge.Height - frcET.Height) / 2
        pbcETSTab.Left = -300
        pbcETTab.Left = -300
        cbcETDay.Clear
        cbcETDay.AddItem ("Mo")
        cbcETDay.SetItemData = 0
        cbcETDay.AddItem ("Tu")
        cbcETDay.SetItemData = 1
        cbcETDay.AddItem ("We")
        cbcETDay.SetItemData = 2
        cbcETDay.AddItem ("Th")
        cbcETDay.SetItemData = 3
        cbcETDay.AddItem ("Fr")
        cbcETDay.SetItemData = 4
        cbcETDay.AddItem ("Sa")
        cbcETDay.SetItemData = 5
        cbcETDay.AddItem ("Su")
        cbcETDay.SetItemData = 6
        cbcETDay.ReSizeFont = "A"
        cbcETDay.Height = grdET.RowHeight(grdET.FixedRows)

        grdMulticast.Width = lacPrgTimes.Left - grdMulticast.Left + 2 * imcPrt.Width
        grdMulticast.ColWidth(MCSELECTEDINDEX) = 0
        grdMulticast.ColWidth(MCSHTTCODEINDEX) = 0
        grdMulticast.ColWidth(MCATTCODEINDEX) = 0
        grdMulticast.ColWidth(MCCALLLETTERSINDEX) = grdMulticast.Width / 6
        grdMulticast.ColWidth(MCWITHINDEX) = grdMulticast.Width / 6
        grdMulticast.ColWidth(MCDATERANGEINDEX) = grdMulticast.Width / 5
        '8/6/19: Added owner column
        grdMulticast.ColWidth(MCOWNERINDEX) = grdMulticast.Width / 5
        grdMulticast.ColWidth(MCMARKETINDEX) = grdMulticast.Width - GRIDSCROLLWIDTH  '(5 * grdStation.Columns(6).Width) / 6
        For ilCol = MCCALLLETTERSINDEX To MCDATERANGEINDEX Step 1
            If ilCol <> MCMARKETINDEX Then
                grdMulticast.ColWidth(MCMARKETINDEX) = grdMulticast.ColWidth(MCMARKETINDEX) - grdMulticast.ColWidth(ilCol)
            End If
        Next ilCol
        grdMulticast.Height = grdPledge.Top - 60
        gGrid_IntegralHeight grdMulticast
        grdMulticast.Height = grdMulticast.Height + 30
        
        grdEvent.ColWidth(EVTPETINFOINDEX) = 0
        grdEvent.ColWidth(EVTSORTINDEX) = 0
        grdEvent.ColWidth(EVTEVENTNOINDEX) = grdEvent.Width * 0.06
        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
            grdEvent.ColWidth(EVTFEEDSOURCEINDEX) = grdEvent.Width * 0.08
        Else
            grdEvent.ColWidth(EVTFEEDSOURCEINDEX) = 0
        End If
        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
            grdEvent.ColWidth(EVTLANGUAGEINDEX) = grdEvent.Width * 0.1
        Else
            grdEvent.ColWidth(EVTLANGUAGEINDEX) = 0
        End If
        grdEvent.ColWidth(EVTVISITTEAMINDEX) = grdEvent.Width * 0.16
        grdEvent.ColWidth(EVTHOMETEAMINDEX) = grdEvent.Width * 0.16
        grdEvent.ColWidth(EVTAIRDATEINDEX) = grdEvent.Width * 0.1
        grdEvent.ColWidth(EVTAIRTIMEINDEX) = grdEvent.Width * 0.11
        grdEvent.ColWidth(EVTCARRYINDEX) = grdEvent.Width * 0.04
        grdEvent.ColWidth(EVTNOTCARRIEDINDEX) = grdEvent.Width * 0.04
        grdEvent.ColWidth(EVTUNDECIDEDINDEX) = grdEvent.Width * 0.04
        
        grdEvent.ColWidth(EVTVISITTEAMINDEX) = grdEvent.Width - GRIDSCROLLWIDTH - 15
        For ilCol = 0 To EVTUNDECIDEDINDEX Step 1
            If ilCol <> EVTVISITTEAMINDEX Then
                grdEvent.ColWidth(EVTVISITTEAMINDEX) = grdEvent.ColWidth(EVTVISITTEAMINDEX) - grdEvent.ColWidth(ilCol)
            End If
        Next ilCol
        grdEvent.ColWidth(EVTVISITTEAMINDEX) = (grdEvent.ColWidth(EVTHOMETEAMINDEX) + grdEvent.ColWidth(EVTVISITTEAMINDEX)) \ 2
        grdEvent.ColWidth(EVTHOMETEAMINDEX) = grdEvent.ColWidth(EVTVISITTEAMINDEX)
        'Align columns to left
        gGrid_AlignAllColsLeft grdEvent
        grdEvent.TextMatrix(0, EVTEVENTNOINDEX) = "Event"
        grdEvent.TextMatrix(1, EVTEVENTNOINDEX) = "#"
        grdEvent.TextMatrix(0, EVTFEEDSOURCEINDEX) = "Feed"
        grdEvent.TextMatrix(1, EVTFEEDSOURCEINDEX) = "Source"
        grdEvent.TextMatrix(0, EVTLANGUAGEINDEX) = "Language"
        grdEvent.TextMatrix(0, EVTVISITTEAMINDEX) = ""  '"Visiting"
        grdEvent.TextMatrix(1, EVTVISITTEAMINDEX) = ""  '"Team"
        grdEvent.TextMatrix(0, EVTHOMETEAMINDEX) = ""   '"Home"
        grdEvent.TextMatrix(1, EVTHOMETEAMINDEX) = ""   '"Team"
        grdEvent.TextMatrix(0, EVTAIRDATEINDEX) = "Air"
        grdEvent.TextMatrix(1, EVTAIRDATEINDEX) = "Date"
        grdEvent.TextMatrix(0, EVTAIRTIMEINDEX) = "Start"
        grdEvent.TextMatrix(1, EVTAIRTIMEINDEX) = "Time"
        
        grdEvent.TextMatrix(0, EVTCARRYINDEX) = "Carry Status"
        grdEvent.TextMatrix(1, EVTCARRYINDEX) = "Yes"
        grdEvent.TextMatrix(0, EVTNOTCARRIEDINDEX) = "Carry Status"
        grdEvent.TextMatrix(1, EVTNOTCARRIEDINDEX) = "No"
        grdEvent.TextMatrix(0, EVTUNDECIDEDINDEX) = "Carry Status"
        grdEvent.TextMatrix(1, EVTUNDECIDEDINDEX) = "?"
        gGrid_IntegralHeight grdEvent
        gGrid_FillWithRows grdEvent
        grdEvent.MergeCells = 2
        grdEvent.MergeRow(0) = True
        grdEvent.Row = 0
        grdEvent.Col = EVTCARRYINDEX
        grdEvent.CellAlignment = flexAlignCenterTop


        'Setting the location of the grid must be before mSort or we can't debug the user
        'control during the mSort call to setup the grid
        udcContactGrid.Move pbcArrow.Width, 0, frcTab(4).Width - pbcArrow.Width, frcTab(4).Height
        udcContactGrid.Action 2 'Init
        udcContactGrid.Source = "A"

        mSort optPSSort(0), cboPSSort, imAffShttCode()
        If sgAgreementCallSource = "S" Then
            'mPresetAgreement
            tmcStart.Enabled = True
        Else
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        End If
        lbcMulticast.Height = grdPledge.Top - 60
        imFirstTime = False
        imFieldChgd = False
    ElseIf (sgAgreementCallSource = "S") And (imFieldChgd = False) Then
        mPresetAgreement
        imFirstTime = False
        imFieldChgd = False
    End If
    sgAgreementCallSource = ""
End Sub

Private Sub Form_Click()
    mETSetShow
    mPledgeSetShow
    If pbcClickFocus.Enabled Then
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    Dim llOffset As Long
    'grdDayparts.Width = frcTab(2).Width
    Me.Visible = False
    Me.Width = Screen.Width / 1.05   '1.05  '1.15
    Me.Height = Screen.Height / 1.3 '15    '1.45    '1.25
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmAgmnt
    gCenterForm frmAgmnt
    cbcMarketRep.ReSizeFont = "A"
    cbcMarketRep.Height = txtStartDate.Height
    cbcMarketRep.SetDropDownWidth cbcMarketRep.Width
    cbcServiceRep.ReSizeFont = "A"
    cbcServiceRep.Height = txtStartDate.Height
    cbcServiceRep.SetDropDownWidth cbcServiceRep.Width
    cbcContractPDF.ReSizeFont = "A"
    cbcContractPDF.Height = txtStartDate.Height
    cbcContractPDF.SetDropDownWidth cbcContractPDF.Width
    cbcContractPDF.PopUpListDirection "A"
    cbcSeason.ReSizeFont = "A"
    cbcSeason.Height = txtStartDate.Height
    cbcSeason.SetDropDownWidth cbcSeason.Width
    lacEventKey.ForeColor = vbBlack
    lacGreenkey.Left = lacEventKey.Left + lacEventKey.Width + 120
    lacOrangeKey.Left = lacGreenkey.Left + lacGreenkey.Width + 120
    lacBlueKey.Left = lacOrangeKey.Left + lacOrangeKey.Width + 120
    lacRedKey.Left = lacBlueKey.Left + lacBlueKey.Width + 120
    '7701 remove
'    frcSendLogEMail.Left = frcPosting.Left
'    frcSendLogEMail.Top = frcPosting.Top
'    llOffset = txtIDCReceiverID.Top - txtXDReceiverID.Top
'    lacIDCReceiverID.Top = lacIDCReceiverID.Top - llOffset
'    txtIDCReceiverID.Top = txtIDCReceiverID.Top - llOffset
'    frcIdcGroup.Top = frcIdcGroup.Top - llOffset

End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim llIdx As Long
    Dim ilLoop As Integer
    Dim llOffset As Long
    Dim ilRet As Integer
    Dim ilValue1 As Integer

    '6419 temporarily hide grouping
    frcIdcGroup.Visible = False
    '6592 CBS added as option
    ckcExportTo(4).Enabled = True
    bFormWasAlreadyResized = False
    'frcExport.Visible = False
    'frcMulticast.Visible = False
    lacMulticast(0).Visible = False
    lacMulticast(1).Visible = False
    'lbcMulticast.Visible = False
    grdMulticast.Visible = False
    frcPosting.Visible = False
    frcLogType.Visible = False
    frcPostType.Visible = False
    frcSendLogEMail.Visible = False
    txtEmailAddr.Visible = False
    txtLogPassword.Visible = False
    lblWebEmail.Visible = False
    lblWebPW.Visible = False
    imIgnoreTimeTypeChg = False
    cmdGenPassword.Visible = False
    imShowGridBox = False
    lmTopRow = -1
    imFromArrow = False
    lmEnableRow = -1
    lmEnableCol = -1
    imLastPledgeColSorted = -1
    imLastPledgeSort = -1
    smPledgeByEvent = "N"
    imLastEventColSorted = -1
    imLastEventSort = -1
    imCloseListBox = True
    bmBypassTestDP = False
    bmSaving = False
    imcPrt.Picture = frmDirectory!imcPrinter.Picture
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        rbcExportType(1).Enabled = False
        ckcExportTo(2).Visible = False
        '8000
        ckcUnivision.Visible = False
    Else
        If Not gUsingWeb Then
            'rbcExportType(1).Enabled = False
            'rbcExportType(3).Enabled = False
            ckcExportTo(0).Enabled = False
            ckcExportTo(1).Enabled = False
        End If
        If Not gUsingUnivision Then
            'rbcExportType(2).Enabled = False
            ckcExportTo(2).Enabled = False
            ckcExportTo(2).Visible = False
            '8000
            ckcUnivision.Enabled = False
            ckcUnivision.Visible = False
        Else
            ckcUnivision.Enabled = True
            ckcUnivision.Visible = True
        End If
        ckcExportTo(3).Enabled = False
    End If
    mMousePointer vbHourglass
    frmAgmnt.Caption = "Agreements - " & sgClientName
        
    'Me.Width = Screen.Width / 1.1   '1.05  '1.15
    'Me.Height = Screen.Height / 1.15    '1.45    '1.25
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    optSSSort(0).Visible = False
    optSSSort(1).Visible = False
    lblSort.Visible = False
    optSSSort(0).Value = True
    optExAll(0).Value = True
    txtRetDate.Visible = False
    imBSMode = False
    imInChg = False
    imDateChgd = False
    imTabIndex = TABMAIN
    imIgnoreTabClick = False
    imIntegralSet = False
    imAllowInsert = False
    cmdAdjustDates.Visible = False
    igLiveDayPart = False

    ReDim tgDat(0 To 0) As DAT
    ReDim tmAdjustDates(0 To 0)
   'D.S. 07/11/01
    'Gather up all agreement start or end dates that are >= 1/1/2070. Show a button
    '"cmdAdjustDates" on the Agreement Airing screen "Adjust Dates?". If the user clicks
    'yes then "mAdjustDates" is called to correct year 20XX to year 19XX
    On Error GoTo ErrHand
'   D.S. 4/21/04 The code below is slow and no longer necessary
'    SQLQuery = "SELECT attAgreeStart, attAgreeEnd, attCode"
'    SQLQuery = SQLQuery & " FROM att"
'    SQLQuery = SQLQuery & " WHERE (attAgreeStart >= " & "'" & Format$("1/1/2070", sgSQLDateForm) & "'"
'    SQLQuery = SQLQuery & " Or attAgreeEnd >= " & "'" & Format$("1/1/2070", sgSQLDateForm) & "'" & ")"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If rst.EOF = False Then
'        cmdAdjustDates.Visible = True
'        llIdx = 0
'        While Not rst.EOF
'            tmAdjustDates(llIdx).lAttCode = rst!attCode
'            tmAdjustDates(llIdx).sAttAgreeStart = Format$(rst!attAgreeStart, "mm/dd/yyyy")
'            tmAdjustDates(llIdx).sAttAgreeEnd = Format$(rst!attAgreeEnd, "mm/dd/yyyy")
'            llIdx = llIdx + 1
'            ReDim Preserve tmAdjustDates(0 To llIdx)
'            rst.MoveNext
'        Wend
'    End If
    bmShowDates = False
    '3/23/15: Add Send Delays to XDS
    bmSupportXDSDelay = False
    bmDefaultEstDay = False
    SQLQuery = "SELECT siteShowContrDate, siteSupportXDSDelay, siteDefaultEstDay From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!siteShowContrDate = "Y" Then
            bmShowDates = True
        End If
        '3/23/15: Add Send Delays to XDS
        If rst!siteSupportXDSDelay = "Y" Then
            bmSupportXDSDelay = True
            'ckcSendDelays.Visible = True
            ckcSendDelays.Enabled = True
        Else
            'ckcSendDelays.Visible = False
            ckcSendDelays.Enabled = False
        End If
        If rst!siteDefaultEstDay = "Y" Then
            bmDefaultEstDay = True
        End If
    End If
    If Not bmShowDates Then
        lacStartDate.Visible = False
        txtStartDate.Visible = False
        lacDropdate.Visible = False
        txtDropDate.Visible = False
        lacEndDate.Visible = False
        txtEndDate.Visible = False
        llOffset = lacOnAirDate.Left - lacStartDate.Left
        lacOnAirDate.Left = lacOnAirDate.Left - llOffset
        txtOnAirDate.Left = txtOnAirDate.Left - llOffset
        lacOffAirDate.Left = lacOffAirDate.Left - llOffset
        txtOffAirDate.Left = txtOffAirDate.Left - llOffset
        lacHistorialDate.Left = lacHistorialDate.Left - llOffset
        txtHistorialDate.Left = txtHistorialDate.Left - llOffset
    End If
    cboSSSort.Clear
    For ilLoop = 0 To UBound(tgStatusTypes) Step 1
        tmStatusTypes(ilLoop) = tgStatusTypes(ilLoop)
        If InStr(1, tmStatusTypes(ilLoop).sName, "1-", vbTextCompare) = 1 Then
            tmStatusTypes(ilLoop).sName = "1-Air Live"
        End If
        If InStr(1, tmStatusTypes(ilLoop).sName, "2-", vbTextCompare) = 1 Then
            tmStatusTypes(ilLoop).sName = "2-Air Delay B'cast"  '"2-Air In Daypart"
        End If
    Next ilLoop
    For iLoop = 0 To UBound(tmStatusTypes) Step 1
        If (tmStatusTypes(iLoop).iStatus = 0) Or (tmStatusTypes(iLoop).iStatus = 1) Or (tmStatusTypes(iLoop).iStatus = 4) Or (tmStatusTypes(iLoop).iStatus = 8) Or (tmStatusTypes(iLoop).iStatus = 9) Or (tmStatusTypes(iLoop).iStatus = 10) Then
            'D.S. 11/6/08 Don't show "5-Not Aired Other" as a selectable status in agreements
            'Affiliate Meeting Decisions item 5) f-iii
            If InStr(1, tmStatusTypes(iLoop).sName, "5-Not Aired Other", vbTextCompare) <> 1 Then
                lbcStatus.AddItem Trim$(tmStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        End If
    Next iLoop
    
    mPopMarketRep
    mPopServiceRep
    
    ilRet = gPopVff()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    
    smContractPDFSubFolder = ""
    If sgContractPDFPath = "" Then
        lacContractPDF.Enabled = False
        cbcContractPDF.Enabled False
        cmcBrowse.Enabled = False
        lacPDFPath.Caption = "PDF Path: Not defined in Affliat.ini"
    Else
        lacPDFPath.Caption = "PDF Path: " '& sgContractPDFPath
    End If
    
    'Replaced by Market Rep, attMktRepUstCode
    'mGetAffAE
    
    imFirstTime = True
    If sgUstWin(2) <> "I" Then
        cmdRemap.Enabled = False
        cmdNew.Enabled = False
        cmdErase.Enabled = False
        cmdSave.Enabled = False
        'frcPledgeType.Enabled = False
        cmcPledgeBy(0).Enabled = False
        cmcPledgeBy(1).Enabled = False
        frcTab(0).Enabled = False
        frcTab(1).Enabled = False
        frcTab(2).Enabled = False
        'Leave frcTab(2) enabled so that user can scroll
        imcTrash.Enabled = False
    End If
    If sgUstPledge <> "Y" Then
        'frcPledgeType.Enabled = False
        cmcPledgeBy(0).Enabled = False
        cmcPledgeBy(1).Enabled = False
    End If
    mShowTabs
    ilRet = gPopAvailNames()
    'Dan M 5457
    If gIsSiteXDStation() Then
        bmIsXDSiteStation = True
    Else
        bmIsXDSiteStation = False
    End If
    smCompensation = "N"
    SQLQuery = "Select safFeatures1 From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue1 = Asc(rst!safFeatures1)
        If (ilValue1 And COMPENSATION) = COMPENSATION Then
            smCompensation = "Y"
        End If
    End If
    If smCompensation <> "Y" Then
        frcCompensation.Visible = False
    End If
    '7701 services become combo boxes
    lacExportTo.Visible = False
    For ilLoop = 0 To 6 Step 1
        ckcExportTo(ilLoop).Visible = False
    Next
    lacAudio.Visible = False
    For ilLoop = 0 To 5
        rbcAudio(ilLoop).Visible = False
    Next
    frcSendLogEMail.BorderStyle = 0
    frcLogDelivery.BorderStyle = 0
    frcAudioDelivery.BorderStyle = 0
    frcExclude.BorderStyle = 0
    frcAudioDelivery.Left = frcLogDelivery.Left
    lacLogTitle.Top = 0
    lacLogTitle.Left = lacAudioTitle.Left
    lacAudioTitle.Top = 0
    frcSendLogEMail.Left = lacVoiceTracked.Left
    lacSendLogEMail.Left = 0
    '8000
    ckcUnivision.Left = frcSendLogEMail.Left
    'lacXDReceiverID.Left = lbcAudioDelivery.Left
    lacXDReceiverID.Visible = True
    txtXDReceiverID.Visible = True
    lacIDCReceiverID.Visible = True
    txtIDCReceiverID.Visible = True
    'default email to 'yes'
    rbcSendLogEMail(0).Value = True
    If Not gUsingWeb Then
        frcSendLogEMail.Visible = False
    End If
    '8418
    mPopVendorList
   ' mLoadDeliveryServices
    'end 7701
    If sgUsingServiceAgreement <> "Y" Then
        frcService.Enabled = False
    Else
        frcService.Enabled = True
    End If
    mMousePointer vbDefault
    Exit Sub
ErrHand:
    gMsg = ""
    gHandleError "AffErrorLog.txt", "frmAgmnt-form-load"
End Sub

Private Sub Form_Resize()
    If bFormWasAlreadyResized Then
        Exit Sub
    End If
    bFormWasAlreadyResized = True
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
'    grdDayparts.Groups(0).Width = grdDayparts.Width * 0.55
'    grdDayparts.Groups(1).Width = grdDayparts.Width - grdDayparts.Groups(0).Width - grdDayparts.Columns(0).Left - GRIDSCROLLWIDTH 'Me.Width * 0.41
    lacAttCode.Left = frcSelect.Left + frcSelect.Width + 30
    lacAttCode.Top = 60
    TabStrip1.Left = frcSelect.Left
    TabStrip1.Height = cmdCancel.Top - (frcSelect.Top + frcSelect.Height + 150)  'TabStrip1.ClientTop - TabStrip1.Top + (10 * frcTab(0).Height) / 9
    TabStrip1.Width = frcSelect.Width
    frcTab(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(2).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(3).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(4).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcEvent.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    frcTab(3).BorderStyle = 0
    frcTab(4).BorderStyle = 0
    frcEvent.BorderStyle = 0
    grdMulticast.Top = frcTab(2).Top
    grdMulticast.Left = frcTab(2).Left + lacMulticast(0).Left + lacMulticast(0).Width
    pbcEmbeddedOrROS.Width = pbcArial.TextWidth("Embedded  ")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgAgreementVisible = False
    rst_Webl.Close
    rst_Ast.Close
    rst_Aet.Close
    rst_Shtt.Close
    rst_Pft.Close
    rst_ept.Close
    rst_crf.Close
   ' rst_ief.Close
    rst_Pet.Close
    rst_Gsf.Close
    Erase tmAdjustDates
    Erase tgAirPlaySpec
    Erase tgBreakoutSpec
    Erase tgDPSelection
    Erase tmETAvailInfo
    Erase tgDat
    Erase imAffVefCode
    Erase imAffShttCode
    Erase tmOverlapInfo
    Erase tmAvailDat
    Erase tmMCDat
    Erase tmPetInfo
    If sgAgreementCallSource = "S" Then
        frmStationSearch.SetFocus
    End If
    Set frmAgmnt = Nothing
End Sub



Private Sub frcTab_Click(Index As Integer)
    If Index = 2 Then
        mETSetShow
        mPledgeSetShow
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub grdET_EnterCell()
    mETSetShow
End Sub

Private Sub grdET_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llColLeftPos As Long
    
    If sgUstWin(2) <> "I" Then
        grdET.Redraw = True
        'If pbcClickFocus.Enabled Then
        '    pbcClickFocus.SetFocus
        'End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        grdET.Redraw = True
        'If pbcClickFocus.Enabled Then
        '    pbcClickFocus.SetFocus
        'End If
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdET, X, Y)
    If Not ilFound Then
        grdET.Redraw = True
        'If pbcClickFocus.Enabled Then
        '    pbcClickFocus.SetFocus
        'End If
        Exit Sub
    End If
    If Not mETColOk() Then
        grdET.Redraw = True
        'If pbcClickFocus.Enabled Then
        '    pbcClickFocus.SetFocus
        'End If
        Exit Sub
    End If
    grdET.Redraw = True
    mETEnableBox
End Sub

Private Sub grdET_Scroll()
    mETSetShow
    pbcPledgeFocus.SetFocus
End Sub

Private Sub grdPledge_Click()
    Dim llRow As Long
    
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If Trim$(grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)) <> "" Then
        If Not mPledgeColAllowed(grdPledge.Col) Then
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    Else
        If grdPledge.Col > STATUSINDEX Then
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    End If
    If grdPledge.Col >= grdPledge.Cols - 1 Then
        Exit Sub
    End If
'    lmTopRow = grdPledge.TopRow
'    llRow = grdPledge.Row
'    If grdPledge.TextMatrix(llRow, 9) = "" Then
'        grdPledge.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdPledge.TextMatrix(llRow, 9) = ""
'        grdPledge.Row = llRow + 1
'        'grdPledge.Col = 0
'        grdPledge.Redraw = True
'    End If
'    mPledgeEnableBox

End Sub

Private Sub grdPledge_EnterCell()

    'If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
    '    gMsgBox "This agreement may not be changed.  Spots have been posted against it." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "The Last Date that Spots were Posted was on " & smLastPostedDate
    '    Exit Sub
    'End If

    ''D.S. 12/17/04
    ''If optTimeType(0).Value = False And optTimeType(1).Value = False Then
    '2/5/11:  Allow input without avails defined
    'If UBound(tgDat) <= LBound(tgDat) Then
    '    gMsgBox "Either Dayparts or Avails must be selected before entering pledge information", vbOKOnly
    '    Exit Sub
    'End If

    mETSetShow
    mPledgeSetShow
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    
End Sub

Private Sub grdPledge_GotFocus()
    If grdPledge.Col >= grdPledge.Cols - 1 Then
        Exit Sub
    End If
    'grdPledge_Click
End Sub

Private Sub grdPledge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdPledge.TopRow
    grdPledge.Redraw = False
End Sub

Private Sub grdPledge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llColLeftPos As Long
    Dim ilBypass As Integer
    
    If sgUstWin(2) <> "I" Then
        grdPledge.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        grdPledge.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If (Y > grdPledge.RowHeight(0)) And (Y < grdPledge.RowHeight(1) + grdPledge.RowHeight(0)) Then
        If (X > imColPos(STARTTIMEFDINDEX)) And (X < grdPledge.ColWidth(STARTTIMEFDINDEX) + imColPos(STARTTIMEFDINDEX)) Then
            mPledgeSortCol STARTTIMEFDINDEX
        End If
        If (X > imColPos(AIRPLAYINDEX)) And (X < grdPledge.ColWidth(AIRPLAYINDEX) + imColPos(AIRPLAYINDEX)) Then
            mPledgeSortCol AIRPLAYINDEX
        End If
        If (X > imColPos(STARTTIMEPDINDEX)) And (X < grdPledge.ColWidth(STARTTIMEPDINDEX) + imColPos(STARTTIMEPDINDEX)) Then
            mPledgeSortCol STARTTIMEPDINDEX
        End If
        grdPledge.Redraw = True
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdPledge, X, Y)
    If Not ilFound Then
        grdPledge.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
        ilBypass = False
        If (grdPledge.Col = ESTIMATETIMEINDEX) And (grdPledge.CellBackColor <> LIGHTYELLOW) Then
            ilBypass = True
        End If
        If Not ilBypass Then
            gMsgBox "This agreement pledge informatiom may not be changed.  Spots have been posted against it." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "The Last Week that Spots were Posted was on " & smLastPostedDate
            grdPledge.Redraw = True
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim$(grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)) <> "" Then
        If Not mPledgeColAllowed(grdPledge.Col) Then
            grdPledge.Redraw = True
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    Else
        If grdPledge.Col > STATUSINDEX Then
            grdPledge.Redraw = True
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    End If
    If grdPledge.Col >= grdPledge.Cols - 1 Then
        grdPledge.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdPledge.TopRow
    llRow = grdPledge.Row
    If grdPledge.TextMatrix(llRow, STATUSINDEX) = "" Then
        grdPledge.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdPledge.TextMatrix(llRow, STATUSINDEX) = ""
        grdPledge.Row = llRow + 1
        'grdPledge.Col = 0
        grdPledge.Redraw = True
    End If
    grdPledge.Redraw = True
    mPledgeEnableBox
End Sub

Private Sub grdPledge_Scroll()
'    If (lmTopRow <> -1) And (lmTopRow <> grdPledge.TopRow) Then
'        grdPledge.TopRow = lmTopRow
'        lmTopRow = -1
'    End If
    If grdPledge.Redraw = False Then
        grdPledge.Redraw = True
        If lmTopRow <> -1 Then
            grdPledge.TopRow = lmTopRow
        End If
        grdPledge.Refresh
        grdPledge.Redraw = False
    End If
    If (imShowGridBox) And (grdPledge.Row >= grdPledge.FixedRows) And (grdPledge.Col >= MONFDINDEX) And (grdPledge.Col < grdPledge.Cols - 1) Then

'        If grdPledge.RowIsVisible(grdPledge.Row) Then
'            pbcArrow.Move grdPledge.Left - pbcArrow.Width, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + (grdPledge.RowHeight(grdPledge.Row) - pbcArrow.Height) / 2
'            pbcArrow.Visible = True
'            If ((grdPledge.Col >= MONFDINDEX) And (grdPledge.Col <= SUNFDINDEX)) Or ((grdPledge.Col >= MONPDINDEX) And (grdPledge.Col <= SUNPDINDEX)) Then
'                pbcDay.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 30, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 30
'                pbcDay.Visible = True
'                If ckcDay.Enabled Then
'                    ckcDay.SetFocus
'                End If
'            ElseIf (grdPledge.Col = DAYFEDINDEX) Then
'                pbcDayFed.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 30, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 30
'                pbcDayFed.Visible = True
'                If pbcDayFed.Enabled Then
'                    pbcDayFed.SetFocus
'                End If
'            ElseIf (grdPledge.Col = STARTTIMEFDINDEX) Or (grdPledge.Col = ENDTIMEFDINDEX) Or (grdPledge.Col = STARTTIMEPDINDEX) Or (grdPledge.Col = ENDTIMEPDINDEX) Then
'                txtTime.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 30, grdPledge.ColWidth(grdPledge.Col) - 30, grdPledge.RowHeight(grdPledge.Row) - 30
'                txtTime.Visible = True
'                If txtTime.Enabled Then
'                    txtTime.SetFocus
'                End If
'            Else
'                txtDropdown.Move grdPledge.Left + imColPos(grdPledge.Col) + 30, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + 15, grdPledge.ColWidth(grdPledge.Col) + 2 * grdPledge.ColWidth(grdPledge.Col + 1), grdPledge.RowHeight(grdPledge.Row) - 15
'                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
'                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
'                txtDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcStatus.Visible = True
'                If txtDropdown.Enabled Then
'                    txtDropdown.SetFocus
'                End If
'            End If
'        Else
'            If pbcPledgeFocus.Enabled Then
'                pbcPledgeFocus.SetFocus
'            End If
'            pbcDay.Visible = False
'            pbcDayFed.Visible = False
'            txtTime.Visible = False
'            txtDropdown.Visible = False
'            cmcDropDown.Visible = False
'            lbcStatus.Visible = False
'            pbcArrow.Visible = False
'        End If
'    Else
'        If pbcPledgeFocus.Enabled Then
'            pbcPledgeFocus.SetFocus
'        End If
'        pbcArrow.Visible = False
'        imFromArrow = False
        pbcClickFocus.SetFocus
    End If
End Sub



Private Sub imcPrt_Click()
    Dim iLoop As Integer
    Dim sRange As String
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sEndDate As String
    Dim sFdDay As String
    Dim sFdSTime As String
    Dim sFdETime As String
    Dim sPdDay As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim sStatus As String
    Dim llRow As Long
    Dim llCol As Long
    Dim slPdDayFed As String
    Dim ilFirstFdDay As Integer
    Dim ilFirstPdDay As Integer
    
    mMousePointer vbHourglass
    mETSetShow
    mPledgeSetShow
    lmTopRow = -1
    Printer.Print ""
    Printer.Print Tab(65); Format$(Now)
    Printer.Print ""
    If txtOnAirDate.Text = "" Then
        sRange = ""
    Else
        If gIsDate(txtOnAirDate.Text) = False Then
            sRange = ""
        Else
            sRange = Format$(txtOnAirDate.Text, sgShowDateForm)
        End If
    End If
    If txtOffAirDate.Text = "" Then
        sOffAir = "12/31/2069"
    Else
        If gIsDate(txtOffAirDate.Text) = False Then
            sOffAir = "12/31/2069"
        Else
            sOffAir = Format$(txtOffAirDate.Text, sgShowDateForm)
        End If
    End If
    If txtDropDate.Text = "" Then
        sDropDate = "12/31/2069"
    Else
        If gIsDate(txtDropDate.Text) = False Then
            sDropDate = "12/31/2069"
        Else
            sDropDate = Format$(txtDropDate.Text, sgShowDateForm)
        End If
    End If
    If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
        sEndDate = sDropDate
    Else
        sEndDate = sOffAir
    End If
    If (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/2069")) Or (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/69")) Then
        If sRange <> "" Then
            sRange = sRange & "-TFN"
        End If
    Else
        If sRange <> "" Then
            sRange = sRange & "-" & sEndDate
        Else
            sRange = "Thru " & sEndDate
        End If
    End If
    Printer.Print Trim$(cboPSSort.Text) & " " & Trim$(cboSSSort.Text) & " " & sRange
    Printer.Print ""
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        ilFirstFdDay = -1
        sFdDay = ""
        For llCol = MONFDINDEX To SUNFDINDEX Step 1
            If Trim$(grdPledge.TextMatrix(llRow, llCol)) <> "" Then
                If ilFirstFdDay = -1 Then
                    ilFirstFdDay = llCol
                End If
                Select Case llCol
                    Case MONFDINDEX
                        sFdDay = sFdDay & "Mo "
                    Case TUEFDINDEX
                        sFdDay = sFdDay & "Tu "
                    Case WEDFDINDEX
                        sFdDay = sFdDay & "We "
                    Case THUFDINDEX
                        sFdDay = sFdDay & "Th "
                    Case FRIFDINDEX
                        sFdDay = sFdDay & "Fr "
                    Case SATFDINDEX
                        sFdDay = sFdDay & "Sa "
                    Case SUNFDINDEX
                        sFdDay = sFdDay & "Su "
                End Select
            End If
        Next llCol
        sFdSTime = Trim$(grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX))
        sFdETime = Trim$(grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX))
        sStatus = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        ilFirstPdDay = -1
        sPdDay = ""
        For llCol = MONPDINDEX To SUNPDINDEX Step 1
            If Trim$(grdPledge.TextMatrix(llRow, llCol)) <> "" Then
                If ilFirstPdDay = -1 Then
                    ilFirstPdDay = llCol
                End If
                Select Case llCol
                    Case MONPDINDEX
                        sPdDay = sPdDay & "Mo "
                    Case TUEPDINDEX
                        sPdDay = sPdDay & "Tu "
                    Case WEDPDINDEX
                        sPdDay = sPdDay & "We "
                    Case THUPDINDEX
                        sPdDay = sPdDay & "Th "
                    Case FRIPDINDEX
                        sPdDay = sPdDay & "Fr "
                    Case SATPDINDEX
                        sPdDay = sPdDay & "Sa "
                    Case SUNPDINDEX
                        sPdDay = sPdDay & "Su "
                End Select
            End If
        Next llCol
        If ilFirstPdDay - MONPDINDEX < ilFirstFdDay - MONFDINDEX Then
            slPdDayFed = grdPledge.TextMatrix(llRow, DAYFEDINDEX)
        ElseIf ilFirstPdDay - MONPDINDEX > ilFirstFdDay - MONFDINDEX Then
            slPdDayFed = "A"
        Else
            slPdDayFed = ""
        End If
        If slPdDayFed = "B" Then
            slPdDayFed = "Before"
        ElseIf slPdDayFed = "A" Then
            slPdDayFed = "After"
        End If
        sPdSTime = Trim$(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX))
        sPdETime = Trim$(grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX))
        Printer.Print sFdDay; Tab(25); sFdSTime; Tab(37); sFdETime; Tab(49); sStatus; Tab(70); sPdDay; Tab(95); slPdDayFed; Tab(102); sPdSTime; Tab(114); sPdETime
    Next llRow
    Printer.EndDoc
    mMousePointer vbDefault

End Sub

Private Sub imcTrash_Click()
    Dim iLoop As Integer
    Dim llRow As Long
    Dim llRows As Long
    
    mETSetShow
    mPledgeSetShow
    llRow = grdPledge.Row
    llRows = grdPledge.Rows
    If (llRow < 0) Or (llRow > grdPledge.Rows - 1) Then
        Exit Sub
    End If
    lmTopRow = -1
    If grdPledge.TextMatrix(llRow, CODEINDEX) <> "" Then
        imFieldChgd = True
    End If
    '2/1/11: Mark as Not Carried instead of removing it
    'grdPledge.RemoveItem llRow
    'gGrid_FillWithRows grdPledge
    txtDropdown.Text = "9-Not Carried"
    mSetPledgeFromStatus llRow, STATUSINDEX
    If pbcClickFocus.Enabled Then
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub lbcStatus_Click()
    txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        If txtDropdown.Enabled Then
            txtDropdown.SetFocus
        End If
        lbcStatus.Visible = False
    End If
End Sub

Private Sub optBarCode_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optBarCode_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optCarryCmml_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optCarryCmml_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optComp_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optComp_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optExAll_Click(Index As Integer)
    Dim sStation As String
    Dim iLoop As Integer
    Dim iFound As Integer
    Dim iTest As Integer
    Dim sRange As String
    Dim sMarket As String
    Dim sEndDate As String
    Dim llStaCode As Long
    Dim slMulticast As String
    Dim ilAddSpot As Integer
    
    On Error GoTo ErrHand
    
'    If optExAll(Index).Value = False Then
'        Exit Sub
'    End If
     
    mMousePointer vbHourglass
    If sgUstWin(2) = "I" Then
        cmdRemap.Enabled = True
    End If
    ClearControls
    cboSSSort.Clear
    'DoEvents
    If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then  'PSSort contains Stations/Markets
        imVefCode = 0
        mSetEventTitles
        'imShttCode = cboPSSort.Columns(2).Text
        If imShttCode <= 0 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        If optExAll(0).Value = True Then    'Get Affiliates with agreements
            SQLQuery = "SELECT vefType, vefName, vefCode, attOnAir, attOffAir, attDropDate, attMulticast, attCode"
            'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, att"
            SQLQuery = SQLQuery & " FROM VEF_Vehicles, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode "
            SQLQuery = SQLQuery + " AND attShfCode = " & imShttCode & ")"
            SQLQuery = SQLQuery + " ORDER BY vefName"
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF = False Then
                While Not rst.EOF
                    If DateValue(gAdjYear(rst!attDropDate)) < DateValue(gAdjYear(rst!attOffAir)) Then
                        sEndDate = Format$(rst!attDropDate, sgShowDateForm)
                    Else
                        sEndDate = Format$(rst!attOffAir, sgShowDateForm)
                    End If
                    'cboSSSort.AddItem Trim$(rst(0).Value) & "|" & rst(1).Value
                    If (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                        sRange = ""
                    Else
                        sRange = Format$(Trim$(rst!attOnAir), sgShowDateForm)
                    End If
                    If (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(sEndDate)) = DateValue("12/31/69")) Then  'Or (rst!attOffAir = "12/31/69") Then
                        If sRange <> "" Then
                            sRange = sRange & "-TFN"
                        End If
                    Else
                        If sRange <> "" Then
                            sRange = sRange & "-" & sEndDate    'rst!attOffAir
                        Else
                            sRange = "Thru " & sEndDate 'rst!attOffAir
                        End If
                    End If
                    If rst!attMulticast = "Y" Then
                        slMulticast = " Multicast"
                    Else
                        slMulticast = ""
                    End If
                    
                    '10/07/14 Added If wrapper
                    If chkActive.Value = vbUnchecked Or (chkActive.Value = vbChecked And gDateValue(sEndDate) >= gDateValue(gNow())) Then
                        If sgShowByVehType = "Y" Then
                            cboSSSort.AddItem Trim$(rst!vefType) & ":" & Trim$(rst!vefName) & " " & sRange & slMulticast
                        Else
                            cboSSSort.AddItem Trim$(rst!vefName) & " " & sRange & slMulticast
                        End If
                        cboSSSort.ItemData(cboSSSort.NewIndex) = rst!attCode
                    End If
                    rst.MoveNext
                Wend
            End If
        Else        'Get Affiliate without agreements
            SQLQuery = "SELECT vefName, vefCode"
            'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, att"
            SQLQuery = SQLQuery & " FROM VEF_Vehicles, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode "
            
            If chkActive.Value = vbChecked Then
                SQLQuery = SQLQuery + " AND vefState = " & "'" & "A" & "'"
            End If
            SQLQuery = SQLQuery + " AND attShfCode = " & imShttCode & ")"
            SQLQuery = SQLQuery + " ORDER BY vefName"
            Set rst = gSQLSelectCall(SQLQuery)
            ReDim imAffVefCode(0 To 0) As Integer
            If rst.EOF = False Then
                cboSSSort.Clear
                While Not rst.EOF
                    imAffVefCode(UBound(imAffVefCode)) = rst!vefCode  'rst(1).Value
                    ReDim Preserve imAffVefCode(0 To UBound(imAffVefCode) + 1) As Integer
                    rst.MoveNext
                Wend
            End If
            For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                '10/07/14
                If (tgVehicleInfo(iLoop).sState = "A") Or (tgVehicleInfo(iLoop).sState <> "A" And optPSSort(3).Value) Then
                    iFound = False
                    For iTest = 0 To UBound(imAffVefCode) - 1 Step 1
                        If imAffVefCode(iTest) = tgVehicleInfo(iLoop).iCode Then
                            iFound = True
                            Exit For
                        End If
                    Next iTest
                    If Not iFound Then
                        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                        If mVehicleAddTest(iLoop) Then
                            cboSSSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                            cboSSSort.ItemData(cboSSSort.NewIndex) = tgVehicleInfo(iLoop).iCode
                        End If
                        'End If
                    End If
                End If
            Next iLoop
        End If
    Else
        imShttCode = 0
        'imVefCode = cboPSSort.Columns(2).Text
        If imVefCode <= 0 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        If optExAll(0).Value = True Then        'Get existing Affiliates
            'SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode, attOnAir, attOffAir, attDropDate, attCode, mktname"
            '''SQLQuery = SQLQuery + " FROM shtt, att"
            ''SQLQuery = "SELECT shttCallLetters, shttCode, attOnAir, attOffAir, attDropDate, attCode"
            'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, att"
            ''SQLQuery = SQLQuery + " FROM shtt, att"
            'SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode"
            'SQLQuery = SQLQuery + " AND attVefCode = " & imVefCode & ")"
            'Doug- On 11/17/06 I removed the Market from the call
            SQLQuery = "SELECT shttCallLetters, shttCode, attOnAir, attOffAir, attDropDate, attMulticast, attCode"
            SQLQuery = SQLQuery + " FROM shtt , att"
            SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode"
            
'            xxx If chkActive.Value = vbChecked Then
'                SQLQuery = SQLQuery + " AND attDropDate >= " & "'" & smCurDate & "'"
'            End If
            
            
            SQLQuery = SQLQuery + " AND attVefCode = " & imVefCode & ")"
            If optSSSort(0).Value Then
                'Doug- On 11/17/06 I removed the Market from the call
                'SQLQuery = SQLQuery + " ORDER BY shttCallLetters, shttMarket"
                ''SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
                Set rst = gSQLSelectCall(SQLQuery)
                While Not rst.EOF
                    If DateValue(gAdjYear(rst!attDropDate)) < DateValue(gAdjYear(rst!attOffAir)) Then
                        sEndDate = Format$(rst!attDropDate, sgShowDateForm)
                    Else
                        sEndDate = Format$(rst!attOffAir, sgShowDateForm)
                    End If
                    'cboSSSort.AddItem Trim$(rst(0).Value) & ", " & Trim$(rst(1).Value) & "|" & rst(2).Value
                    If (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/1970")) Or (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/70")) Then   'Or (rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                        sRange = ""
                    Else
                        sRange = Format$(Trim$(rst!attOnAir), sgShowDateForm)
                    End If
                    'If (rst!attOffAir = "12/31/2069") Or (rst!attOffAir = "12/31/69") Then
                    If (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/2069")) Or (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/69")) Then 'Or (rst!attOffAir = "12/31/69") Then
                        If sRange <> "" Then
                            sRange = sRange & "-TFN"
                        End If
                    Else
                        If sRange <> "" Then
                            sRange = sRange & "-" & sEndDate    'rst!attOffAir
                        Else
                            sRange = "Thru " & sEndDate 'rst!attOffAir
                        End If
                    End If
                    'If IsNull(rst!shttMarket) = True Then
                    '    sMarket = ""
                    'Else
                    '    sMarket = rst!shttMarket  'Trim$(rst!shttMarket)
                    'End If
                    'cboSSSort.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket) & " " & sRange
                    
                    llStaCode = gBinarySearchStation(rst!shttCallLetters)
                    If llStaCode <> -1 Then
                        sMarket = Trim$(tgStationInfo(llStaCode).sMarket)
                    Else
                        sMarket = ""
                    End If
                    
                    If rst!attMulticast = "Y" Then
                        slMulticast = " Multicast"
                    Else
                        slMulticast = ""
                    End If
                    
                    ilAddSpot = True
                    If sMarket = "" Then
                        If chkActive.Value = vbChecked Then
                            If (DateValue(gAdjYear(sEndDate)) >= DateValue(smCurDate)) Then
                                cboSSSort.AddItem Trim$(rst!shttCallLetters) & " " & sRange & slMulticast
                            Else
                                ilAddSpot = False
                            End If
                        Else
                            cboSSSort.AddItem Trim$(rst!shttCallLetters) & " " & sRange & slMulticast
                        End If
                    Else
                        If chkActive.Value = vbChecked Then
                            If (DateValue(gAdjYear(sEndDate)) >= DateValue(smCurDate)) Then
                                cboSSSort.AddItem Trim$(rst!shttCallLetters) & ", " & sMarket & " " & sRange & slMulticast
                            Else
                                ilAddSpot = False
                            End If
                        Else
                            cboSSSort.AddItem Trim$(rst!shttCallLetters) & ", " & sMarket & " " & sRange & slMulticast
                        End If
                    End If
                    
                    
                    'Doug- On 11/17/06 I removed this Market code
                    'If IsNull(rst!mktName) = True Then
                    '    sMarket = ""
                    '    cboSSSort.AddItem Trim$(rst!shttCallLetters) & " " & sRange
                    'Else
                    '    sMarket = rst!mktName  'Trim$(rst!shttMarket)
                    '    cboSSSort.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket) & " " & sRange
                    'End If
                    If ilAddSpot Then
                        cboSSSort.ItemData(cboSSSort.NewIndex) = rst!attCode
                    End If
                    rst.MoveNext
                Wend
            Else
                'SQLQuery = SQLQuery + " ORDER BY shttMarket, shttCallLetters"
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
                Set rst = gSQLSelectCall(SQLQuery)
                While Not rst.EOF
                    If DateValue(gAdjYear(rst!attDropDate)) < DateValue(gAdjYear(rst!attOffAir)) Then
                        sEndDate = Format$(rst!attDropDate, sgShowDateForm)
                    Else
                        sEndDate = Format$(rst!attOffAir, sgShowDateForm)
                    End If
                    'cboSSSort.AddItem Trim$(rst(1).Value) & ", " & Trim$(rst(0).Value) & "|" & rst(2).Value
                    If (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/1970")) Or (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/70")) Then   'Or (rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                        sRange = ""
                    Else
                        sRange = Format$(Trim$(rst!attOnAir), sgShowDateForm)
                    End If
                    'If (rst!attOffAir = "12/31/2069") Or (rst!attOffAir = "12/31/69") Then
                    If (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/2069")) Or (DateValue(gAdjYear(sEndDate)) = DateValue("12/31/69")) Then 'Or (rst!attOffAir = "12/31/69") Then
                        If sRange <> "" Then
                            sRange = sRange & "-TFN"
                        End If
                    Else
                        If sRange <> "" Then
                            sRange = sRange & "-" & sEndDate    'rst!attOffAir
                        Else
                            sRange = "Thru " & sEndDate 'rst!attOffAir
                        End If
                    End If
                    'Doug- On 11/17/06 I replaced this call with the Binary match
                    ''If IsNull(rst!shttMarket) = True Then
                    ''    sMarket = ""
                    ''Else
                    ''    sMarket = rst!shttMarket  'Trim$(rst!shttMarket)
                    ''End If
                    ''cboSSSort.AddItem Trim$(sMarket) & ", " & Trim$(rst!shttCallLetters) & " " & sRange
                    'If IsNull(rst!mktName) = True Then
                    '    sMarket = ""
                    '    cboSSSort.AddItem Trim$(rst!shttCallLetters) & " " & sRange
                    'Else
                    '    sMarket = rst!mktName  'Trim$(rst!shttMarket)
                    '    cboSSSort.AddItem Trim$(sMarket) & ", " & Trim$(rst!shttCallLetters) & " " & sRange
                    'End If
                    llStaCode = gBinarySearchStation(rst!shttCallLetters)
                    If llStaCode <> -1 Then
                        sMarket = Trim$(tgStationInfo(llStaCode).sMarket)
                    Else
                        sMarket = ""
                    End If
                    
                    If rst!attMulticast = "Y" Then
                        slMulticast = " Multicast"
                    Else
                        slMulticast = ""
                    End If
                    
                    '10/22/19: Check if active only TTP 9470
                    If (chkActive.Value = vbUnchecked) Or ((chkActive.Value = vbChecked) And (DateValue(gAdjYear(sEndDate)) >= DateValue(smCurDate))) Then
                        If sMarket = "" Then
                            cboSSSort.AddItem Trim$(rst!shttCallLetters) & " " & sRange & slMulticast
                        Else
                            cboSSSort.AddItem Trim$(sMarket) & ", " & Trim$(rst!shttCallLetters) & " " & sRange & slMulticast
                        End If
                        cboSSSort.ItemData(cboSSSort.NewIndex) = rst!attCode
                    End If
                    rst.MoveNext
                Wend
            End If
            
        Else        'Get Affiliates without agreements
            ''mSort optSSSort(0), cboSSSort
            SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode"
            'SQLQuery = SQLQuery + " FROM shtt, att"
            'SQLQuery = "SELECT shttCallLetters, shttCode"
            'Doug- On 11/17/06 I removed the Left Outer Join
            'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, att"
            SQLQuery = SQLQuery + " FROM shtt, att"
            'SQLQuery = SQLQuery + " FROM shtt, att"
            SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode"
            SQLQuery = SQLQuery + " AND attVefCode = " & imVefCode & ")"
            'SQLQuery = SQLQuery + " ORDER BY shttCallLetters, shttMarket"
            SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
            Set rst = gSQLSelectCall(SQLQuery)
            ReDim imAffShttCode(0 To 0) As Integer
            While Not rst.EOF
                imAffShttCode(UBound(imAffShttCode)) = rst!shttCode   'rst(2).Value
                ReDim Preserve imAffShttCode(0 To UBound(imAffShttCode) + 1) As Integer
                rst.MoveNext
            Wend
            mSort optSSSort(0), cboSSSort, imAffShttCode()
        End If
    End If
    'If (optExAll(0).Value) Or (optPSSort(0).Value) Or (optPSSort(1).Value) Then
    '    If cboSSSort.Rows > 0 Then
    '        cboSSSort.MoveFirst
    '        cboSSSort.Text = cboSSSort.Columns(0).Text
    '        cboSSSort_Click
    '    End If
    'Else
    '    If cboSSSort.Rows > 1 Then
    '        cboSSSort.MoveFirst
    '        cboSSSort.MoveNext
    '        cboSSSort.Text = cboSSSort.Columns(0).Text
    '        cboSSSort_Click
    '    End If
    'End If
    'If cboSSSort.ListCount > 0 Theniu
    '    cboSSSort.ListIndex = 0
    'End If
    mShowTabs
    mMousePointer vbDefault
    Exit Sub
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-optExAll"
End Sub

Private Sub optExAll_GotFocus(Index As Integer)
    imIgnoreTabs = False
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub optFormerNCR_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optFormerNCR_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optNCR_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optNCR_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optPost_Click(Index As Integer)
    imFieldChgd = True
'    If optPost(2).Value = True Then
'        frcExport.Visible = True
'    Else
'        frcExport.Visible = False
'        frcLogType.Visible = False
'        frcPostType.Visible = False
'        lblWebPW.Visible = False
'        lblWebEmail.Visible = False
'        txtLogPassword.Visible = False
'        cmdGenPassword.Visible = False
'        txtEmailAddr.Visible = False
'        rbcExportType(0).Value = True
'    End If
End Sub

Private Sub optPost_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optPrintCP_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optPrintCP_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optPSSort_Click(Index As Integer)
    Dim iLoop As Integer
    
    If optPSSort(Index).Value = False Then
        Exit Sub
    End If
    mMousePointer vbHourglass
    DoEvents
    If sgUstWin(2) = "I" Then
        cmdRemap.Enabled = True
    End If
    ClearControls
    cboPSSort.Clear
    cboSSSort.Clear
    DoEvents
    imShttCode = 0
    imVefCode = 0
    mSetEventTitles
    
    If optPSSort(0).Value = True Then
        optSSSort(0).Visible = False
        optSSSort(1).Visible = False
        lblSort.Visible = False
        ReDim imAffShttCode(0 To 0) As Integer
        mSort optPSSort(0), cboPSSort, imAffShttCode()
    ElseIf optPSSort(1).Value = True Then
        optSSSort(0).Visible = False
        optSSSort(1).Visible = False
        lblSort.Visible = False
        ReDim imAffShttCode(0 To 0) As Integer
        mSort optPSSort(0), cboPSSort, imAffShttCode()
        
    ElseIf (optPSSort(2).Value = True) Or (optPSSort(3).Value = True) Then
        'Add sort by vehicle to primary sort
        optSSSort(0).Visible = True
        optSSSort(1).Visible = True
        lblSort.Visible = True
        optSSSort(0).Value = True
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'Bypass Wegener and OLA vehicles
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            'D.S. 09/12/02 Supports choice of active or all vehicles
            '6/21/12: check if vehicle should be added
            If mVehicleAddTest(iLoop) Then
                If optPSSort(2).Value = True Then
                    'Active Vehicles Only
                    If tgVehicleInfo(iLoop).sState = "A" Then
                        cboPSSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                        cboPSSort.ItemData(cboPSSort.NewIndex) = tgVehicleInfo(iLoop).iCode
                    End If
                Else
                    'All Vehicles Active and Dormant
                    cboPSSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                    cboPSSort.ItemData(cboPSSort.NewIndex) = tgVehicleInfo(iLoop).iCode
                End If
            End If
            'End If
        Next iLoop
        
        optSSSort(0).Visible = True
        optSSSort(1).Visible = True
        lblSort.Visible = True
    End If
    mShowTabs
    mMousePointer vbDefault
    Exit Sub
End Sub
    
Private Sub optPSSort_GotFocus(Index As Integer)
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub optRadarClearType_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optRadarClearType_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optSendTape_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optSendTape_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optSigned_Click(Index As Integer)
    imFieldChgd = True
    If optSigned(1).Value = True Then
        txtRetDate.Visible = True
    ElseIf optSigned(1).Value = False Then
        txtRetDate.Visible = False
    End If
End Sub

Private Sub optSigned_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub optSSSort_Click(Index As Integer)
    Dim iVehicle As Integer
    If optSSSort(Index).Value = False Then
        Exit Sub
    End If
    If optExAll(0).Value Then
        optExAll_Click 0
    Else
        optExAll_Click 1
    End If
    mShowTabs
    mMousePointer vbDefault
    Exit Sub
    
End Sub


Private Sub optSSSort_GotFocus(Index As Integer)
    imIgnoreTabs = True
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub optSuppressNotices_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optSuppressNotices_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub mTimeType(Index As Integer)
    'Index: 0=Daypart, 1=Avail
    Dim sMsg As String
    Dim ilRet As Integer
    Dim ilCol As Integer
    Dim slOnAir As String
    Dim slOffAir As String
    Dim slDrop As String
    Dim slCDStartTime As String
    

    If Not mLoadPledgeOk() Then
        Exit Sub
    End If

    'If imDatLoaded Then
        mMoveDaypart
        lgPledgeAttCode = lmAttCode
        igDayPartShttCode = imShttCode
        igDayPartVefCode = imVefCode
        igDayPartVefCombo = imVefCombo
        If edcNoAirPlays.Text <> "" Then
            igNoAirPlays = Val(edcNoAirPlays.Text)
            If cbcAirPlayNo.ListIndex > 0 Then
                igDefaultAirPlayNo = Val(cbcAirPlayNo.List(cbcAirPlayNo.ListIndex))
            Else
                igDefaultAirPlayNo = 1
            End If
        Else
            igNoAirPlays = 1
            igDefaultAirPlayNo = 1
        End If
        If Index = 1 Then  'Avails
            'grdDayparts.Groups(0).Caption = "Feed"
            chkMonthlyWebPost.Visible = True
            For ilCol = MONFDINDEX To ENDTIMEFDINDEX Step 1
                grdPledge.TextMatrix(0, ilCol) = "Feed"
            Next ilCol
            igAvails = True
            igCDTapeDayPart = False
            igLiveDayPart = False
            If imIgnoreTimeTypeChg Then
                Exit Sub
            End If
            igPledgeExist = False
            If (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "") Then
                igPledgeExist = True
            End If
            mGetDayParts
            frmAgmntPledgeSpec.Show vbModal
            If igAgmntReturn Then
                imDatLoaded = True
                mPopPledge
                sgAirPlay1TimeType = ""         '12-10-11
            End If
            'imDatLoaded = False
            'ReDim tgDat(0 To 0) As DAT
            'tgDat(0).sFdSTime = ""
            'mLoadPledge False, Index
        ElseIf Index = 0 Then  'Live Dayparts
            If chkMonthlyWebPost.Value = vbChecked Then
                gMsgBox "Monthly Posting is Checked and is Not Allowable using Dayparts.  Monthly will be Unchecked.", vbOKOnly, "Error"
                chkMonthlyWebPost.Value = vbUnchecked
                chkMonthlyWebPost.Visible = False
                Exit Sub
            End If
        
            'grdDayparts.Groups(0).Caption = "Feed"
            For ilCol = MONFDINDEX To ENDTIMEFDINDEX Step 1
                grdPledge.TextMatrix(0, ilCol) = "Feed"
            Next ilCol
            igAvails = False
            igCDTapeDayPart = False
            igLiveDayPart = True
            If imIgnoreTimeTypeChg Then
                Exit Sub
            End If
            igPledgeExist = False
            If (Trim$(grdPledge.TextMatrix(grdPledge.FixedRows, STATUSINDEX)) <> "") Then
                igPledgeExist = True
            End If
            mGetDayParts
            frmAgmntPledgeSpec.Show vbModal
            If igAgmntReturn Then
                imDatLoaded = True
                mPopPledge
                sgAirPlay1TimeType = ""         '12-10-11
            End If
            'imDatLoaded = False
            'ReDim tgDat(0 To 0) As DAT
            'tgDat(0).sFdSTime = ""
            'mLoadPledge False, Index
        End If
    'Else
    '    If imIgnoreTimeTypeChg Then
    '        Exit Sub
    '    End If
    '    mLoadPledge False, Index
    'End If
    slOnAir = Format("1/1/1970")
    If gIsDate(txtOnAirDate.Text) Then
        slOnAir = Format(txtOnAirDate.Text, sgShowDateForm)
    Else
        lacPrgTimes.Caption = "Program Times: "
        Exit Sub
    End If
    slOffAir = Format("12/31/2069", "m/d/yyyy")
    If gIsDate(txtOffAirDate.Text) Then
        slOffAir = Format(txtOffAirDate.Text, sgShowDateForm)
    End If
    slDrop = Format("12/31/2069", "m/d/yyyy")
    If gIsDate(txtDropDate.Text) Then
        slDrop = Format(txtDropDate.Text, sgShowDateForm)
    End If
    slCDStartTime = ""
    ilRet = gDetermineAgreementTimes(imShttCode, imVefCode, slOnAir, slOffAir, slDrop, slCDStartTime, sgVehProgStartTime, sgVehProgEndTime)
    lacPrgTimes.Caption = ""
    If sgVehProgStartTime <> "" Then
        sgVehProgStartTime = gCompactTime(sgVehProgStartTime)
        lacPrgTimes.Caption = "Program Times: " & sgVehProgStartTime
        If sgVehProgEndTime <> "" Then
            sgVehProgEndTime = gCompactTime(sgVehProgEndTime)
            lacPrgTimes.Caption = lacPrgTimes.Caption & "-" & sgVehProgEndTime
        End If
    End If
End Sub

Private Sub optVoiceTracked_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optVoiceTracked_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    mETSetShow
    mPledgeSetShow
End Sub

Private Sub pbcDay_Click()
    If ckcDay.Value = vbChecked Then
        ckcDay.Value = vbUnchecked
    Else
        ckcDay.Value = vbChecked
    End If
End Sub

Private Sub pbcDayFed_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        grdPledge.Text = "A"
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcDayFed_Paint
    ElseIf KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
        grdPledge.Text = "B"
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcDayFed_Paint
    End If
    If KeyAscii = Asc(" ") Then
        slStr = grdPledge.Text
        If slStr = "A" Then
            grdPledge.Text = "B"
        ElseIf slStr = "B" Then
            grdPledge.Text = "A"
        Else
            grdPledge.Text = "A"
        End If
        imFieldChgd = True
        bmPledgeDataChgd = True
        pbcDayFed_Paint
    End If
End Sub

Private Sub pbcDayFed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    slStr = grdPledge.Text
    If slStr = "A" Then
        grdPledge.Text = "B"
    ElseIf slStr = "B" Then
        grdPledge.Text = "A"
    Else
        grdPledge.Text = "A"
    End If
    imFieldChgd = True
    bmPledgeDataChgd = True
    pbcDayFed_Paint

End Sub

Private Sub pbcDayFed_Paint()
    pbcDayFed.Cls
    pbcDayFed.CurrentX = 15
    pbcDayFed.CurrentY = 0 'fgBoxInsetY
    If grdPledge.Text = "B" Then
        pbcDayFed.Print "Before"
    ElseIf grdPledge.Text = "A" Then
        pbcDayFed.Print "After"
    Else
        pbcDayFed.Print ""
    End If
End Sub

Private Sub pbcETSTab_GotFocus()
    Dim ilVisible As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim slStr As String
    Dim llETTime As Long
    
    If GetFocus() <> pbcETSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    ilVisible = False
    If txtET.Visible Then
        ilVisible = True
    ElseIf cbcETDay.Visible Then
        ilVisible = True
    End If
    If ilVisible Then
        llEnableRow = lmETEnableRow
        llEnableCol = lmETEnableCol
        mETSetShow
        grdET.Row = llEnableRow
        grdET.Col = llEnableCol
        Do
            Select Case grdET.Col
                Case ETDAYINDEX
                    slStr = Trim$(txtET.Text)
                    If slStr <> "" Then
                        If Not (gIsTime(slStr)) Then
                            Beep
                            grdET.Col = llEnableCol
                            grdET.Row = llEnableRow
                            mETEnableBox
                            Exit Sub
                        End If
                    End If
                    If grdET.Row = grdET.FixedRows Then
                        mETSetShow
                        pbcPledgeSTab_GotFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdET.Row = grdET.Row - 1
                    If Not grdET.RowIsVisible(grdET.Row) Then
                        grdET.TopRow = grdET.TopRow - 1
                    End If
                    grdET.Col = ETTIMEINDEX
                Case ETTIMEINDEX
                    slStr = Trim$(txtET.Text)
                    If (slStr = "") Then
                        Beep
                        grdET.Col = llEnableCol
                        grdET.Row = llEnableRow
                        mETEnableBox
                        Exit Sub
                    End If
                    If Not (gIsTime(slStr)) Then
                        Beep
                        grdET.Col = llEnableCol
                        grdET.Row = llEnableRow
                        mETEnableBox
                        Exit Sub
                    End If
                    If Trim$(grdET.TextMatrix(llEnableRow, ETTIMEINDEX)) <> "" Then
                        llETTime = gTimeToLong(grdET.TextMatrix(llEnableRow, ETTIMEINDEX), False)
                        If (llETTime < lmPdSTime) Or (llETTime >= lmPdETime) Then
                            Beep
                            grdET.Col = llEnableCol
                            grdET.Row = llEnableRow
                            mETEnableBox
                            Exit Sub
                        End If
                    End If
                    grdET.Col = grdET.Col - 1
                Case Else
                    grdET.Col = grdET.Col - 1
            End Select
        Loop While Not mETColOk()
        mETEnableBox
    Else
        lmTopRow = -1
        If UBound(tmETAvailInfo) > LBound(tmETAvailInfo) Then
            grdET.TopRow = grdET.FixedRows
            grdET.Col = ETDAYINDEX
            grdET.Row = grdET.FixedRows
            mETEnableBox
        Else
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub pbcETTab_GotFocus()
    Dim ilVisible As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim slStr As String
    Dim llETTime As Long
    
    If GetFocus() <> pbcETTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    ilVisible = False
    If txtET.Visible Then
        ilVisible = True
    ElseIf cbcETDay.Visible Then
        ilVisible = True
    End If
    If ilVisible Then
        llEnableRow = lmETEnableRow
        llEnableCol = lmETEnableCol
        mETSetShow
        grdET.Row = llEnableRow
        grdET.Col = llEnableCol
        Do
            Select Case grdET.Col
                '12/8/14: If day blank exit
                Case ETDAYINDEX
                    If Trim$(grdET.TextMatrix(llEnableRow, ETDAYINDEX)) = "" Then
                        pbcPledgeTab.SetFocus
                        Exit Sub
                    End If
                    grdET.Col = grdET.Col + 1
                Case ETTIMEINDEX
                    slStr = Trim$(txtET.Text)
                    '9/4/14: Allow Time to be removed (To remove record, Day and time must be blanked out)
                    'If (slStr = "") Then
                    '    pbcPledgeTab.SetFocus
                    '    Exit Sub
                    'End If
                    If slStr <> "" Then
                        If Not (gIsTime(slStr)) Then
                            Beep
                            grdET.Col = llEnableCol
                            grdET.Row = llEnableRow
                            mETEnableBox
                            Exit Sub
                        End If
                        If Trim$(grdET.TextMatrix(llEnableRow, ETTIMEINDEX)) <> "" Then
                            llETTime = gTimeToLong(grdET.TextMatrix(llEnableRow, ETTIMEINDEX), False)
                            If (llETTime < lmPdSTime) Or (llETTime >= lmPdETime) Then
                                Beep
                                grdET.Col = llEnableCol
                                grdET.Row = llEnableRow
                                mETEnableBox
                                Exit Sub
                            End If
                        End If
                    End If
                    If (grdET.Row + 1 >= grdET.Rows) Then
                        'grdET.Rows = grdET.Rows + 1
                        'grdET.Row = grdET.Row + 1
                        'grdET.TextMatrix(grdET.Row, ETEPTCODEINDEX) = 0
                        'If Not grdET.RowIsVisible(grdET.Row) Then
                        '    grdET.TopRow = grdET.TopRow + 1
                        'End If
                        'grdET.Col = ETDAYINDEX
                        'mETEnableBox
                        pbcPledgeTab.SetFocus
                        Exit Sub
                    End If
                    grdET.Row = grdET.Row + 1
                    grdET.Col = ETDAYINDEX
                    'If Not grdET.RowIsVisible(grdET.Row) Then
                    If grdET.Top + grdET.RowPos(grdET.Row) + 15 + txtET.Height > grdET.Top + grdET.Height Then
                        grdET.TopRow = grdET.TopRow + 1
                    End If
                Case Else
                    grdET.Col = grdET.Col + 1
            End Select
        Loop While Not mETColOk()
        mETEnableBox
    Else
        lmTopRow = -1
        grdET.TopRow = grdET.FixedRows
        grdET.Col = ETTIMEINDEX
        grdET.Row = grdET.FixedRows
        'grdET_Click
        mETEnableBox
    End If
End Sub

Private Sub pbcPledgeSTab_GotFocus()
    Dim ilVisible As Integer
    Dim ilNext As Integer
        
    If GetFocus() <> pbcPledgeSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        'grdPledge.SetFocus
        mPledgeEnableBox
        Exit Sub
    End If
    ilVisible = False
    If pbcDay.Visible Then
        ilVisible = True
    ElseIf pbcDayFed.Visible Then
        ilVisible = True
    ElseIf pbcEmbeddedOrROS.Visible Then
        ilVisible = True
    ElseIf txtTime.Visible Then
        ilVisible = True
    ElseIf txtDropdown.Visible Then
        ilVisible = True
    ElseIf txtAirPlay.Visible Then
        ilVisible = True
    ElseIf frcET.Visible Then
        ilVisible = True
    End If
    If ilVisible Then
        mPledgeSetShow
        Do
            ilNext = False
            If grdPledge.Col = MONFDINDEX Then
                If grdPledge.Row > grdPledge.FixedRows Then
                    lmTopRow = -1
                    grdPledge.Row = grdPledge.Row - 1
                    If Not grdPledge.RowIsVisible(grdPledge.Row) Then
                        grdPledge.TopRow = grdPledge.TopRow - 1
                    End If
                    grdPledge.Col = MONFDINDEX
                    If Not mPledgeColAllowed(grdPledge.Col) Then
                        grdPledge.Col = STATUSINDEX
                    End If
                    'grdPledge.SetFocus
                    'mPledgeEnableBox
                Else
                    If pbcClickFocus.Enabled Then
                        pbcClickFocus.SetFocus
                    Else
                        cmdCancel.SetFocus
                    End If
                    Exit Sub
                End If
            ElseIf (grdPledge.Col = STARTTIMEPDINDEX) Then
                If mPdPriorFd(grdPledge.Row) Then
                    grdPledge.Col = grdPledge.Col - 1
                Else
                    grdPledge.Col = grdPledge.Col - 2
                End If
                'grdPledge.SetFocus
                'mPledgeEnableBox
            Else
                grdPledge.Col = grdPledge.Col - 1
                'grdPledge.SetFocus
                'mPledgeEnableBox
            End If
            If mPledgeColAllowed(grdPledge.Col) Then
                mPledgeEnableBox
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdPledge.TopRow = grdPledge.FixedRows
        grdPledge.Col = MONFDINDEX
        grdPledge.Row = grdPledge.FixedRows
        Do
            If mPledgeColAllowed(grdPledge.Col) Then
                Exit Do
            End If
            If grdPledge.Col + 1 >= ENDTIMEPDINDEX Then
                cmdCancel.SetFocus
                Exit Sub
            End If
            grdPledge.Col = grdPledge.Col + 1
        Loop
            
        'grdPledge_Click
        mPledgeEnableBox
    End If
End Sub

Private Sub pbcPledgeTab_GotFocus()
    Dim ilVisible As Integer
    Dim ilBypassCol As Integer
    Dim llCol As Long
    Dim ilDay As Integer
    Dim slStr As String
    Dim llRowIndex As Long
    Dim llRow As Long
    
    If GetFocus() <> pbcPledgeTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(2) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If sgUstPledge <> "Y" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    ilVisible = False
    If pbcDay.Visible Then
        ilVisible = True
    ElseIf txtTime.Visible Then
        ilVisible = True
    ElseIf txtDropdown.Visible Then
        ilVisible = True
    ElseIf pbcDayFed.Visible Then
        ilVisible = True
    ElseIf pbcEmbeddedOrROS.Visible Then
        ilVisible = True
    ElseIf txtAirPlay.Visible Then
        ilVisible = True
    ElseIf frcET.Visible Then
        ilVisible = True
    End If
    If ilVisible Then
        mPledgeSetShow
        If (grdPledge.Col >= ENDTIMEPDINDEX) Or (grdPledge.Col = STATUSINDEX) Or (grdPledge.Col = AIRPLAYINDEX) Then
            ilBypassCol = True
            If grdPledge.Col = STATUSINDEX Then
                slStr = Trim$(txtDropdown.Text)
                If slStr = "" Then
                    Beep
                    grdPledge.Col = grdPledge.Col
                    'grdPledge.SetFocus
                    mPledgeEnableBox
                    Exit Sub
                End If
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex < 0 Then
                    Beep
                    grdPledge.Col = grdPledge.Col
                    'grdPledge.SetFocus
                    mPledgeEnableBox
                    Exit Sub
                End If
                ilBypassCol = Not mPledgeColAllowed(grdPledge.Col + 1)
                If ilBypassCol Then
                    '8/20/11: Air p[lay column could be yellow if only one air play specified
                    '         Test day column
                    ilBypassCol = Not mPledgeColAllowed(grdPledge.Col + 2)
                    If Not ilBypassCol Then
                        grdPledge.Col = grdPledge.Col + 1
                    End If
                End If
            ElseIf grdPledge.Col = AIRPLAYINDEX Then
                ilBypassCol = Not mPledgeColAllowed(grdPledge.Col + 1)
            ElseIf grdPledge.Col = ENDTIMEPDINDEX Then
                ilBypassCol = Not mPledgeColAllowed(grdPledge.Col + 1)
            End If
            If ilBypassCol Then
'                If grdPledge.Row + 1 < grdPledge.Rows Then
                llRow = grdPledge.Rows
                Do
                    llRow = llRow - 1
                Loop While grdPledge.TextMatrix(llRow, STATUSINDEX) = ""
                Do
                    llRow = llRow + 1
                    If (grdPledge.Row + 1 < llRow) Then
                        lmTopRow = -1
                        grdPledge.Row = grdPledge.Row + 1
                        If Not grdPledge.RowIsVisible(grdPledge.Row) Then
                            grdPledge.TopRow = grdPledge.TopRow + 1
                        End If
                        If grdPledge.Col = ESTIMATETIMEINDEX Then
                            If (Trim$(grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)) <> "") And (mPledgeColAllowed(ESTIMATETIMEINDEX)) Then
                                mPledgeEnableBox
                                Exit Do
                            End If
                        End If
                        If Trim$(grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)) <> "" Then
                            'If grdPledge.Enabled Then
                            '    grdPledge.SetFocus
                            'End If
                            For llCol = MONFDINDEX To ESTIMATETIMEINDEX Step 1
                                If mPledgeColAllowed(llCol) Then
                                    grdPledge.Col = llCol
                                    mPledgeEnableBox
                                    Exit Do
                                End If
                            Next llCol
                            If pbcClickFocus.Enabled Then
                                pbcClickFocus.SetFocus
                            Else
                                cmdCancel.SetFocus
                            End If
                        Else
                            imFromArrow = True
                            pbcArrow.Move grdPledge.Left - pbcArrow.Width, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + (grdPledge.RowHeight(grdPledge.Row) - pbcArrow.Height) / 2
                            pbcArrow.Visible = True
                            If pbcArrow.Enabled Then
                                pbcArrow.SetFocus
                            End If
                            Exit Do
                        End If
                    Else
                        If grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX) <> "" Then
                            lmTopRow = -1
                            If grdPledge.Row + 1 >= grdPledge.Rows Then
                                grdPledge.AddItem ""
                            End If
                            grdPledge.Row = grdPledge.Row + 1
                            If Not grdPledge.RowIsVisible(grdPledge.Row) Then
                                grdPledge.TopRow = grdPledge.TopRow + 1
                            End If
                            For llCol = MONFDINDEX To SUNFDINDEX Step 1
                                grdPledge.Col = llCol
                                grdPledge.CellFontName = "Monotype Sorts"
                            Next llCol
                            For llCol = MONPDINDEX To SUNPDINDEX Step 1
                                grdPledge.Col = llCol
                                grdPledge.CellFontName = "Monotype Sorts"
                            Next llCol
                            grdPledge.Col = MONFDINDEX
                            'grdPledge.SetFocus
                            imFromArrow = True
                            pbcArrow.Move grdPledge.Left - pbcArrow.Width, grdPledge.Top + grdPledge.RowPos(grdPledge.Row) + (grdPledge.RowHeight(grdPledge.Row) - pbcArrow.Height) / 2
                            pbcArrow.Visible = True
                            If pbcArrow.Enabled Then
                                pbcArrow.SetFocus
                            End If
                            Exit Do
                        Else
                            If pbcClickFocus.Enabled Then
                                pbcClickFocus.SetFocus
                            End If
                            Exit Do
                        End If
                    End If
                Loop
            Else
                grdPledge.Col = grdPledge.Col + 1
                'grdPledge.SetFocus
                mPledgeEnableBox
            End If
        ElseIf (grdPledge.Col = STARTTIMEFDINDEX) Or (grdPledge.Col = ENDTIMEFDINDEX) Or (grdPledge.Col = STARTTIMEPDINDEX) Or (grdPledge.Col = ENDTIMEPDINDEX) Then
            slStr = Trim$(txtTime.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                grdPledge.Col = grdPledge.Col + 1
                'grdPledge.SetFocus
                mPledgeEnableBox
            Else
                Beep
                grdPledge.Col = grdPledge.Col
                'grdPledge.SetFocus
                mPledgeEnableBox
            End If
        ElseIf (grdPledge.Col = SUNFDINDEX) Then
            For ilDay = MONFDINDEX To SUNFDINDEX Step 1
                If Trim$(grdPledge.TextMatrix(grdPledge.Row, ilDay)) <> "" Then
                    grdPledge.Col = grdPledge.Col + 1
                    'grdPledge.SetFocus
                    mPledgeEnableBox
                    Exit Sub
                End If
            Next ilDay
            Beep
            grdPledge.Col = grdPledge.Col
            'grdPledge.SetFocus
            mPledgeEnableBox
        ElseIf (grdPledge.Col = SUNPDINDEX) Then
            For ilDay = MONPDINDEX To SUNPDINDEX Step 1
                If Trim$(grdPledge.TextMatrix(grdPledge.Row, ilDay)) <> "" Then
                    If mPdPriorFd(grdPledge.Row) Then
                        grdPledge.Col = grdPledge.Col + 1
                    Else
                        grdPledge.Col = grdPledge.Col + 2
                    End If
                    'grdPledge.SetFocus
                    mPledgeEnableBox
                    Exit Sub
                End If
            Next ilDay
            Beep
            grdPledge.Col = grdPledge.Col
            'grdPledge.SetFocus
            mPledgeEnableBox
        Else
            grdPledge.Col = grdPledge.Col + 1
            'grdPledge.SetFocus
            mPledgeEnableBox
        End If
    Else
        lmTopRow = -1
        grdPledge.TopRow = grdPledge.FixedRows
        grdPledge.Col = MONFDINDEX
        grdPledge.Row = grdPledge.FixedRows
        'grdPledge_Click
        Do
        If mPledgeColAllowed(grdPledge.Col) Then
                Exit Do
            End If
            If grdPledge.Col + 1 >= ENDTIMEPDINDEX Then
                cmdCancel.SetFocus
                Exit Sub
            End If
            grdPledge.Col = grdPledge.Col + 1
        Loop
        mPledgeEnableBox
    End If
End Sub

Private Sub pbcSTab_GotFocus(Index As Integer)
    If imIgnoreTabs Then
        imIgnoreTabs = False
        If frcTab(0).Visible Then
            If txtStartDate.Visible And txtStartDate.GetEnabled Then
                txtStartDate.SetFocus
            End If
        ElseIf frcTab(1).Visible Then
            If rbcExportType(0).Value Then
                If rbcExportType(0).Visible And rbcExportType(0).Enabled Then
                    rbcExportType(0).SetFocus
                End If
            ElseIf rbcExportType(1).Value Then
                If rbcExportType(1).Visible And rbcExportType(1).Enabled Then
                    rbcExportType(1).SetFocus
                End If
            Else
                If cmdCancel.Visible And cmdCancel.Enabled Then
                    cmdCancel.SetFocus
                End If
            End If
        ElseIf frcTab(2).Visible Then
            'grdPledge.SetFocus
            mPledgeEnableBox
        ElseIf frcTab(3).Visible Then
        
            If optCarryCmml(0).Value Then
                If optCarryCmml(0).Visible And optCarryCmml(0).Enabled Then
                    optCarryCmml(0).SetFocus
                End If
            ElseIf optCarryCmml(1).Value Then
                If optCarryCmml(1).Visible And optCarryCmml(1).Enabled Then
                    optCarryCmml(1).SetFocus
                End If
            Else
                If cmdCancel.Visible And cmdCancel.Enabled Then
                    cmdCancel.SetFocus
                End If
            End If
        
        End If
        Exit Sub
    End If
    mPledgeSetShow
    If Index = 0 Then
        'gSendKeys "%P", True
        'cmdAffRep.SetFocus
        If cmdSave.Enabled Then
            If cmdSave.Enabled Then
                cmdSave.SetFocus
            End If
        ElseIf cmdNew.Enabled Then
            If cmdNew.Enabled Then
                cmdNew.SetFocus
            End If
        End If
    ElseIf Index = 1 Then
        'gSendKeys "%M", True
        TabStrip1.Tabs(TABMAIN).Selected = True
        'If optPrintCP(0).Value Then
        '    If optPrintCP(0).Enabled Then
        '        optPrintCP(0).SetFocus
        '    End If
        'ElseIf optPrintCP(1).Value Then
        '    If optPrintCP(1).Enabled Then
        '        optPrintCP(1).SetFocus
        '    End If
        'Else
            If txtDropDate.Visible And txtDropDate.GetEnabled Then
                txtDropDate.SetFocus
            End If
        'End If
    ElseIf Index = 2 Then
        'gSendKeys "%D", True
        TabStrip1.Tabs(TABDELIVERY).Selected = True
        If rbcExportType(0).Value Then
            If rbcExportType(0).Enabled Then
                rbcExportType(0).SetFocus
            End If
        ElseIf rbcExportType(1).Value Then
            If rbcExportType(1).Enabled Then
                rbcExportType(1).SetFocus
            End If
        'ElseIf rbcExportType(2).Value Then
        '    If rbcExportType(2).Enabled Then
        '        rbcExportType(2).SetFocus
        '    End If
        'ElseIf rbcExportType(3).Value Then
        '    If rbcExportType(3).Enabled Then
        '        rbcExportType(3).SetFocus
        '    End If
        Else
            If cmdCancel.Enabled Then
                cmdCancel.SetFocus
            End If
        End If
    'ElseIf Index = 3 Then
    '    'gSendKeys "%L", True
    '    'txtOther.SetFocus
    End If
End Sub

Private Sub pbcTab_GotFocus(Index As Integer)
    If imIgnoreTabs Then
        imIgnoreTabs = False
        If frcTab(0).Visible Then
            If txtDropDate.Visible And txtDropDate.GetEnabled Then
                txtDropDate.SetFocus
            End If
        ElseIf frcTab(1).Visible Then
            If rbcExportType(0).Value Then
                If rbcExportType(0).Visible And rbcExportType(0).Enabled Then
                    rbcExportType(0).SetFocus
                End If
            ElseIf rbcExportType(1).Value Then
                If rbcExportType(1).Visible And rbcExportType(1).Enabled Then
                    rbcExportType(1).SetFocus
                End If
            Else
                If cmdCancel.Visible And cmdCancel.Enabled Then
                    cmdCancel.SetFocus
                End If
            End If
        ElseIf frcTab(2).Visible Then
            'grdDayparts.SetFocus
            If pbcPledgeSTab.Enabled Then
                pbcPledgeSTab.SetFocus
            End If
        ElseIf frcTab(3).Visible Then
            If optCarryCmml(0).Value Then
                If optCarryCmml(0).Enabled Then
                    optCarryCmml(0).SetFocus
                End If
            ElseIf optCarryCmml(1).Value Then
                If optCarryCmml(1).Enabled Then
                    optCarryCmml(1).SetFocus
                End If
            Else
                If cmdCancel.Enabled Then
                    cmdCancel.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If
    mPledgeSetShow
    If Index = 0 Then
        'gSendKeys "%D", True
        TabStrip1.Tabs(TABDELIVERY).Selected = True
        'If optCarryCmml(0).Value Then
        '    If optCarryCmml(0).Enabled Then
        '        optCarryCmml(0).SetFocus
        '    End If
        'ElseIf optCarryCmml(1).Value Then
        '    If optCarryCmml(1).Enabled Then
        '        optCarryCmml(1).SetFocus
        '    End If
        'Else
        '    If cmdCancel.Enabled Then
        '        cmdCancel.SetFocus
        '    End If
        'End If
    ElseIf Index = 1 Then
        'gSendKeys "%I", True
        'If optTimeType(0).Value Then
        '    If optTimeType(0).Enabled Then
        '        optTimeType(0).SetFocus
        '    End If
        'ElseIf optTimeType(1).Value Then
        '    If optTimeType(1).Enabled Then
        '        optTimeType(1).SetFocus
        '    End If
        'Else
        '    'grdDayparts.SetFocus
        '    If pbcPledgeSTab.Enabled Then
        '        pbcPledgeSTab.SetFocus
        '    End If
        'End If
        'gSendKeys "%P", True
        TabStrip1.Tabs(TABPLEDGE).Selected = True
        'cbcAirPlayNo.SetFocus
    ElseIf Index = 2 Then
        'gSendKeys "%L", True
        'If optPrintCP(0).Value Then
        '    optPrintCP(0).SetFocus
        'ElseIf optPrintCP(1).Value Then
        '    optPrintCP(1).SetFocus
        'Else
        '    txtLog.SetFocus
        'End If
        'gSendKeys "%I", True
        TabStrip1.Tabs(TABINTERFACE).Selected = True
        'If cmdSave.Enabled Then
        '    If cmdSave.Enabled Then
        '        cmdSave.SetFocus
        '    End If
        'ElseIf cmdNew.Enabled Then
        '    If cmdNew.Enabled Then
        '        cmdNew.SetFocus
        '    End If
        'End If
    'ElseIf Index = 3 Then
    '    gSendKeys "%I", True
    End If
End Sub

Private Sub rbcExportType_Click(Index As Integer)
    '7701
    imFieldChgd = True
    If rbcExportType(0).Value = True Then
        ckcExportTo(0).Value = vbUnchecked
        frcPosting.Visible = True
        'frcAudioDelivery.Visible = False
        frcLogDelivery.Visible = False
    Else
        If gUsingWeb Then
            ckcExportTo(0).Value = vbChecked
        End If
        frcPosting.Visible = False
        'frcAudioDelivery.Visible = True
        frcLogDelivery.Visible = True
    End If
    mVerifyPWExists



'    imFieldChgd = True
'    If rbcExportType(0).Value = True Then
'        frcPosting.Visible = True
'        ckcExportTo(0).Value = vbUnchecked
'        ckcExportTo(1).Value = vbUnchecked
'        ckcExportTo(2).Value = vbUnchecked
'        ckcExportTo(3).Value = vbUnchecked
'        ckcExportTo(0).Enabled = False
'        ckcExportTo(1).Enabled = False
'        ckcExportTo(2).Enabled = False
'        ckcExportTo(3).Enabled = False
'        '6592- CBS
'        ckcExportTo(4).Value = vbUnchecked
'        ckcExportTo(4).Enabled = False
'
'        ckcExportTo(6).Value = vbUnchecked
'        ckcExportTo(6).Enabled = False
'    Else
'        frcPosting.Visible = False
'        If gUsingWeb Then
'            ckcExportTo(0).Value = vbChecked
'            ckcExportTo(0).Enabled = True
'            ckcExportTo(1).Enabled = True
'            '6592
'            ckcExportTo(4).Enabled = True
'        Else
'            ckcExportTo(0).Value = vbUnchecked
'            ckcExportTo(1).Enabled = False
'        End If
'        If gUsingUnivision Then
'            ckcExportTo(2).Enabled = True
'        End If
'        ckcExportTo(3).Enabled = True
'        ckcExportTo(6).Enabled = True
'    End If
'    'If ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) And gUsingWeb Then
'    If ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) And gUsingWeb Then
'        '5/15/11: Remove PostType and LogType
'        'frcLogType.Visible = True
'        'frcPostType.Visible = True
'        frcSendLogEMail.Visible = True
'        'lblWebPW.Visible = True
'        'lblWebEmail.Visible = True
'        'txtLogPassword.Visible = True
'        'cmdGenPassword.Visible = True
'        'txtEmailAddr.Visible = True
'        'mVerifyPWExists
'    Else
'        frcLogType.Visible = False
'        frcPostType.Visible = False
'        frcSendLogEMail.Visible = False
'        'lblWebPW.Visible = False
'        'lblWebEmail.Visible = False
'        'txtLogPassword.Visible = False
'        'cmdGenPassword.Visible = False
'        'txtEmailAddr.Visible = False
'    End If
'
''    If rbcExportType(1).Value = True And gUsingWeb Then
'        mVerifyPWExists
''    End If
End Sub

Private Sub rbcExportType_GotFocus(Index As Integer)
    imIgnoreTabs = True
End Sub

Private Sub rbcLogType_Click(Index As Integer)
    If frcLogType.Visible = True Then
        imFieldChgd = True
        If rbcLogType(1).Value Then
            If rbcPostType(2).Value Then
                rbcPostType(0).Value = True
            End If
            rbcPostType(2).Enabled = False
        ElseIf rbcLogType(1).Value = False Then
            rbcPostType(2).Enabled = True
        End If
    End If
End Sub

Private Sub rbcLogType_GotFocus(Index As Integer)
    imIgnoreTabs = True
End Sub

Private Sub rbcMulticast_Click(Index As Integer)

'    If rbcExportType(0).Value And Index <> 0 Then
'        gMsgBox "CP Control must be Network Web Site First"
'        rbcMulticast(1).Value = False
'    End If
End Sub

Private Sub rbcMulticast_GotFocus(Index As Integer)
    If Index = 0 Then
        'If (Not rbcExportType(1).Value) And (Not rbcExportType(3).Value) Then
        If ckcExportTo(0).Value = vbUnchecked Then
            gMsgBox "CP Control must be set to Network Web Site first."
        End If
        rbcMulticast(1).Value = True
    End If
End Sub

Private Sub rbcPostType_Click(Index As Integer)
    If frcPostType.Visible = True Then
        imFieldChgd = True
    End If
End Sub

Private Sub rbcPostType_GotFocus(Index As Integer)
    imIgnoreTabs = True
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Dim iRet As Integer
    
    If (imTabIndex = TABPLEDGE) And (Not bmBypassTestDP) Then
        iRet = mTestDaypart(False)
        If Not iRet Then
            Cancel = True
            imIgnoreTabClick = True
        End If
    End If
    bmBypassTestDP = False
End Sub

Private Sub TabStrip1_Click()
    If imIgnoreTabClick Then
        imIgnoreTabClick = False
        Exit Sub
    End If
    If imTabIndex = TabStrip1.SelectedItem.Index Then
        Exit Sub
    End If
    mETSetShow
    mPledgeSetShow
'    If TabStrip1.SelectedItem.Index = 3 Then
'        optTimeType(0).Value = False
'        optTimeType(1).Value = False
'        optTimeType(2).Value = False
'    End If
    frcTab(0).Visible = False
    frcTab(1).Visible = False
    frcTab(2).Visible = False
    frcTab(3).Visible = False
    frcTab(4).Visible = False
    frcEvent.Visible = False
    grdMulticast.Visible = False
    udcContactGrid.Action 1
    'frcTab(TabStrip1.SelectedItem.Index - 1).Visible = True
    Select Case TabStrip1.SelectedItem.Index
        Case TABMAIN  'Main
            frcTab(0).Visible = True
        Case TABPERSONNEL  'Personnel
            frcTab(4).Visible = True
        Case TABPLEDGE  'Pledge
            If Not mLoadPledgeOk() Then
                imTabIndex = -1
                'TabStrip1.SetFocus
                'gSendKeys "%M", True
                TabStrip1.Tabs(TABMAIN).Selected = True
                'txtOnAirDate.SetFocus
                pbcClickFocus.SetFocus
                Exit Sub
            End If
            mPopMulticast
            If smPledgeByEvent <> "Y" Then
                frcTab(2).Visible = True
            Else
                frcEvent.Visible = True
            End If
        Case TABDELIVERY  'Delivery
            frcTab(1).Visible = True
        Case TABINTERFACE  'Interface
            frcTab(3).Visible = True
    End Select
    mSetMulticast
    'frcTab(imTabIndex - 1).Visible = False
    imTabIndex = TabStrip1.SelectedItem.Index
    If (imTabIndex = TABPLEDGE) And (Not imDatLoaded) Then
        If IsAgmntDirty = True Then
            mLoadPledge True, -1
        Else
            'If (optTimeType(0).Value) Or (optTimeType(1).Value) Then
            '    mLoadPledge False
            'End If
            If smPledgeByEvent = "Y" Then
                mLoadPledge True, -1
            End If
        End If
    End If
    If imTabIndex = TABPLEDGE Then
        grdPledge.Redraw = True
    End If
    imIgnoreTabs = True
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mPopulateEvents
End Sub

Private Sub txtACName_Change()
    imFieldChgd = True
End Sub

Private Sub txtACName_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtACPhone_Change()
    imFieldChgd = True
End Sub

Private Sub txtACPhone_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtActStartTime_LostFocus()
'    Dim slStr As String
'    Dim slPdTime As String
'    Dim llLibStartTime As Long
'    Dim llPdTime As Long
'    Dim llTimeOffset As Long
'    Dim llRow As Long
'    Dim llRowIndex As Long
'    Dim iIndex As Integer
'
'    slStr = Trim$(txtActStartTime.Text)
'    slPdTime = sgCDStartTime
'    If (gIsTime(slStr)) And (slStr <> "") And (gIsTime(slPdTime)) And (slPdTime <> "") Then
'        llLibStartTime = Hour(slPdTime) * 3600 + 60 * Minute(slPdTime)
'        llPdTime = Hour(slStr) * 3600 + 60 * Minute(slStr)
'        If llLibStartTime <> llPdTime Then
'            llTimeOffset = mGetTimeOffSet(llLibStartTime, llPdTime)
'            For llRow = grdPledge.FixedRows To grdPledge.Rows - grdPledge.FixedRows Step 1
'                slStr = grdPledge.TextMatrix(llRow, STATUSINDEX)
'                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
'                If llRowIndex >= 0 Then
'                    iIndex = lbcStatus.ItemData(llRowIndex)
'                    If tmStatusTypes(iIndex).iPledged < 2 Then
'                        llPdTime = gTimeToCurrency(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX), False)
'                        llPdTime = llPdTime + llTimeOffset
'                        grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = gLongToTime(llPdTime)
'                        llPdTime = gTimeToCurrency(grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX), False)
'                        llPdTime = llPdTime + llTimeOffset
'                        grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = gLongToTime(llPdTime)
'                        grdPledge.TextMatrix(llRow, STATUSINDEX) = "2-Air Delay B'cast"  '"2-Air In Daypart"
'                    End If
'                End If
'            Next llRow
'        End If
'        sgCDStartTime = txtActStartTime.Text
'    End If
End Sub

Private Sub txtComments_Change()
    imFieldChgd = True
End Sub

Private Sub txtComments_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtCP_Change()
    imFieldChgd = True
End Sub

Private Sub txtCP_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDays_Change()
    imFieldChgd = True
End Sub

Private Sub txtDays_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropDate_Change()
    imChgDropDate = True
    imDateChgd = True
    imFieldChgd = True
End Sub

Private Sub txtDropDate_GotFocus()
    txtDropDate.ZOrder
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = txtDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
'        If grdPledge.Text <> slStr Then
'            imFieldChgd = True
'        End If
        imCloseListBox = False
        lbcStatus.ListIndex = llRow
        txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
        txtDropdown.SelStart = ilLen
        txtDropdown.SelLength = Len(txtDropdown.Text)
'        grdPledge.Text = lbcStatus.List(lbcStatus.ListIndex)
'    Else
'        grdPledge.Text = ""
    End If
End Sub

Private Sub txtDropdown_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub txtDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If txtDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub txtDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        gProcessArrowKey Shift, KeyCode, lbcStatus, True ', imLbcArrowSetting
    End If
End Sub

Private Sub txtEmailAddr_Change()
    imFieldChgd = True
End Sub

Private Sub txtEmailAddr_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtEndDate_Change()
    imFieldChgd = True
End Sub

Private Sub txtEndDate_GotFocus()
    txtEndDate.ZOrder
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtEndDate_LostFocus()
    'Removed 11/16- Jim
    'If txtOffAirDate.Text = "" Then
    '    If txtEndDate.Text <> "" Then
    '        If gIsDate(txtEndDate.Text) = True Then
    '            txtOffAirDate.Text = txtEndDate.Text
    '        End If
    '    End If
    'End If
End Sub

Private Sub txtET_Change()
    Dim slStr As String
    
    Select Case grdET.Col
        Case ETTIMEINDEX
            slStr = Trim$(txtET.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                grdET.CellForeColor = vbBlack
                slStr = gConvertTime(slStr)
                If Second(slStr) = 0 Then
                    slStr = Format$(slStr, sgShowTimeWOSecForm)
                Else
                    slStr = Format$(slStr, sgShowTimeWSecForm)
                End If
                If grdET.Text <> slStr Then
                    imFieldChgd = True
                    '8/12/16: Separtated Pledge and Estimate so that est can be saved in the past
                    'bmPledgeDataChgd = True
                    bmETDataChgd = True
                End If
                grdET.Text = slStr
            '9/4/14: Allow Time to be removed (To remove record, Day and time must be blanked out)
            ElseIf (slStr = "") Then
                grdET.Text = slStr
            End If
    End Select
End Sub

Private Sub txtET_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtLdMult_Change()
    imFieldChgd = True
End Sub

Private Sub txtLdMult_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtLog_Change()
    imFieldChgd = True
End Sub

Private Sub txtLog_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtLogPassword_Change()
    imFieldChgd = True
End Sub

Private Sub txtLogPassword_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtNoCDs_Change()
    imFieldChgd = True
End Sub

Private Sub txtNoCDs_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtOffAirDate_Change()
    imDateChgd = True
    imFieldChgd = True
End Sub

Private Sub txtOffAirDate_GotFocus()
    txtOffAirDate.ZOrder
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtOnAirDate_Change()
    imChgOnAirDate = True
    If IsAgmntDirty = True Then
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        mEnableControls
        imOkToChange = True
    End If
    imDateChgd = True
    imFieldChgd = True
End Sub

Private Sub txtOnAirDate_GotFocus()
    txtOnAirDate.ZOrder
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtOther_Change()
    imFieldChgd = True
End Sub

Private Sub txtOther_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtRetDate_Change()
    imFieldChgd = True
End Sub

Private Sub txtRetDate_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtStartDate_Change()
    imFieldChgd = True
    '5589 Dan dick says add flag here
    imDateChgd = True
End Sub

Private Sub txtStartDate_GotFocus()
    txtStartDate.ZOrder
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Function mFailsIDCTest() As Boolean
    Dim blRet As Boolean
    
   ' blRet = False
    'Test if IDC receiver defined only if Audio IDC selected.
    'D.L. 11/3/14 added If statement below
    '7701
    blRet = mMultiListIsData(Vendors.iDc, lbcAudioDelivery)
'    With lbcAudioDelivery
'        If .ListIndex > -1 Then
'            If .ItemData(.ListIndex) = Vendors.iDc Then
'                blRet = True
'            End If
'        End If
'    End With
    If blRet Then
        blRet = False
    'If rbcAudio(3).Value = True Then
        'Is this vehicle in evt table and in that table defined as idc export?
        SQLQuery = "select count(*) as myCount from evt_Export_Vehicles inner join eht_Export_Header on evtEhtCode = ehtCode where ehtExportType = 'D'"
        SQLQuery = SQLQuery & " AND ehtSubType = 'S' and evtVefCode = " & imVefCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            'also, is IDCSiteId not set?
            If rst!mycount > 0 And Len(txtIDCReceiverID.Text) = 0 Then
                blRet = True
            End If
        End If
    End If
    mFailsIDCTest = blRet
End Function

Private Function mSave(iAsk As Integer) As Integer
    Dim sStation As String
    Dim iSigned As Integer
    Dim iComp As Integer
    Dim iPost As Integer
    Dim iSendTape As Integer
    Dim iNoCDs As Integer
    Dim iPrintCP As Integer
    Dim sSuppressNotices As String
    Dim sNCR As String                      '7-6-09
    Dim sFormerNCR As String                '7-6-09
    Dim iCarryCmml As Integer
    Dim iBarCodes As Integer
    Dim iTimeType As Integer
    Dim iLoad As Integer
    Dim iLoop As Integer
    Dim sFdStTime As String
    Dim sFdEdTime As String
    Dim sPdStTime As String
    Dim sPdEdTime As String
    Dim AgreeStart As String
    Dim AgreeEnd As String
    Dim OnAir As String
    Dim OffAir As String
    Dim RetDate As String
    Dim Drop As String
    Dim CurTime As String
    Dim sACName As String
    Dim sACPhone As String
    Dim iNewRec As Integer
    Dim ilRet As Integer
    Dim sComment As String
    Dim ilExportType As Integer
    Dim ilLogType As Integer
    Dim ilPostType As Integer
    Dim ilSendLogEMail As Integer
    Dim slPassword As String
    Dim slEmailAddr As String
    Dim slTemp As String
    Dim rst_pw As ADODB.Recordset
    Dim llTemp As Long
    Dim slMulticast As String
    Dim slStr As String
    Dim slStr2 As String
    Dim ilLoop As Integer
    Dim llAttCode As Long
    Dim tmp_rst As ADODB.Recordset
    Dim slLabelID As String
    Dim slShipInfo As String
    Dim ilMultiOK As Integer
    Dim slRetStr As String
    Dim llAffAE As Long
    Dim slRadarClearType As String
    Dim ilTemp As Integer
    Dim slForbidSplitLive As String
    Dim slVoiceTracked As String
    Dim llXDReceiverID As Long
    Dim slIDCReceiverID As String
    Dim slMonthlyWebPost As String
    Dim slWebInterface As String
    Dim slExportToWeb As String
    Dim slExportToUnivision As String
    Dim slExportToMarketron As String
    Dim ilMktRepUstCode As Integer
    Dim ilServRepUstCode As Integer
    Dim slCDStartTime As String
    Dim slPledgeType As String
    Dim llCode As Long
    Dim ilNoAirPlays As Integer
    Dim llRow As Long
    Dim slDates As String
    Dim blShowMulticastMsg As Boolean
    Dim slAttSQLQuery As String
    Dim llMCAttCode As Long
    Dim ilMCShttCode As Integer
    Dim slAddStr As String
    Dim slChgStr As String
    Dim ilShttIndex As Integer
    Dim llMCOnAir As Long
    '5352
    Dim slInvalidEmail As String
    'Dim slProgContProv As String
    'Dim slCommAudProv As String
    '6466
    Dim slIDCGroup As String
    '11/19/13: Create CPTT for multi-cast (ttp 5385)
    Dim llSvAttCode As Long
    Dim ilSvShttCode As Long
    Dim slAudioDelivery As String
    '6592
    Dim slExportToCBS As String
    Dim slExportToJelli As String
    '3/23/15: Add Send Delays to XDS
    Dim slSendDelays As String
    Dim slSendNotCarried As String
    
    Dim slService As String '10/28/14
    '7701
    Dim ilLogVendor() As Integer
    Dim ilAudioVendor() As Integer
    Dim blFoundOne As Boolean
    Dim ilCount As Integer
    '6/28/18
    Dim llNewMCAttCode As Long
    Dim llMCOffAir As Long
    Dim llMCDropDate As Long
    '4/3/19
    Dim slExcludeFillSpot As String
    Dim slExcludeCntrTypeQ As String    'Per Inquiry
    Dim slExcludeCntrTypeR As String    'Direct Response
    Dim slExcludeCntrTypeT As String    'Remnant
    Dim slExcludeCntrTypeM As String    'Promo
    Dim slExcludeCntrTypeS As String    'PSA
    Dim slExcludeCntrTypeV As String    'Reservation
    
    
    On Error GoTo ErrHand
    'ttp 5352
    slInvalidEmail = udcContactGrid.InValidEmails()
    If Len(slInvalidEmail) > 0 Then
        If MsgBox("Do you wish to continue to save, or cancel?  The following email(s) are invalid: " & slInvalidEmail, vbOKCancel + vbInformation, "Invalid Email") = vbCancel Then
            mSave = False
            Exit Function
        End If
    End If
    '6466
    For iLoop = 0 To 2
        If optIDCGroup(iLoop).Value Then
            Select Case iLoop
                Case 1
                    slIDCGroup = "S"
                Case 2
                    slIDCGroup = "L"
                Case Else
                    slIDCGroup = "N"
            End Select
            Exit For
        End If
    Next iLoop
    ' ttp 5319 Dan M  IDC 3/22/12
    If mFailsIDCTest() Then
        If MsgBox("Warning - this Agreement may require a Site ID - Ok to continue or Cancel to stop save", vbOKCancel + vbQuestion, "IDC Vehicle") = vbCancel Then
            mSave = False
            Exit Function
        End If
    End If
    '7701
    If mIsDeliveryInconsistent() Then
        mSave = False
        Exit Function
    End If
    '8017
    If lbcLogDelivery.SelCount > 2 Then
        gMsgBox "Cannot have more than 2 vendors in Log Delivery.  Save cancelled", vbExclamation
        mSave = False
        Exit Function
    End If
    If lbcAudioDelivery.SelCount > 2 Then
        gMsgBox "Cannot have more than 2 vendors in Audio Delivery.  Save cancelled", vbExclamation
        mSave = False
        Exit Function
    End If
    '4/4/13: Add test that
    If smPledgeByEvent = "Y" Then
        If UBound(tmPetInfo) <= LBound(tmPetInfo) Then
            gMsgBox "The Pledge by Event information must be defined"
            mSave = False
            Exit Function
        End If
    End If
    '12/15/14: Remove / if last character in the date fields
    '11/15/14: Remove backslash if last character
    If right(txtOffAirDate.Text, 1) = "/" Then
        txtOffAirDate.Text = Left(txtOffAirDate.Text, Len(txtOffAirDate.Text) - 1)
    End If
    If right(txtDropDate.Text, 1) = "/" Then
        txtDropDate.Text = Left(txtDropDate.Text, Len(txtDropDate.Text) - 1)
    End If
    If right(txtOnAirDate.Text, 1) = "/" Then
        txtOnAirDate.Text = Left(txtOnAirDate.Text, Len(txtOnAirDate.Text) - 1)
    End If
    If right(txtStartDate.Text, 1) = "/" Then
        txtStartDate.Text = Left(txtStartDate.Text, Len(txtStartDate.Text) - 1)
    End If
    If right(txtEndDate.Text, 1) = "/" Then
        txtEndDate.Text = Left(txtEndDate.Text, Len(txtEndDate.Text) - 1)
    End If

    'D.S. 11/14/06
    'If Not imOkToChange Then
        'D.S. 11/14/06 I don't know what this code was for, but it goofs things up
        'If Len(Trim(txtOffAirDate.Text)) < 1 And Len(Trim(txtDropDate.Text)) < 1 And lmAttCode <> 0 Then
        '    mSave = True
        '    Exit Function
        'End If
        ' The user can only change the off air or drop date when this flag is set.
        If Len(Trim(txtOffAirDate.Text)) > 0 Then
            If gIsDate(txtOffAirDate.Text) = True Then
                If gIsDate(smLastPostedDate) = True Then
                    If DateValue(smLastPostedDate) >= DateValue(txtOffAirDate.Text) Then
                        mMousePointer vbDefault
                        gMsgBox "The off air date must be greater than " & Format$(smLastPostedDate, sgShowDateForm)
                        If imTabIndex = TABMAIN Then
                            txtOffAirDate.SetFocus
                        End If
                        mSave = False
                        Exit Function
                    End If
                End If
            End If
        End If

        If Len(Trim(txtDropDate.Text)) > 0 Then
            If gIsDate(txtDropDate.Text) = True Then
                If gIsDate(smLastPostedDate) = True Then
                    If DateValue(smLastPostedDate) >= DateValue(txtDropDate.Text) Then
                        mMousePointer vbDefault
                        gMsgBox "The drop date date must be greater than " & Format$(smLastPostedDate, sgShowDateForm)
                        If imTabIndex = TABMAIN Then
                            If txtDropDate.Visible And txtDropDate.GetEnabled Then
                                txtDropDate.SetFocus
                            End If
                        End If
                        mSave = False
                        Exit Function
                    End If
                End If
            End If
        End If
    'End If

    ilRet = mValidateRows
    If Not ilRet Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If

    slLabelID = Trim(txtLabelID.Text)
    slShipInfo = Trim(txtShipInfo.Text)
    
    'web email check
    'If gUsingWeb And ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) Then
    If gUsingWeb And ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) Then
        ilRet = gTestForMultipleEmail(txtEmailAddr.Text, "Reg")
        If ilRet = False Then
            mMousePointer vbDefault
            If imTabIndex = TABPERSONNEL Then
                If txtEmailAddr.Enabled Then
                    txtEmailAddr.SetFocus
                End If
            End If
            gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Web Email Address Before Continuing", vbExclamation
            gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Web Email Address Before Continuing", "AffErrorLog.Txt", False
            mSave = False
            Exit Function
        End If
    End If
    
    'D.S. 01/12/09
    'web agreement password check
    'If gUsingWeb And rbcExportType(1).Value = True Then
    '    If txtLogPassword.text = "" Then
    '        gMsgBox "A valid password must be supplied"
    '        mSave = False
    '        Exit Function
    '    End If
    'End If
    
    'D.S. 01/12/09
    'web station password check
'    If gUsingWeb And rbcExportType(1).Value = True Then
'        SQLQuery = "SELECT shttWebPW, shttWebEmail"
'        SQLQuery = SQLQuery & " FROM shtt"
'        SQLQuery = SQLQuery + " WHERE (shttCode = " & imShttCode & ")"
'        Set rst_pw = gSQLSelectCall(SQLQuery)
'
'        If rst_pw.EOF = False Then
'            slTemp = Trim$(rst_pw!shttWebPW)
'            If slTemp = "" Then
'                SQLQuery = "Update SHTT Set shttWebPW = '" & txtLogPassword.text & "' Where shttCode = " & imShttCode
'                'cnn.Execute SQLQuery, rdExecDirect
'                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                    GoSub ErrHand:
'                End If
'            End If
'            slTemp = Trim$(rst_pw!shttWebEmail)
'            If slTemp = Trim$(rst_pw!shttWebEmail) Then
'                txtEmailAddr.text = gFixQuote(txtEmailAddr.text)
'                SQLQuery = "Update SHTT Set shttWebEmail = '" & txtEmailAddr.text & "' Where shttCode = " & imShttCode
'                'cnn.Execute SQLQuery, rdExecDirect
'                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                    GoSub ErrHand:
'                End If
'            End If
'        End If
'    End If
    
    'Save the Y/N value if the agreement is part of a multicast or not.
    
    
    slMonthlyWebPost = "N"
    
    If chkMonthlyWebPost.Value = vbChecked Then
        '1/20/11: It is belived that this test was added for Radio America
        'If required by Radio America, I site question or vehicle question should be added
        'Sara guess is that this test was required to make sure spots from one program
        'did not overlap spots from another program
        ilRet = mPostMonthly()
        If Not ilRet Then
            mMousePointer vbDefault
            gMsgBox "When using Monthly Posting all Delayed times must have a time span of time equal to or less than the feed time span."
            mSave = False
            Exit Function
        End If
        slMonthlyWebPost = "Y"
    Else
        slMonthlyWebPost = "N"
    End If
    
'    If rbcMulticast(0).Value Then
'        slMulticast = "Y"
'    Else
'        slMulticast = "N"
'    End If
    'If optTimeType(0).Value Then
    '    slPledgeType = "D"
    'ElseIf optTimeType(1).Value Then
    '    slPledgeType = "A"
    'Else
    '    slPledgeType = ""
    'End If
    'Determine whether avails or dayparts
    '0 = dayparts, 1 = avails, 2 = CD/Tape
    'iTimeType = -1
    'For iLoop = 0 To 2
    '    If optTimeType(iLoop) Then
    '        iTimeType = iLoop
    '        Exit For
    '    End If
    'Next iLoop
    slPledgeType = sgAirPlay1TimeType
    If (slPledgeType <> "C") And (slPledgeType <> "A") And (slPledgeType <> "D") Then
        'Check Pledge definition
        slPledgeType = "A"
        For iLoop = 0 To UBound(tgDat) - 1 Step 1
            If gTimeToLong(Format$(tgDat(iLoop).sFdETime, "h:m:ssam/pm"), True) - gTimeToLong(Format$(tgDat(iLoop).sFdSTime, "h:m:ssam/pm"), False) > AVAIL_OR_DP_TIME Then
                slPledgeType = "D"
                Exit For
            End If
        Next iLoop
    End If
    If slPledgeType = "C" Then
        iTimeType = 2
    ElseIf slPledgeType = "A" Then
        iTimeType = 1
    ElseIf slPledgeType = "D" Then
        iTimeType = 0
    Else
        
        iTimeType = -1
    End If
    ilMktRepUstCode = 0
    If cbcMarketRep.ListIndex >= 1 Then
        ilMktRepUstCode = cbcMarketRep.GetItemData(cbcMarketRep.ListIndex)
    End If
    ilServRepUstCode = 0
    If cbcServiceRep.ListIndex >= 1 Then
        ilServRepUstCode = cbcServiceRep.GetItemData(cbcServiceRep.ListIndex)
    End If
    
    llAffAE = 0
    'If cboAffAE.ListIndex > 0 Then
    '    llAffAE = cboAffAE.ItemData(cboAffAE.ListIndex)
    'End If
    
    'Set default CD Start Time
    slCDStartTime = sgCDStartTime
    If sgCDStartTime = "" Then
        sgCDStartTime = "00:00:00"
    End If
    
    mSave = False
    If sgUstWin(2) <> "I" Then
        mMousePointer vbDefault
        gMsgBox "Not Allowed to Save.", vbOKOnly
        Exit Function
    End If
    If (imShttCode <= 0) Or (imVefCode <= 0) Then
        mMousePointer vbDefault
        Exit Function
    End If
    If Not mTestZone() Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If
    If sgUstPledge = "Y" Then
        'Check Air Plays
        ilRet = mCheckAirPlays()
        If ilRet = 1 Then
            mMousePointer vbDefault
            MsgBox "Air Plays number missing"
            mSave = False
            Exit Function
        End If
        If ilRet = 2 Then
            mMousePointer vbDefault
            MsgBox "Not All Air Plays Defined"
            mSave = False
            Exit Function
        End If
        If ilRet = 3 Then
            mMousePointer vbDefault
            MsgBox "Too many Air Plays Defined"
            mSave = False
            Exit Function
        End If
        If ilRet = 4 Then
            mMousePointer vbDefault
            ilRet = gMsgBox("Warning: All Pledges set to 9 - Not Carried" & Chr$(13) & Chr$(10) & "Proceed With Save?", vbYesNo)
            If ilRet = vbNo Then
                mSave = False
                Exit Function
            End If
            mMousePointer vbHourglass
        End If
        If Not mTestDaypart(True) Then
            mMousePointer vbDefault
            mSave = False
            Exit Function
        End If
        If Not mCheckSendDelays() Then
            mMousePointer vbDefault
            mSave = False
            Exit Function
        End If
        '9/15/14: Moved here from below
        If Not mMoveDaypart() Then
            mMousePointer vbDefault
            mSave = False
            Exit Function
        End If
        If Not mCheckEstTimes() Then
            mMousePointer vbDefault
            MsgBox "Estimated Times in Red are outside of Pledge Times"
            mSave = False
            Exit Function
        End If
        '9/4/14: Verify that estimate day defined if estimate time defined
        If Not mTestEstimateTimes() Then
            mMousePointer vbDefault
            mSave = False
            Exit Function
        End If
        '9/15/14: Moved above Estimjate time testing
        'If Not mMoveDaypart() Then
        '    mMousePointer vbDefault
        '    mSave = False
        '    Exit Function
        'End If
    End If
    
    'Check for Daypart Conflicts
    If igLiveDayPart = True Then
        ilRet = mTestForDPConflict()
        If ilRet = 1 Then
            mMousePointer vbDefault
            If igLiveDayPart Then
                gMsgBox "Daypart Conflict Exist in the Feed Area." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "                   Please Correct.", vbOKOnly
            Else
                gMsgBox "Daypart Conflict Exist in the Sold Area." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "                   Please Correct.", vbOKOnly
            End If
            mSave = False
            Exit Function
        ElseIf ilRet = 2 Then
            mMousePointer vbDefault
            gMsgBox "Daypart Conflict Exist in the Pledge Area." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "                   Please Correct.", vbOKOnly
            mSave = False
            Exit Function
        End If
    End If
    
    mSave = True
    'smCurDate = Format(gNow(), sgShowDateForm)
    CurTime = Format(gNow(), sgShowTimeWSecForm)
    
    'CP Control check
    'If rbcExportType(0).Value = False And rbcExportType(1).Value = False And rbcExportType(2).Value = False And rbcExportType(3).Value = False Then
    If rbcExportType(0).Value = False And rbcExportType(1).Value = False Then
            mMousePointer vbDefault
            gMsgBox "You must specify Affidavit Control on the Agreement Delivery Tab."
            mSave = False
            Exit Function
    End If
    '7701 Dan if not manual, then exportTo is always checked(always goes to web), so this test isn't needed
'    If rbcExportType(1).Value = True Then
'        '7701
'        If (ckcExportTo(0).Value = vbUnchecked) And mRetrieveMultiListString(lbcLogDelivery) = "" Then
'        '6592 added CBS (4)
'       ' If (ckcExportTo(0).Value = vbUnchecked) And (ckcExportTo(1).Value = vbUnchecked) And (ckcExportTo(2).Value = vbUnchecked) And (ckcExportTo(3).Value = vbUnchecked) And (ckcExportTo(4).Value = vbUnchecked) Then
'        'If (ckcExportTo(0).Value = vbUnchecked) And (ckcExportTo(1).Value = vbUnchecked) And (ckcExportTo(2).Value = vbUnchecked) And (ckcExportTo(3).Value = vbUnchecked) Then
'            mMousePointer vbDefault
'            gMsgBox "You must specify Export To on the Agreement Delivery Tab."
'            mSave = False
'            Exit Function
'        End If
'    End If
    slExportToWeb = "N"
    slExportToMarketron = "N"
    '7701 many of these fields aren't used, but still used in sql
    slWebInterface = ""
    slExportToUnivision = "N"
    slExportToCBS = "N"
    slExportToJelli = "N"
    'If rbcExportType(3).Value = True Then
    '    ilExportType = 1
    '    slWebInterface = "C"
    'Else
    '    ' 0 = manual, 1 = web, 2 = Univision
    '    For iLoop = 0 To 2
    '        If rbcExportType(iLoop).Value = True Then
    '            ilExportType = iLoop
    '            Exit For
    '        End If
    '    Next iLoop
    'End If
'7701
    ilAudioVendor = mRetrieveMultiListIntegers(lbcAudioDelivery)
    ilLogVendor = mRetrieveMultiListIntegers(lbcLogDelivery)
    If rbcExportType(0).Value = True Then
        ilExportType = 0
'7701
    Else
        ilExportType = 1
        If mMultiListIsData(Vendors.NetworkConnect, lbcLogDelivery) Then
            slExportToMarketron = "Y"
        End If
'        With lbcLogDelivery
'            If .ListIndex > -1 Then
'                ilLogVendor = .ItemData(.ListIndex)
'                'Used below
'                If ilLogVendor = Vendors.NetworkConnect Then
'                    slExportToMarketron = "Y"
'                End If
'            End If
'        End With
        If gUsingWeb Then
             slExportToWeb = "Y"
        Else
            slExportToWeb = "N"
        End If
    End If
    '8000
    If ckcUnivision.Value = vbChecked Then
        slExportToUnivision = "Y"
    End If
'    Else
'        ilExportType = 1
'        If gUsingWeb Then
'            If ckcExportTo(1).Value = vbChecked Then
'                slWebInterface = "C"
'            '    slExportToWeb = "Y"
'            End If
'            'If ckcExportTo(0).Value = vbChecked Then
'                slExportToWeb = "Y"
'            'End If
'        Else
'            slExportToWeb = "N"
'            slWebInterface = "N"
'        End If
'        If ckcExportTo(2).Value = vbChecked Then
'            slExportToUnivision = "Y"
'        End If
'        If ckcExportTo(3).Value = vbChecked Then
'            slExportToMarketron = "Y"
'        End If
'        '6592
'        If ckcExportTo(4).Value = vbChecked Then
'            slExportToCBS = "Y"
'            slExportToWeb = "Y"
'        End If
'        If ckcExportTo(6).Value = vbChecked Then
'            slExportToJelli = "Y"
'        End If
'    End If
    If slExportToMarketron = "Y" Then
        If gDateValue(smMktronActiveDate) = gDateValue("1/1/1970") Then
            smMktronActiveDate = Format(gNow(), sgShowDateForm)
        End If
    Else
        smMktronActiveDate = "1/1/1970"
    End If
'    With lbcAudioDelivery
'        If .ListIndex > -1 Then
'            ilAudioVendor = .ItemData(.ListIndex)
'        End If
'    End With

'    If rbcAudio(0).Value Then
'        slAudioDelivery = "X"
'    ElseIf rbcAudio(1).Value Then
'        slAudioDelivery = "B"
'     ElseIf rbcAudio(2).Value Then
'        slAudioDelivery = "W"
'    ElseIf rbcAudio(3).Value Then
'        slAudioDelivery = "I"
'    ElseIf rbcAudio(4).Value Then
'        slAudioDelivery = "P"
'    Else
'        slAudioDelivery = "N"
'    End If

'   7701  frcLogType disabled and never visible (5/11)
'    'if network website is selected then make sure that they selected a Log Format
'    'If gUsingWeb And ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) Then
'    If frcLogType.Visible = True Then
'        If gUsingWeb And ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) Then
'            If rbcLogType(0).Value = False And rbcLogType(1).Value = False And rbcLogType(2).Value = False Then
'                mMousePointer vbDefault
'                gMsgBox "You must specify Log Format type on the Agreement Delivery Tab."
'                mSave = False
'                Exit Function
'            End If
'        End If
'    End If
    
    ilLogType = 0
'   7701 frclogType never visible, as above
'    'This indicate how the logs shows the times
'    ' 0 = exact time, 1 = dayparts, 2 = break numbers
'    If frcLogType.Visible = True Then
'        For iLoop = 0 To 2
'            If rbcLogType(iLoop).Value = True Then
'                ilLogType = iLoop
'                Exit For
'            End If
'        Next iLoop
'    End If
'
    'if network website is selected then make sure that they selected a Posting Method
'7701
    If frcPostType.Visible = True Then
        If gUsingWeb And ((rbcExportType(1).Value = True)) Then
            If rbcPostType(0).Value = False And rbcPostType(1).Value = False And rbcPostType(2).Value = False Then
                mMousePointer vbDefault
                gMsgBox "You must specify Posting Method on the Agreement Delivery Tab."
                mSave = False
                Exit Function
            End If
        End If
    End If
    'If gUsingWeb And ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) Then
'    If frcPostType.Visible = True Then
'        If gUsingWeb And ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) Then
'            If rbcPostType(0).Value = False And rbcPostType(1).Value = False And rbcPostType(2).Value = False Then
'                mMousePointer vbDefault
'                gMsgBox "You must specify Posting Method on the Agreement Delivery Tab."
'                mSave = False
'                Exit Function
'            End If
'        End If
'    End If

    ilPostType = 0
    ' 0 = exact time, 1 = ordered times, 2 = start time
    If frcPostType.Visible = True Then
        For iLoop = 0 To 2
            If rbcPostType(iLoop).Value = True Then
                ilPostType = iLoop
                Exit For
            End If
        Next iLoop
    End If
' removed 7701
'    'if network website is selected then make sure that they selected Send E-Mail Notification
'    'If gUsingWeb And ((rbcExportType(1).Value = True) Or (rbcExportType(3).Value = True)) Then
'    If gUsingWeb And ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) Then
'        If rbcSendLogEMail(0).Value = False And rbcSendLogEMail(1).Value = False Then
'            mMousePointer vbDefault
'            gMsgBox "You must specify Send E-Mail Notification on the Agreement Delivery Tab."
'            mSave = False
'            Exit Function
'        End If
'    End If
    
    '7/28/15: Check Rights
    If Not udcContactGrid.VerifyRights("A") Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If

    ' 0 = Yes, 1 = No
    For iLoop = 0 To 1
        If rbcSendLogEMail(iLoop).Value = True Then
            ilSendLogEMail = iLoop
            Exit For
        End If
    Next iLoop
    
    'D.S. 01/12/09
    'slPassword = gFixQuote(Trim$(txtLogPassword.text))
    'slEmailAddr = gFixQuote(Trim$(txtEmailAddr.text))
    
    
'    If Trim$(txtProgContProv.Text) = "" Then
'        slProgContProv = " "
'    Else
'        slProgContProv = Trim$(txtProgContProv.Text)
'    End If
'
'    If Trim$(txtCommAudProv.Text) = "" Then
'        slCommAudProv = " "
'    Else
'        slCommAudProv = Trim$(txtCommAudProv.Text)
'    End If
    
    
    'Determine if contract signed
    '0 = Not returned, 1 = returned, 2 = rejected
    iSigned = -1
    For iLoop = 0 To 2
        If optSigned(iLoop).Value Then
            iSigned = iLoop
            Exit For
        End If
    Next iLoop
    
    iPrintCP = -1
    For iLoop = 0 To 1
        'Yes = 0 No = 1 to print
        If optPrintCP(iLoop).Value Then
            iPrintCP = iLoop
            Exit For
        End If
    Next iLoop
    
    sSuppressNotices = "N"
    If optSuppressNotices(0).Value Then
        sSuppressNotices = "Y"
    End If
    
    '7-6-09
    If optNCR(0).Value Then     'user manually set it as NCR
        sNCR = "Y"
    Else
        sNCR = "N"
        If smPreviousNCR = "Y" Then  'if previous value of attncr was a non-compliant agreement, and now
                                    'changing to compliant, set flag to indicate former offender
            sFormerNCR = "Y"
        End If
    End If
    If optFormerNCR(0).Value Then       'if already former offender, leave alone
        sFormerNCR = "Y"                'if former ncr is already Y, dont turn it off
    End If
   
    '10/28/14
    slService = "N"
    If optService(0).Value Then
        slService = "Y"
    End If
   
    '3/23/15: Add Send Delays to XDS
    slSendDelays = "N"
    '7701
    If bmSupportXDSDelay And ckcSendDelays.Value = vbChecked Then
        If mMultiListIsData(Vendors.XDS_Break, lbcAudioDelivery) Or mMultiListIsData(Vendors.XDS_ISCI, lbcAudioDelivery) Then
            slSendDelays = "Y"
        End If
'        With lbcAudioDelivery
'            If .ListIndex > -1 Then
'                If .ItemData(.ListIndex) = Vendors.XDS_Break Or .ItemData(.ListIndex) = Vendors.XDS_ISCI Then
'                    slSendDelays = "Y"
'                End If
'            End If
'        End With
    End If
    slSendNotCarried = "N"
    '7701
    If ckcSendNotCarried.Value = vbChecked Then
        If mMultiListIsData(Vendors.XDS_Break, lbcAudioDelivery) Or mMultiListIsData(Vendors.XDS_ISCI, lbcAudioDelivery) Then
            slSendNotCarried = "Y"
        End If
    End If
    
    'If bmSupportXDSDelay And ckcSendDelays.Value = vbChecked And (rbcAudio(0).Value Or rbcAudio(1).Value) Then
'        slSendDelays = "Y"
'    End If
   
   '0 = yes/carry, 1 = no/don't carry
    iCarryCmml = -1
    For iLoop = 0 To 1
        If optCarryCmml(iLoop).Value Then
            iCarryCmml = iLoop
            Exit For
        End If
    Next iLoop
    
   '0 = yes, 1 = no
    iSendTape = -1
    For iLoop = 0 To 1
        If optSendTape(iLoop).Value Then
            iSendTape = iLoop
            Exit For
        End If
    Next iLoop
        
    'relic code - hidden from users
    '0 = yes, 1 = no
    'Determine if bar codes peferred
    iBarCodes = -1
    For iLoop = 0 To 1
        If optBarCode(iLoop).Value Then
            iBarCodes = iLoop
            Exit For
        End If
    Next iLoop
    
    '4/29/14
    'Determine perferred method of compensation
    ''0 = network, 1 = affiliate, 2 = barter
    '0=Barter; 1=Affiliate; 2=Network
    If smCompensation = "Y" Then
        iComp = -1
        For iLoop = 0 To 2
            If optComp(iLoop).Value Then
                iComp = iLoop
                Exit For
            End If
        Next iLoop
    Else
        iComp = 0
    End If
    'Determine Posting Method
    '0 = CP reveipt Only, 1 = Spot Count, 2 = Exact times by date, 3 = exact time by adv.
    If ilExportType = 0 Then
        iPost = -1
        For iLoop = 0 To 3
            If optPost(iLoop).Value Then
                iPost = iLoop
                Exit For
            End If
        Next iLoop
    Else
        'Exact Time by Date
        iPost = 2
    End If
    
    If ckcProhibitSplitCopy.Value = vbChecked Then
        slForbidSplitLive = "Y"
    Else
        slForbidSplitLive = "N"
    End If
    
    'Test for proper dates and convert datatypes
    If txtStartDate.Text = "" Then
        AgreeStart = Format("1/1/1970", "m/d/yyyy")   'Placeholder value to prevent using Nulls/outer joins
    Else
        If gIsDate(txtStartDate.Text) = False Then
            mMousePointer vbDefault
            mSave = False
            Beep
            If imTabIndex = TABMAIN Then
                If txtStartDate.Visible And txtStartDate.GetEnabled Then
                    txtStartDate.SetFocus
                End If
            End If
            Exit Function
        Else
            AgreeStart = gAdjYear(Format(txtStartDate.Text, sgShowDateForm))
        End If
    End If
    
    
    If txtEndDate.Text = "" Then
        AgreeEnd = Format("12/31/2069", "m/d/yyyy")
    Else
        If gIsDate(txtEndDate.Text) = False Then
            mMousePointer vbDefault
            mSave = False
            Beep
            If imTabIndex = TABMAIN Then
                If txtEndDate.GetEnabled Then
                    txtEndDate.SetFocus
                End If
            End If
            mSave = False
            Exit Function
        Else
            AgreeEnd = gAdjYear(Format(txtEndDate.Text, sgShowDateForm))
        End If
    End If
    
    
    If txtOnAirDate.Text = "" Then
        OnAir = Format("1/1/1970", "m/d/yyyy")
    Else
        If gIsDate(txtOnAirDate.Text) = False Then
            mMousePointer vbDefault
            mSave = False
            Beep
            If imTabIndex = TABMAIN Then
                If txtOnAirDate.GetEnabled Then
                    txtOnAirDate.SetFocus
                End If
            End If
            mSave = False
            Exit Function
        Else
            OnAir = gAdjYear(Format(txtOnAirDate.Text, sgShowDateForm))
        End If
    End If
    
    'D.S. 12/2/05  Don't let them save an agreement that's ON AIR date is not on a Monday
    'If imChgDropDate And imChgOnAirDate Then
        If Weekday(OnAir) <> vbMonday Then
            mMousePointer vbDefault
            gMsgBox "On Air Date Must be a Monday", vbOKOnly
            If imTabIndex = TABMAIN Then
                If txtOnAirDate.GetEnabled Then
                    txtOnAirDate.SetFocus
                End If
            End If
            mSave = False
            Exit Function
        End If
    'End If
    
    'D.S. 12/2/05  Don't let them save an agreement that's ON AIR date <= the START DATE
'    If DateValue(gAdjYear(txtOnAirDate.text)) < DateValue(gAdjYear(txtStartDate.text)) Then
'        gMsgBox "On Air Date Must Equal to or After the Start Date", vbOKOnly
'        If imTabIndex = 1 Then
'            If txtStartDate.Enabled Then
'                txtStartDate.SetFocus
'            End If
'        End If
'        mSave = False
'        Exit Function
'    End If
    
    If txtOffAirDate.Text = "" Then
        OffAir = Format("12/31/2069", "m/d/yyyy")
    Else
        If gIsDate(txtOffAirDate.Text) = False Then
            mMousePointer vbDefault
            mSave = False
            Beep
            If imTabIndex = TABMAIN Then
                If txtOffAirDate.GetEnabled Then
                    txtOffAirDate.SetFocus
                End If
            End If
            mSave = False
            Exit Function
        Else
            OffAir = gAdjYear(Format(txtOffAirDate.Text, sgShowDateForm))
        End If
    End If
    
    
    If iSigned = 1 Then
        If txtRetDate.Text = "" Then
            RetDate = Format("1/1/1970", "m/d/yyyy")
        Else
            If gIsDate(txtRetDate.Text) = False Then
                mMousePointer vbDefault
                mSave = False
                Beep
                If imTabIndex = TABMAIN Then
                    If txtRetDate.Enabled Then
                        txtRetDate.SetFocus
                    End If
                End If
                mSave = False
                Exit Function
            Else
                RetDate = Format(txtRetDate.Text, sgShowDateForm)
            End If
        End If
    Else
        RetDate = Format("12/31/2069", "m/d/yyyy")
    End If
    
    If txtDropDate.Text = "" Then
        Drop = Format("12/31/2069", "m/d/yyyy")
    Else
        If gIsDate(txtDropDate.Text) = False Then
            mMousePointer vbDefault
            mSave = False
            Beep
            If imTabIndex = TABMAIN Then
                If txtDropDate.Visible And txtDropDate.GetEnabled Then
                    txtDropDate.SetFocus
                End If
            End If
            mSave = False
            Exit Function
        Else
            Drop = gAdjYear(Format(txtDropDate.Text, sgShowDateForm))
            If IsAgmntDirty = False Then
                If DateValue(gAdjYear(Drop)) <= DateValue(gAdjYear(OnAir)) Then
                    Drop = Format("12/31/2069", "m/d/yyyy")
                Else
                    If gMsgBox("Retain Drop Date, " & Format$(Drop, "m/d/yy") & "?", vbYesNo) = vbNo Then
                        Drop = Format("12/31/2069", "m/d/yyyy")
                    End If
                End If
            End If
        End If
    End If
    
    'Check that the dates don't affect any Posted CP's
    If imSource = 1 Then
        If (DateValue(gAdjYear(OffAir)) < DateValue(smSvOffAirDate)) Or (DateValue(gAdjYear(Drop)) < DateValue(smSvDropDate)) Then
            If DateValue(gAdjYear(OffAir)) < DateValue(gAdjYear(Drop)) Then
                If DateValue(gAdjYear(OffAir)) < DateValue(smLastPostedDate) Then
                    mMousePointer vbDefault
                    gMsgBox "Off Air Date Must be After the Last Posted Date " & smLastPostedDate, vbOKOnly
                    If imTabIndex = TABMAIN Then
                        If txtOffAirDate.GetEnabled Then
                            txtOffAirDate.SetFocus
                        End If
                    End If
                    mSave = False
                    Exit Function
                End If
            Else
                If DateValue(gAdjYear(Drop)) < DateValue(smLastPostedDate) Then
                    mMousePointer vbDefault
                    gMsgBox "Drop Date Must be After the Last Posted Date " & smLastPostedDate, vbOKOnly
                    If imTabIndex = TABMAIN Then
                        If txtDropDate.Visible And txtDropDate.GetEnabled Then
                            txtDropDate.SetFocus
                        End If
                    End If
                    mSave = False
                    Exit Function
                End If
            End If
        End If
        If DateValue(gAdjYear(OnAir)) > DateValue(smSvOnAirDate) Then
            If DateValue(smSvOnAirDate) < DateValue(smLastPostedDate) Then
                mMousePointer vbDefault
                gMsgBox "On Air Date can not be advanced as Dates previously Posted through " & smLastPostedDate, vbOKOnly
                If imTabIndex = TABMAIN Then
                    If txtOnAirDate.GetEnabled Then
                        txtOnAirDate.SetFocus
                    End If
                End If
                mSave = False
                Exit Function
            End If
        End If
    End If
    
    If txtLdMult.Text = "" Then
        iLoad = 1
    Else
        iLoad = Val(txtLdMult.Text)
    End If
    
    If txtNoCDs.Text = "" Then
        iNoCDs = 1
    Else
        iNoCDs = Val(txtNoCDs.Text)
    End If
    
    sACName = Trim$(txtACName.Text)
    
    sACPhone = Trim$(txtACPhone.Text)
    If (StrComp(smShttACName, sACName, 1) = 0) And (StrComp(smShttACPhone, sACPhone, 1) = 0) Then
        sACName = ""
        sACPhone = ""
    Else
        sACName = gFixQuote(sACName)
    End If
    sComment = Trim$(txtComments.Text)
    sComment = gFixQuote(sComment)
    slRadarClearType = "N"
    If optRadarClearType(0).Value Then
        slRadarClearType = "C"
    ElseIf optRadarClearType(1).Value Then
        slRadarClearType = "P"
    ElseIf optRadarClearType(3).Value Then
        slRadarClearType = "E"
    End If
    
    slVoiceTracked = "N"
    If optVoiceTracked(0).Value Then
        slVoiceTracked = "Y"
    End If
    llXDReceiverID = Val(txtXDReceiverID.Text)
    slIDCReceiverID = Trim$(txtIDCReceiverID.Text)
    
    ilNoAirPlays = Val(edcNoAirPlays.Text)
    '11/4/11: Clear load factor if air plays defined.  This would only happen with old agreements where
    '         the load factor was defined prior to v6.0
    If ilNoAirPlays > 1 Then
        iLoad = 1
    End If
    '6/28/18
    'If mDateOverlap(lmAttCode, DateValue(gAdjYear(OnAir)), DateValue(gAdjYear(OffAir)), DateValue(gAdjYear(Drop))) Then
    If mDateOverlap(lmAttCode, DateValue(gAdjYear(OnAir)), DateValue(gAdjYear(OffAir)), DateValue(gAdjYear(Drop)), True) Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If
    
    '6/28/18: Moved to after all questions have been asked about the save since this routine updates files.
    'If Not mAdjOverlapAgmnts(DateValue(OnAir), DateValue(OffAir), DateValue(Drop)) Then
    '    mMousePointer vbDefault
    '    mSave = False
    '    Exit Function
    'End If
    
    'Only compare agreements if the multicast y/n radio button is set to yes
'    If rbcMulticast(0).Value Then
'        For ilLoop = 0 To UBound(tmAssocStnMulticastInfo) - 1 Step 1
'            If tmAssocStnMulticastInfo(ilLoop).iSelected = 1 Then
'                SQLQuery = "SELECT attCode, attMulticast from att WHERE AttCode = " & tmAssocStnMulticastInfo(ilLoop).lAttCode
'                'SQLQuery = "SELECT attCode, attMulticast from att WHERE attShfCode = " & tmAssocStnMulticastInfo(ilLoop).iShttCode & " And attVefCode = " & imVefCode
'                'SQLQuery = SQLQuery & " AND attDropDate >= " & "'" & Format$(Now(), sgSQLDateForm) & "'"
'                'SQLQuery = SQLQuery & " AND ((attOnAir <= " & "'" & Format$(Now(), sgSQLDateForm) & "') OR (attOnAir = " & "'" & Format(OnAir, sgSQLDateForm) & "'))"
'                Set tmp_rst = gSQLSelectCall(SQLQuery)
'                If Not tmp_rst.EOF Then
'                    llAttCode = tmp_rst!attCode
'                Else
'                    gMsgBox "No matching agreement was found to multicast with.", vbCritical
'                    Exit Function
'                End If
'                ilMultiOK = True
'                ilRet = gCompare2AgrmntsByPostAndLogType(llAttCode)
'                If Not ilRet Then
'                    slStr = gGetCallLettersByAttCode(lmAttCode)
'                    slStr2 = gGetCallLettersByAttCode(llAttCode)
'                    If ilRet = 1 Then
'                        gMsgBox "Station " & slStr & " Post Type is different from Station " & slStr2 & "."
'                        mSave = False
'                        Exit Function
'                    End If
'                    If ilRet = 2 Then
'                        gMsgBox "Station " & slStr & " Log Type is different from Station " & slStr2 & "."
'                        mSave = False
'                        Exit Function
'                    End If
'                    If ilRet = 3 Then
'                        gMsgBox "Station " & slStr & " CP Control is different from Station " & slStr2 & "."
'                        mSave = False
'                        Exit Function
'                    End If
'                End If
'                slRetStr = gCompare2AgrmntsPledges(llAttCode, imVefCode)
'                If slRetStr <> "" And slRetStr <> "Error" Then
'                    slStr = gGetCallLettersByAttCode(llAttCode)
'                    gMsgBox "Compared to station " & slStr & slRetStr
'                    ilMultiOK = False
'                End If
'                If slRetStr = "Error" Then
'                    mSave = False
'                    Exit Function
'                End If
'                If Not ilMultiOK Then
'                    ilRet = gMsgBox("Conflict(s) exist with this agreement and the agreement(s) that you are trying to form a Multicast with.  Do you want to Save this agreement and have Multicast set back to NO? " & vbCrLf & vbCrLf & "If you answer YES you can adjust the Multicast information and try saving again later.", vbYesNo)
'                    If ilRet = vbYes Then
'                        slMulticast = "N"
'                        rbcMulticast(1).Value = True
'                    Else
'                        mSave = False
'                        rbcMulticast(1).Value = True
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next ilLoop
'    End If
    
    slMulticast = "N"
    imRepopAgnmt = "N"
    slAddStr = ""
    slChgStr = ""
    blShowMulticastMsg = False
    'If grdMulticast.Visible Then
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                slMulticast = "Y"
                'Force repopultation of agreement
                imRepopAgnmt = "Y"
                slDates = grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX)
                '8/12/16: Separated Pledge and Estimate
                If (Left$(grdMulticast.TextMatrix(llRow, MCWITHINDEX), 2) = "No") Or (bmPledgeDataChgd) Or (bmETDataChgd) Or (Left$(grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX), 2) = "No") Or (IsAgmntDirty = False) Then
                    blShowMulticastMsg = True
                End If
                If (Left$(grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX), 2) = "No") Then
                    slAddStr = slAddStr & " " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
                '8/12/16: Separated Pledge and Estimate
                ElseIf (Left$(grdMulticast.TextMatrix(llRow, MCWITHINDEX), 2) = "No") Or (bmPledgeDataChgd) Or (bmETDataChgd) Then
                    slChgStr = slChgStr & " " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
                End If
                'Check That Export type and Export To match
                llMCAttCode = Val(grdMulticast.TextMatrix(llRow, MCATTCODEINDEX))
                '3/21/14: Multicast failing because attExportCBS not in SQL call
                'SQLQuery = "SELECT attExportType, attExportToWeb, attWebInterface, attExportToUnivision, attExportToMarketron FROM ATT WHERE attCode = " & llMCAttCode
                SQLQuery = "SELECT * FROM ATT WHERE attCode = " & llMCAttCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    If ilExportType <> rst!attExportType Then
                        mMousePointer vbDefault
                        gMsgBox "'Affidavit Control' setting on Delivery tab does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
                        mSave = False
                        Exit Function
                    End If
'7701
'                    If ((slExportToWeb = "Y") And (Trim$(rst!attExportToWeb) <> "Y")) Or ((slExportToWeb <> "Y") And (Trim$(rst!attExportToWeb) = "Y")) Then
'                        mMousePointer vbDefault
'                        gMsgBox "Export To: 'Network Web Site' on Delivery tab setting does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
'                        mSave = False
'                        Exit Function
'                    End If
'                    If ((slWebInterface = "C") And (Trim$(rst!attWebInterface) <> "C")) Or ((slWebInterface <> "C") And (Trim$(rst!attWebInterface) = "C")) Then
'                        mMousePointer vbDefault
'                        gMsgBox "Export To: 'Cumulus' on Delivery tab setting does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
'                        mSave = False
'                        Exit Function
'                    End If
'                    If ((slExportToUnivision = "Y") And (Trim$(rst!attExportToUnivision) <> "Y")) Or ((slExportToUnivision <> "Y") And (Trim$(rst!attExportToUnivision) = "Y")) Then
'                        mMousePointer vbDefault
'                        gMsgBox "Export To: 'Univision' on Delivery tab setting does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
'                        mSave = False
'                        Exit Function
'                    End If
'                    If ((slExportToMarketron = "Y") And (Trim$(rst!attExportToMarketron) <> "Y")) Or ((slExportToMarketron <> "Y") And (Trim$(rst!attExportToMarketron) = "Y")) Then
'                        mMousePointer vbDefault
'                        gMsgBox "Export To: 'Marketron' on Delivery tab setting does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
'                        mSave = False
'                        Exit Function
'                    End If
'                    '6592 CBS
'                    If ((slExportToCBS = "Y") And (Trim$(rst!attExportToCBS) <> "Y")) Or ((slExportToCBS <> "Y") And (Trim$(rst!attExportToCBS) = "Y")) Then
'                        mMousePointer vbDefault
'                        gMsgBox "Export To: 'CBS' on Delivery tab setting does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
'                        mSave = False
'                        Exit Function
'                    End If
                End If
                '7701
                blFoundOne = False
                ilCount = 0
                SQLQuery = "SELECT * FROM VAT_Vendor_Agreement WHERE vatAttCode = " & llMCAttCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    'is multi choice not in current?  Stop the save
                    Do While Not rst.EOF
                        If Not mMultiListIsData(rst!vatwvtvendorid, lbcLogDelivery) And Not mMultiListIsData(rst!vatwvtvendorid, lbcAudioDelivery) Then
                            blFoundOne = True
                            Exit Do
                        End If
                        ilCount = ilCount + 1
                        rst.MoveNext
                    Loop
                    'is the count of multi to current not the same? Stop the save!
                    If Not blFoundOne Then
                        If ilCount <> UBound(ilLogVendor) + UBound(ilAudioVendor) Then
                            blFoundOne = True
                        End If
                    End If
                    'for the multi 7701
                    If blFoundOne Then
                        mSetMultilist lbcLogDelivery, 0
                        mSetMultilist lbcAudioDelivery, 0
                        rst.MoveFirst
                        Do While Not rst.EOF
                            If Not mSetMultilist(lbcLogDelivery, rst!vatwvtvendorid) Then
                                mSetMultilist lbcAudioDelivery, rst!vatwvtvendorid
                            End If
                           rst.MoveNext
                        Loop
                        mMousePointer vbDefault
                        gMsgBox "Delivery selection does not match " & Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX)) & ". Terminate and add new Multicast Agreements with matching delivery selections"
                        mSave = False
                        Exit Function
                    End If
                End If
            End If
        Next llRow
    'End If
    
    If (slMulticast = "Y") And (blShowMulticastMsg) Then
        ilShttIndex = gBinarySearchStationInfoByCode(imShttCode)
        If smSvMulticast <> "Y" Then
            If ilShttIndex <> -1 Then
                slAddStr = slAddStr & " " & Trim$(tgStationInfoByCode(ilShttIndex).sCallLetters)
            End If
        '8/12/16: Separated Pledge and Estimate
        ElseIf bmPledgeDataChgd Or bmETDataChgd Then
            If ilShttIndex <> -1 Then
                slChgStr = slChgStr & " " & Trim$(tgStationInfoByCode(ilShttIndex).sCallLetters)
            End If
        End If
        slStr = ""
        If (slAddStr <> "") And (slChgStr <> "") Then
            slStr = "The Following Station Agreement will be Added: " & slAddStr
            slStr = slStr & sgCR & sgLF & "The Following Station Agreement will be Updated: " & slChgStr
        ElseIf slAddStr <> "" Then
            slStr = "The Following Station Agreement will be Added: " & slAddStr
        ElseIf (slChgStr <> "") Then
            slStr = "The Following Station Agreement will be Updated: " & slChgStr
        End If
        If slStr <> "" Then
            mMousePointer vbDefault
            ilRet = gMsgBox("Multicast defined as follows: " & sgCR & sgLF & Trim$(slStr) & ", Continue with Save", vbYesNo)
            If ilRet = vbNo Then
                mSave = False
                Exit Function
            End If
            mMousePointer vbHourglass
        End If
    End If
    If sgVehProgStartTime = "" Then
        slCDStartTime = ""
        ilRet = gDetermineAgreementTimes(imShttCode, imVefCode, OnAir, OffAir, Drop, slCDStartTime, sgVehProgStartTime, sgVehProgEndTime)
        lacPrgTimes.Caption = ""
        If sgVehProgStartTime <> "" Then
            sgVehProgStartTime = gCompactTime(sgVehProgStartTime)
            lacPrgTimes.Caption = "Program Times: " & sgVehProgStartTime
            If sgVehProgEndTime <> "" Then
                sgVehProgEndTime = gCompactTime(sgVehProgEndTime)
                lacPrgTimes.Caption = lacPrgTimes.Caption & "-" & sgVehProgEndTime
            End If
        End If
    End If
    
    '6/28/18: Moved code from above because it modifies files
    If Not mAdjOverlapAgmnts(DateValue(OnAir), DateValue(OffAir), DateValue(Drop)) Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If
    
    '4/3/19
    If ckcExcludeFillSpot.Value = vbChecked Then
        slExcludeFillSpot = "Y"
    Else
        slExcludeFillSpot = "N"
    End If
    If ckcExcludeCntrTypeQ.Value = vbChecked Then
        slExcludeCntrTypeQ = "Y"
    Else
        slExcludeCntrTypeQ = "N"
    End If
    If ckcExcludeCntrTypeR.Value = vbChecked Then
        slExcludeCntrTypeR = "Y"
    Else
        slExcludeCntrTypeR = "N"
    End If
    If ckcExcludeCntrTypeT.Value = vbChecked Then
        slExcludeCntrTypeT = "Y"
    Else
        slExcludeCntrTypeT = "N"
    End If
    If ckcExcludeCntrTypeM.Value = vbChecked Then
        slExcludeCntrTypeM = "Y"
    Else
        slExcludeCntrTypeM = "N"
    End If
    If ckcExcludeCntrTypeS.Value = vbChecked Then
        slExcludeCntrTypeS = "Y"
    Else
        slExcludeCntrTypeS = "N"
    End If
    If ckcExcludeCntrTypeV.Value = vbChecked Then
        slExcludeCntrTypeV = "Y"
    Else
        slExcludeCntrTypeV = "N"
    End If
    
    
    '5457 5589(if xdId changed, set flag to true to update below) Dan M 7375 changed name from mXDigitalIDChanged because now test 'by isci' or 'by break
    mXDigitalChange
    '9452
    If bmSendNotCarriedChange And slSendNotCarried = "Y" Then
            mSetCpttAstStatus lmAttCode, imVefCode, OnAir
    End If

        ''D.S. 8/2/05
        'llTemp = gFindAttHole()
        'If llTemp = -1 Then
        '    mSave = False
        '    mMousePointer vbDefault
        '    Exit Function
        'End If
    SQLQuery = "INSERT INTO att(attCode, attShfCode, attVefCode, attAgreeStart, "
    SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, attSigned, attSignDate, "
    SQLQuery = SQLQuery & "attLoad, attTimeType, attComp, attBarCode, attDropDate, "
    SQLQuery = SQLQuery & "attUsfCode, attEnterDate, attEnterTime, attNotice, "
    SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, attACName, "
    SQLQuery = SQLQuery & "attACPhone, attGenLog, attGenCP, attPostingType, attPrintCP, "
    SQLQuery = SQLQuery & "attExportType, attLogType, attPostType, attWebPW, attWebEmail, "
    '7-6-09 NCR flags
    SQLQuery = SQLQuery & "attSendLogEMail, attSuppressNotice, attLabelID, attLabelShipInfo, "
    SQLQuery = SQLQuery & "attComments, attGenOther, attStartTime, attMulticast, attRadarClearType, "
    SQLQuery = SQLQuery & "attArttCode, attStatus, attNCR, attFormerNCR, attForbidSplitLive, attXDReceiverID, attVoiceTracked, attMonthlyWebPost, attWebInterface, "
    SQLQuery = SQLQuery & "attContractPrinted, "
    SQLQuery = SQLQuery & "attMktRepUstCode, "
    SQLQuery = SQLQuery & "attServRepUstCode, "
    SQLQuery = SQLQuery & "attVehProgStartTime, "
    SQLQuery = SQLQuery & "attVehProgEndTime, "
    SQLQuery = SQLQuery & "attExportToWeb, "
    SQLQuery = SQLQuery & "attExportToUnivision, "
    SQLQuery = SQLQuery & "attExportToMarketron, "
    SQLQuery = SQLQuery & "attExportToCBS, "
    SQLQuery = SQLQuery & "attExportToClearCh, "
    SQLQuery = SQLQuery & "attPledgeType, "
    SQLQuery = SQLQuery & "attNoAirPlays, "
    SQLQuery = SQLQuery & "attDesignVersion, "
    '6466
    SQLQuery = SQLQuery & "attIDCReceiverID, attIDCGroupType, "
    SQLQuery = SQLQuery & "attMktronActiveDate, "
    SQLQuery = SQLQuery & "attSentToXDSStatus, "
    SQLQuery = SQLQuery & "attAudioDelivery, "
    SQLQuery = SQLQuery & "attExportToJelli, "
    '3/23/15: Add Send Delays to XDS
    SQLQuery = SQLQuery & "attSendDelayToXDS, "
    SQLQuery = SQLQuery & "attXDSSendNotCarry, "
    SQLQuery = SQLQuery & "attServiceAgreement, "
    '4/3/19
    SQLQuery = SQLQuery & "attExcludeFillSpot, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeQ, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeR, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeT, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeM, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeS, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeV, "
    
    SQLQuery = SQLQuery & "attUnused "
    SQLQuery = SQLQuery & ")"
    SQLQuery = SQLQuery & " VALUES"
    SQLQuery = SQLQuery & "(" & "ReplaceAtt" & ", " & "ReplaceShtt" & ", " & imVefCode & ", '" & Format$(AgreeStart, sgSQLDateForm) & "', '" & Format$(AgreeEnd, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(OnAir, sgSQLDateForm) & "', '" & Format$(OffAir, sgSQLDateForm) & "', " & iSigned & ", "
    SQLQuery = SQLQuery & "'" & Format$(RetDate, sgSQLDateForm) & "', " & iLoad & ", " & iTimeType & ", "
    SQLQuery = SQLQuery & iComp & ", " & iBarCodes & ", '" & Format$(Drop, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & igUstCode & ", '" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format$(CurTime, sgSQLTimeForm) & "', '" & txtDays.Text & "', "
    SQLQuery = SQLQuery & iCarryCmml & ", " & iNoCDs & ", " & iSendTape & ", '" & sACName & "', "
    SQLQuery = SQLQuery & "'" & sACPhone & "', '" & txtLog.Text & "', '" & txtCP.Text & "', " & iPost & ", " & iPrintCP & ", "
    SQLQuery = SQLQuery & ilExportType & ", " & ilLogType & ", " & ilPostType & ", '" & slPassword & "', '" & slEmailAddr & "', "
    'SQLQuery = SQLQuery & "'" & sComment & "', '" & txtOther.Text & "')"
    '7-6-09 NCR flags
    'SQLQuery = SQLQuery & ilSendLogEMail & ", '" & sSuppressNotices & "', '" & slLabelID & "', '" & slShipInfo & "', '" & sComment & "', '" & txtOther.text & "', '" & Format$(sgCDStartTime, sgSQLTimeForm) & "', '" & slMulticast & "', '" & slRadarClearType & "', " & ilAffAE & ", '" & sNCR & "', '" & sFormerNCR & "', '" & slForbidSplitLive & "' )"
    SQLQuery = SQLQuery & ilSendLogEMail & ", '" & sSuppressNotices & "', '" & slLabelID & "', '" & slShipInfo & "', "
    SQLQuery = SQLQuery & "'" & sComment & "', '" & txtOther.Text & "', '" & Format$(sgCDStartTime, sgSQLTimeForm) & "', '" & slMulticast & "', '" & slRadarClearType & "', "
    SQLQuery = SQLQuery & llAffAE & ", 'C'" & ", '" & sNCR & "', '" & sFormerNCR & "', '" & slForbidSplitLive & "', " & llXDReceiverID & ", '" & slVoiceTracked & "', '" & slMonthlyWebPost & "', '" & slWebInterface & "', "
    SQLQuery = SQLQuery & "'" & "N" & "', "
    SQLQuery = SQLQuery & ilMktRepUstCode & ", "
    SQLQuery = SQLQuery & ilServRepUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(sgVehProgStartTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(sgVehProgEndTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & slExportToWeb & "', "
    SQLQuery = SQLQuery & "'" & slExportToUnivision & "', "
    SQLQuery = SQLQuery & "'" & slExportToMarketron & "', "
    'CBS 6592
    SQLQuery = SQLQuery & "'" & slExportToCBS & "', "
    SQLQuery = SQLQuery & "'" & "N" & "', "
    SQLQuery = SQLQuery & "'" & slPledgeType & "', "
    SQLQuery = SQLQuery & ilNoAirPlays & ", "
    SQLQuery = SQLQuery & 2 & ", "
    '6466
    SQLQuery = SQLQuery & "'" & slIDCReceiverID & "', '" & slIDCGroup & "', "
    SQLQuery = SQLQuery & "'" & Format$(smMktronActiveDate, sgSQLDateForm) & "', "
    If IsAgmntDirty = False Then
        SQLQuery = SQLQuery & "'" & "N" & "', "
    Else
        SQLQuery = SQLQuery & "'" & "M" & "', "
    End If
    SQLQuery = SQLQuery & "'" & slAudioDelivery & "', "
    SQLQuery = SQLQuery & "'" & slExportToJelli & "', "
    '3/23/15: Add Send Delays to XDS
    SQLQuery = SQLQuery & "'" & slSendDelays & "', "
    SQLQuery = SQLQuery & "'" & slSendNotCarried & "', "
    SQLQuery = SQLQuery & "'" & slService & "', "
    '4-3-19
    SQLQuery = SQLQuery & "'" & slExcludeFillSpot & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeQ & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeR & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeT & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeM & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeS & "', "
    SQLQuery = SQLQuery & "'" & slExcludeCntrTypeV & "', "
    
    SQLQuery = SQLQuery & "'" & "" & "'"
    SQLQuery = SQLQuery & ")"
    
    slAttSQLQuery = SQLQuery
    'Adding a new agreement
    If IsAgmntDirty = False Then
        SQLQuery = Replace(slAttSQLQuery, "ReplaceShtt", Trim$(Str$(imShttCode)), , , vbTextCompare)

        If iAsk Then
            mMousePointer vbDefault
            If gMsgBox("Save all changes?", vbYesNo) = vbNo Then
                mSave = False
                Exit Function
            End If
            mMousePointer vbHourglass
        End If
        lmAttCode = gInsertAndReturnCode(SQLQuery, "att", "attCode", "ReplaceAtt")
        If lmAttCode <= 0 Then
            mSave = False
            Exit Function
        End If
        iNewRec = True
    'Or updating an existing station's data
    Else
        mMousePointer vbDefault
        If mCPDelete(OnAir, OffAir, Drop) = False Then
            mSave = False
            mMousePointer vbDefault
            Exit Function
        End If
        mMousePointer vbHourglass
        SQLQuery = "UPDATE att SET "
        SQLQuery = SQLQuery & "attShfCode = " & imShttCode & ", "
        SQLQuery = SQLQuery & "attVefCode = " & imVefCode & ", "
        SQLQuery = SQLQuery & "attAgreeStart = '" & Format$(AgreeStart, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attAgreeEnd = '" & Format$(AgreeEnd, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attOnAir = '" & Format$(OnAir, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attOffAir = '" & Format$(OffAir, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attSigned = " & iSigned & ", "
        SQLQuery = SQLQuery & "attSignDate = '" & Format$(RetDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attLoad = " & iLoad & ", "
        SQLQuery = SQLQuery & "attTimeType = " & iTimeType & ", "
        SQLQuery = SQLQuery & "attComp = " & iComp & ", "
        SQLQuery = SQLQuery & "attStartTime = '" & Format$(sgCDStartTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "attBarCode = " & iBarCodes & ", "
        SQLQuery = SQLQuery & "attDropDate = '" & Format$(Drop, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attUsfCode = " & igUstCode & ", "
        SQLQuery = SQLQuery & "attEnterDate = '" & Format$(smCurDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attEnterTime = '" & Format$(CurTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "attNotice = '" & txtDays.Text & "', "
        SQLQuery = SQLQuery & "attCarryCmml = " & iCarryCmml & ", "
        SQLQuery = SQLQuery & "attNoCDs = " & iNoCDs & ", "
        SQLQuery = SQLQuery & "attSendTape = " & iSendTape & ", "
        SQLQuery = SQLQuery & "attACName = '" & sACName & "', "
        SQLQuery = SQLQuery & "attACPhone = '" & sACPhone & "', "
        SQLQuery = SQLQuery & "attGenLog = '" & txtLog.Text & "', "
        SQLQuery = SQLQuery & "attGenCP = '" & txtCP.Text & "', "
        SQLQuery = SQLQuery & "attPostingType = " & iPost & ", "
        SQLQuery = SQLQuery & "attPrintCP = " & iPrintCP & ", "
        SQLQuery = SQLQuery & "attComments = '" & sComment & "', "
        SQLQuery = SQLQuery & "attGenOther = '" & txtOther.Text & "', "
        SQLQuery = SQLQuery & "attExportType = " & ilExportType & ", "
        SQLQuery = SQLQuery & "attLogType = " & ilLogType & ", "
        SQLQuery = SQLQuery & "attPostType = " & ilPostType & ", "
        SQLQuery = SQLQuery & "attSendLogEMail = " & ilSendLogEMail & ", "
        SQLQuery = SQLQuery & "attWebPW = '" & slPassword & "', "
        SQLQuery = SQLQuery & "attWebEmail = '" & slEmailAddr & "', "
        SQLQuery = SQLQuery & "attMulticast = '" & slMulticast & "', "
        SQLQuery = SQLQuery & "attRadarClearType = '" & slRadarClearType & "', "
        SQLQuery = SQLQuery & "attArttCode = " & llAffAE & ", "
        '7-6-09 NCR flags
        SQLQuery = SQLQuery & "attSuppressNotice = '" & sSuppressNotices & "', "
        SQLQuery = SQLQuery & "attLabelID = '" & slLabelID & "', "
        SQLQuery = SQLQuery & "attLabelShipInfo = '" & slShipInfo & "',"
        SQLQuery = SQLQuery & "attNCR = '" & sNCR & "',"
        '7-6-09 NCR flags
        SQLQuery = SQLQuery & "attFormerNCR = '" & sFormerNCR & "', "
        SQLQuery = SQLQuery & "attForbidSplitLive = '" & slForbidSplitLive & "', "
        SQLQuery = SQLQuery & "attXDReceiverID = '" & llXDReceiverID & "', "
        SQLQuery = SQLQuery & "attVoiceTracked = '" & slVoiceTracked & "', "
        SQLQuery = SQLQuery & "attMonthlyWebPost = '" & slMonthlyWebPost & "', "
        SQLQuery = SQLQuery & "attWebInterface = '" & slWebInterface & "', "
        SQLQuery = SQLQuery & "attContractPrinted ='" & "N" & "', "
        SQLQuery = SQLQuery & "attMktRepUstCode = " & ilMktRepUstCode & ", "
        SQLQuery = SQLQuery & "attServRepUstCode = " & ilServRepUstCode & ", "
        SQLQuery = SQLQuery & "attVehProgStartTime = '" & Format$(sgVehProgStartTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "attVehProgEndTime = '" & Format$(sgVehProgEndTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "attExportToWeb = '" & slExportToWeb & "', "
        SQLQuery = SQLQuery & "attExportToUnivision = '" & slExportToUnivision & "', "
        SQLQuery = SQLQuery & "attExportToMarketron = '" & slExportToMarketron & "', "
        '6592
        SQLQuery = SQLQuery & "attExportToCBS = '" & slExportToCBS & "', "
        SQLQuery = SQLQuery & "attExportToClearCh = '" & "N" & "', "
        SQLQuery = SQLQuery & "attPledgeType = '" & slPledgeType & "', "
        SQLQuery = SQLQuery & "attNoAirPlays = " & ilNoAirPlays & ", "
        '6466
        SQLQuery = SQLQuery & "attIDCReceiverID = '" & slIDCReceiverID & "', attIDCGroupType = '" & slIDCGroup & "', "
        'SQLQuery = SQLQuery & "attIDCReceiverID = '" & slIDCReceiverID & "', "
        SQLQuery = SQLQuery & "attMktronActiveDate = '" & Format$(smMktronActiveDate, sgSQLDateForm) & "', "
        '5589
        If bmXDIdChanged Or imDateChgd Then
            SQLQuery = SQLQuery & "attSentToXDSStatus = 'M', "
        End If
        SQLQuery = SQLQuery & "attAudioDelivery = '" & slAudioDelivery & "', "
        SQLQuery = SQLQuery & "attExportToJelli = '" & slExportToJelli & "', "
        '3/23/15: Add Send Delays to XDS
        SQLQuery = SQLQuery & "attSendDelayToXDS = '" & slSendDelays & "', "
        SQLQuery = SQLQuery & "attXDSSendNotCarry = '" & slSendNotCarried & "', "
        SQLQuery = SQLQuery & "attServiceAgreement = '" & slService & "', "
        '4/3/19
        SQLQuery = SQLQuery & "attExcludeFillSpot = '" & slExcludeFillSpot & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeQ = '" & slExcludeCntrTypeQ & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeR = '" & slExcludeCntrTypeR & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeT = '" & slExcludeCntrTypeT & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeM = '" & slExcludeCntrTypeM & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeS = '" & slExcludeCntrTypeS & "', "
        SQLQuery = SQLQuery & "attExcludeCntrTypeV = '" & slExcludeCntrTypeV & "', "
             
        SQLQuery = SQLQuery & "attUnused = '" & "" & "'"
        SQLQuery = SQLQuery & " WHERE attCode = " & lmAttCode & ""
        If iAsk Then
            mMousePointer vbDefault
            If gMsgBox("Save all changes?", vbYesNo) = vbNo Then
                mSave = False
                mMousePointer vbDefault
                Exit Function
            End If
            mMousePointer vbHourglass
        End If
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSave"
            mSave = False
            Exit Function
        End If
        iNewRec = False
    End If
    '7701
    If Not gSaveVATMulti(lmAttCode, ilLogVendor, ilAudioVendor) Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If
'    If (imDatLoaded) And (sgUstPledge = "Y") Then
'        'Delete Dayparts or Avails
'        SQLQuery = "DELETE FROM dat"
'        SQLQuery = SQLQuery + " WHERE (datAtfCode = " & lmAttCode & ")"
'        'cnn.Execute SQLQuery, rdExecDirect
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'        'Since DAT are removed, remove the matching ept
'        SQLQuery = "DELETE FROM ept"
'        SQLQuery = SQLQuery + " WHERE (eptAttCode = " & lmAttCode & ")"
'        'cnn.Execute SQLQuery, rdExecDirect
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'        For iLoop = 0 To UBound(tgDat) - 1 Step 1
'            sFdStTime = Format$(tgDat(iLoop).sFdSTime, sgShowTimeWSecForm)
'            sFdEdTime = Format$(tgDat(iLoop).sFdETime, sgShowTimeWSecForm)
'            If Len(Trim$(tgDat(iLoop).sPdSTime)) = 0 Or Asc(tgDat(iLoop).sPdSTime) = 0 Then
'                sPdStTime = sFdStTime
'            Else
'                sPdStTime = Format$(tgDat(iLoop).sPdSTime, sgShowTimeWSecForm)
'            End If
'            If Len(Trim$(tgDat(iLoop).sPdETime)) = 0 Or Asc(tgDat(iLoop).sPdETime) = 0 Then
'                sPdEdTime = sPdStTime
'            Else
'                sPdEdTime = Format$(tgDat(iLoop).sPdETime, sgShowTimeWSecForm)
'            End If
'            'To avoid duplicate key when two saves are done in a row.
'            'always set lCode to zero (0).  This should not be required because of the 'Delete From'
'            'If IsDirty = False Then
'                tgDat(iLoop).lCode = 0
'            'End If
'            'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
'            SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
'            SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
'            SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
'            SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
'            SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime)"
'            'SQLQuery = SQLQuery & " VALUES (" & tgDat(iLoop).lCode & ", " & lmAttCode & ", " & imShttCode & ", " & imVefCode
'            SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & lmAttCode & ", " & imShttCode & ", " & imVefCode
'
'            'If optTimeType(0).Value Or optTimeType(1).Value Or optTimeType(2).Value Then
'            '    If optTimeType(0).Value Then       'Live Dayparts
'            '        SQLQuery = SQLQuery & ",0,"
'            '    ElseIf optTimeType(1).Value Then   'Avails
'            '        SQLQuery = SQLQuery & ",1,"
'            '    ElseIf optTimeType(2).Value Then   'CD/Tape Dayparts
'            '        SQLQuery = SQLQuery & ",2,"
'            '    End If
'            'Else
'            '    SQLQuery = SQLQuery & "," & tgDat(0).iDACode & ","
'            'End If
'            SQLQuery = SQLQuery & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(0) & ", " & tgDat(iLoop).iFdDay(1) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(2) & ", " & tgDat(iLoop).iFdDay(3) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(4) & ", " & tgDat(iLoop).iFdDay(5) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(6) & ", "
'            SQLQuery = SQLQuery & "'" & Format$(sFdStTime, sgSQLTimeForm) & "','" & Format$(sFdEdTime, sgSQLTimeForm) & "',"
'            SQLQuery = SQLQuery & tgDat(iLoop).iFdStatus & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(0) & ", " & tgDat(iLoop).iPdDay(1) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(2) & ", " & tgDat(iLoop).iPdDay(3) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(4) & ", " & tgDat(iLoop).iPdDay(5) & ","
'            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(6) & ", "
'
'            If Asc(tgDat(iLoop).sPdDayFed) = 0 Then
'                SQLQuery = SQLQuery & "'" & " " & "', "
'            Else
'                SQLQuery = SQLQuery & "'" & tgDat(iLoop).sPdDayFed & "', "
'            End If
'
'
'
'            SQLQuery = SQLQuery & "'" & Format$(sPdStTime, sgSQLTimeForm) & "','" & Format$(sPdEdTime, sgSQLTimeForm) & "',"
'            SQLQuery = SQLQuery & tgDat(iLoop).iAirPlayNo & ", "
'            SQLQuery = SQLQuery & "'" & tgDat(iLoop).sEstimatedTime & "')"
'            ''cnn.Execute SQLQuery, rdExecDirect
'            'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            '    GoSub ErrHand:
'            'End If
'            llCode = gInsertAndReturnCode(SQLQuery, "dat", "datCode", "Replace")
'            If llCode > 0 Then
'                tgDat(iLoop).lCode = llCode
'            Else
'                mSave = False
'                Exit Function
'            End If
'            '5/22/07:  No Required as DAT is deleted (never updated)
'            'SQLQuery = "SELECT MAX(datCode) from dat"
'            'Set rst = gSQLSelectCall(SQLQuery)
'            'If Not rst.EOF Then
'            '    tgDat(iLoop).lCode = rst(0).Value
'            'End If
'        Next iLoop
'    End If
    If Not mSavePledgeInfo(lmAttCode, imShttCode, imVefCode) Then
        mMousePointer vbDefault
        mSave = False
        Exit Function
    End If
    ilRet = mSetUsedForAtt(imShttCode, True)
    ilRet = mCheckHistDate()
    ilRet = mAdjCPTT(iNewRec, OnAir, OffAir, Drop)
    If ilRet Then
        'cnn.CommitTrans
    End If
    
    ilRet = mSaveContractPDF(lmAttCode, imShttCode, imVefCode)
    ilRet = mSaveEstimatedInfo(lmAttCode, imShttCode, imVefCode)
    ilRet = mSaveEventInfo(lmAttCode, imShttCode, imVefCode)
    'Save contacts
    udcContactGrid.Action 5
    If gDateValue(OffAir) <= gDateValue(Drop) Then
        ilRet = mUpdateIDCExport(OnAir, OffAir)
    Else
        ilRet = mUpdateIDCExport(OnAir, Drop)
    End If
    
    '2/17/16: Spot Build call only required for new or change agreements without any posting
    If (imOkToChange) Or (smLastPostedDate = "1/1/1970") Or (bmETDataChgd And (bmPledgeDataChgd = False)) Then
        If gDateValue(OffAir) < gDateValue(Drop) Then
            gSetStationSpotBuilder "A", imVefCode, imShttCode, gDateValue(OnAir), gDateValue(OffAir)
        Else
            gSetStationSpotBuilder "A", imVefCode, imShttCode, gDateValue(OnAir), gDateValue(Drop)
        End If
    End If
    
'    '5457 Dan M
'    mXDigitalChange
    '7701 no need to change.
    ' Update the web if this records export type is set to Web. 0 = manual, 1 = web, 2 = Univision
    If (ilExportType = 1) And ((ckcExportTo(0).Value = vbChecked) Or (ckcExportTo(1).Value = vbChecked)) Then
        Dim LogType, PostType As Integer
        LogType = 0
        If frcLogType.Visible = True Then
            If rbcLogType(0).Value Then
                LogType = 0
            ElseIf rbcLogType(1).Value Then
                LogType = 1
            ElseIf rbcLogType(2).Value Then
                LogType = 2
            End If
        End If
        PostType = 0
        If frcPostType.Visible = True Then
            If rbcPostType(0).Value Then
                PostType = 0
            ElseIf rbcPostType(1).Value Then
                PostType = 1
            ElseIf rbcPostType(2).Value Then
                PostType = 2
            End If
        End If
        
        If frcLogType.Visible = True Then
            If smInitialattWebEmail <> txtEmailAddr.Text Or _
                smInitialattWebPW <> txtLogPassword.Text Or _
                smInitialattLogType <> LogType Or _
                smInitialattPostType <> PostType Then
                Call UpdateWebSite
                smInitialattWebEmail = txtEmailAddr.Text
                smInitialattWebPW = txtLogPassword.Text
                smInitialattLogType = LogType
                smInitialattPostType = PostType
            End If
        Else
            If smInitialattWebEmail <> txtEmailAddr.Text Or _
                smInitialattWebPW <> txtLogPassword.Text Then
                'smInitialattWebPW <> txtLogPassword.Text Or_
                'smInitialattLogType <> LogType Or _
                'smInitialattPostType <> PostType Then
                Call UpdateWebSite
                smInitialattWebEmail = txtEmailAddr.Text
                smInitialattWebPW = txtLogPassword.Text
                'smInitialattLogType = LogType
                'smInitialattPostType = PostType
            End If
        End If
    End If
        
    '2/15/11: Doug, more stuff needs to be sent to Web at this point
    If slMulticast = "Y" Then
        'Update other multicast agreements if required
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                slDates = grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX)
                If Left(slDates, 2) <> "No" Then
                    slStr = grdMulticast.TextMatrix(llRow, MCWITHINDEX)
                    llMCAttCode = Val(grdMulticast.TextMatrix(llRow, MCATTCODEINDEX))
                    SQLQuery = "SELECT attOnAir, attOffAir, attDropDate FROM ATT WHERE attCode = " & llMCAttCode
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        'Test Date to determine if agreement needs to be split
                        ilMCShttCode = Val(grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX))
                        llMCOnAir = DateValue(gAdjYear(rst!attOnAir))
                        '6/28/18: Handle case where agreement advanced passed the original date and split not required
                        llMCOffAir = DateValue(gAdjYear(rst!attOffAir))
                        llMCDropDate = DateValue(gAdjYear(rst!attDropDate))
                        If (DateValue(gAdjYear(OnAir)) > llMCOffAir) Or (DateValue(gAdjYear(OnAir)) > llMCDropDate) Then
                            ilRet = mSplitAgreement(True, Val(grdMulticast.TextMatrix(llRow, MCATTCODEINDEX)), DateAdd("d", -1, OnAir), OffAir, CurTime, ilNoAirPlays, llNewMCAttCode)
                            ilSvShttCode = imShttCode
                            imShttCode = ilMCShttCode
                            ilRet = mDateOverlap(llNewMCAttCode, DateValue(gAdjYear(OnAir)), DateValue(gAdjYear(OffAir)), DateValue(gAdjYear(Drop)), False)
                            ilRet = mAdjOverlapAgmnts(DateValue(OnAir), DateValue(OffAir), DateValue(Drop))
                            imShttCode = ilSvShttCode
                        ElseIf DateValue(gAdjYear(OnAir)) > llMCOnAir Then
                            'Split Agreement
                            ilRet = mSplitAgreement(False, Val(grdMulticast.TextMatrix(llRow, MCATTCODEINDEX)), DateAdd("d", -1, OnAir), OffAir, CurTime, ilNoAirPlays, llNewMCAttCode)
                            '6/28/18
                            ilSvShttCode = imShttCode
                            imShttCode = ilMCShttCode
                            ilRet = mDateOverlap(llNewMCAttCode, DateValue(gAdjYear(OnAir)), DateValue(gAdjYear(OffAir)), DateValue(gAdjYear(Drop)), False)
                            ilRet = mAdjOverlapAgmnts(DateValue(OnAir), DateValue(OffAir), DateValue(Drop))
                            imShttCode = ilSvShttCode
                        Else
                            If (DateValue(gAdjYear(OffAir)) <> DateValue(gAdjYear(rst!attOffAir))) Or (DateValue(gAdjYear(Drop)) <> DateValue(gAdjYear(rst!attDropDate))) Then
                                SQLQuery = "UPDATE att SET "
                                SQLQuery = SQLQuery & "attOffAir = '" & Format$(OffAir, sgSQLDateForm) & "', "
                                SQLQuery = SQLQuery & "attDropDate = '" & Format$(Drop, sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery & " WHERE attCode = " & llMCAttCode
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    mMousePointer vbDefault
                                    gHandleError "AffErrorLog.txt", "AffAgmnt-mSave"
                                    mSave = False
                                    Exit Function
                                End If
                            End If
                            If Left(slStr, 3) = "Not" Then
                                'Set station as Multicast
                                SQLQuery = "UPDATE att SET"
                                SQLQuery = SQLQuery & "attMulticast = 'Y' ,"
                                SQLQuery = SQLQuery & "attNoAirPlays = " & ilNoAirPlays
                                SQLQuery = SQLQuery & " WHERE attCode = " & llMCAttCode
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    mMousePointer vbDefault
                                    gHandleError "AffErrorLog.txt", "AffAgmnt-mSave"
                                    mSave = False
                                    Exit Function
                                End If
                                bmPledgeDataChgd = True
                            End If
                            '7701
                            If Not gSaveVATMulti(llMCAttCode, ilLogVendor, ilAudioVendor) Then
                                mMousePointer vbDefault
                                mSave = False
                                Exit Function
                            End If
                            '8/12/16: Added ET test
                            If bmPledgeDataChgd Or bmETDataChgd Then
                                'Update Pledge
                                If Not mSavePledgeInfo(llMCAttCode, ilMCShttCode, imVefCode) Then
                                    mMousePointer vbDefault
                                    mSave = False
                                    Exit Function
                                End If
                                ilRet = mSetUsedForAtt(ilMCShttCode, False)
                                ilRet = mSaveEstimatedInfo(llMCAttCode, ilMCShttCode, imVefCode)
                                ilRet = mSaveEventInfo(llMCAttCode, ilMCShttCode, imVefCode)
                            Else
                                ilRet = mSetUsedForAtt(ilMCShttCode, False)
                            End If
                        End If
                    End If
                ElseIf Left(slDates, 2) = "No" Then
                    'Add agreement
                    ilMCShttCode = Val(grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX))
                    SQLQuery = Replace(slAttSQLQuery, "ReplaceShtt", Trim$(Str$(ilMCShttCode)), , , vbTextCompare)
                    llMCAttCode = gInsertAndReturnCode(SQLQuery, "att", "attCode", "ReplaceAtt")
                    If llMCAttCode <= 0 Then
                        mMousePointer vbDefault
                        mSave = False
                        Exit Function
                    End If
                    If Not mSavePledgeInfo(llMCAttCode, ilMCShttCode, imVefCode) Then
                        mMousePointer vbDefault
                        mSave = False
                        Exit Function
                    End If
                    '7701
                    If Not gSaveVATMulti(llMCAttCode, ilLogVendor, ilAudioVendor) Then
                        mMousePointer vbDefault
                        mSave = False
                        Exit Function
                    End If
                    ilRet = mSetUsedForAtt(ilMCShttCode, False)
                    ilRet = mSaveContractPDF(llMCAttCode, ilMCShttCode, imVefCode)
                    ilRet = mSaveEstimatedInfo(llMCAttCode, ilMCShttCode, imVefCode)
                    ilRet = mSaveEventInfo(llMCAttCode, ilMCShttCode, imVefCode)
                    '11/19/13: Create CPTT for multi-cast (ttp 5385)
                    llSvAttCode = lmAttCode
                    ilSvShttCode = imShttCode
                    lmAttCode = llMCAttCode
                    imShttCode = ilMCShttCode
                    ilRet = mAdjCPTT(True, OnAir, OffAir, Drop)
                    lmAttCode = llSvAttCode
                    imShttCode = ilSvShttCode
                End If
                '2/17/16: Spot Build call only required for new or change agreements without any posting
                If (imOkToChange) Or (smLastPostedDate = "1/1/1970") Or (bmETDataChgd And (bmPledgeDataChgd = False)) Then
                    If gDateValue(OffAir) < gDateValue(Drop) Then
                        gSetStationSpotBuilder "A", imVefCode, ilMCShttCode, gDateValue(OnAir), gDateValue(OffAir)
                    Else
                        gSetStationSpotBuilder "A", imVefCode, ilMCShttCode, gDateValue(OnAir), gDateValue(Drop)
                    End If
                End If
            End If
        Next llRow
    End If

    smSvOnAirDate = OnAir
    smSvOffAirDate = OffAir
    smSvDropDate = Drop

    mSave = True
    imFieldChgd = False
    '8/12/16: Separated Pledge and Estimate
    bmPledgeDataChgd = False
    bmETDataChgd = False
    '5457
    mXDigitalContact True
    '9452 stay on form, but make another change?
    bmSendNotCarriedChange = False
    Exit Function
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "", "Agreement-mSave"
    mSave = False
End Function

Private Sub mXDigitalChange()
    Dim ilVendor As Integer
    Dim slNewXD As String
    
    '7701
    '7375
    slNewXD = mRetrieveMultiListString(lbcAudioDelivery)
    If smPreviousXDAudioDelivery <> slNewXD Then
        If mMultiListIsData(Vendors.XDS_Break, lbcAudioDelivery) Or mMultiListIsData(Vendors.XDS_ISCI, lbcAudioDelivery) Then
            smPreviousXDAudioDelivery = slNewXD
            smPreviousXDReceiver = "Dan changed this value to force an update"
        End If
    End If
'    '7375
'    If smPreviousXDAudioDelivery <> rbcAudio(0).Value & rbcAudio(1).Value Then
'        If rbcAudio(0).Value = True Or rbcAudio(1).Value = True Then
'            smPreviousXDAudioDelivery = rbcAudio(0).Value & rbcAudio(1).Value
'            smPreviousXDReceiver = "Dan changed this value to force an update"
'        End If
'    End If
'ttp 5457
    If smPreviousXDReceiver <> txtXDReceiverID.Text Then
        bmXDIdChanged = True
        '7912
        SQLQuery = "Update VAT_Vendor_Agreement set vatSentToWeb = '' WHERE vatWvtVendorID = " & Vendors.XDS_Break & " AND vatAttCode = " & lmAttCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "AffAgmnt-mXDigitalChange"
            Exit Sub
        End If
        SQLQuery = "UPDATE shtt set shttSentToXDSStatus = 'M' WHERE shttCode = " & imShttCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "AffAgmnt-mXDigitalChange"
            Exit Sub
        End If
        smPreviousXDReceiver = txtXDReceiverID.Text
    Else
        bmXDIdChanged = False
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "mXDigitalChange"
End Sub
Private Function UpdateWebSite() As Boolean
    On Error GoTo ErrHand
    Dim SQLQuery As String
    Dim slStr As String
    Dim slVefName As String
    Dim hmToHeader As Integer
    Dim iRet As Integer
    Dim cprst As ADODB.Recordset
    Dim FTPAddress As String
    Dim slTemp As String
    Dim slFileName As String
    Dim slTemp1 As String
    Dim slSuppressLog As String
    Dim slUseActual As String
    Dim llVpf As Long

    Call gLoadOption(sgWebServerSection, "FTPAddress", FTPAddress)
    
    slTemp1 = gGetComputerName()
    If slTemp1 = "N/A" Then
        slTemp1 = "Unknown"
    End If
    slTemp1 = slTemp1 & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    slFileName = "WebHeaders_" & slTemp1

    frmProgressMsg.Show
    frmProgressMsg.SetMessage 0, "Updating Web Site..." & vbCrLf & vbCrLf & "[" & FTPAddress & "]"
    DoEvents
    mMousePointer vbHourglass
    UpdateWebSite = False
    SQLQuery = "Select vefName From VEF_Vehicles Where vefCode = " & imVefCode & ""
    Set cprst = gSQLSelectCall(SQLQuery)
    If cprst.EOF Then
        mMousePointer vbDefault
        frmProgressMsg.SetMessage 1, "Unable to find the vehicle name for this record." & vbCrLf & "Web site not updated."
        Exit Function
    End If
    slVefName = cprst!vefName
    
    'SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW"
    'SQLQuery = SQLQuery + " FROM shtt, cptt, att"
    SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attMulticast, attMonthlyWebPost"
    'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
    SQLQuery = SQLQuery + " FROM shtt, cptt, att"
    SQLQuery = SQLQuery + " WHERE (ShttCode = " & imShttCode & ""
    SQLQuery = SQLQuery + " AND attCode = " & lmAttCode & ""
    SQLQuery = SQLQuery + " AND cpttatfCode = " & lmAttCode & ")"
    Set cprst = gSQLSelectCall(SQLQuery)
    If cprst.EOF Then
        mMousePointer vbDefault
        frmProgressMsg.SetMessage 1, "Web site not updated.   Please verify that Logs have been generated for this vehicle."
        Exit Function
    End If
    
    '***** Important Note *******
    'If you change the way that the headers are built then you have to make the
    'same changes to the way the header is built in frmStations and in frmWebExportSchdSpot.
    'They both put out headers when either station information is changed or
    'a web export is done.
    smAttWebInterface = Trim$(gGetWebInterface(lmAttCode))
    
    slUseActual = "N"
    llVpf = gBinarySearchVpf(CLng(imVefCode))
    If (llVpf <> -1) Then
        If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
            slUseActual = "Y"
        End If
    End If
    
    slSuppressLog = "N"
    If (llVpf <> -1) Then
        If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And SUPPRESSWEBLOG) = SUPPRESSWEBLOG Then
            slSuppressLog = "Y"
        End If
    End If

    slStr = gBuildWebHeaders(cprst, imVefCode, slVefName, imShttCode, smAttWebInterface, True, "A", "", "", slUseActual, slSuppressLog)
    
    Call gLoadOption(sgWebServerSection, "WebExports", smWebExports)
    smWebExports = gSetPathEndSlash(smWebExports, True)
    sToFileHeader = smWebExports & slFileName
    'hmToHeader = FreeFile
    'iRet = 0
    'Open sToFileHeader For Output Lock Write As hmToHeader
    iRet = gFileOpen(sToFileHeader, "Output Lock Write", hmToHeader)
    If iRet <> 0 Then
        mMousePointer vbDefault
        frmProgressMsg.SetMessage 1, "Unable to open file " & sToFileHeader & vbCrLf & "Web site not updated"
        Exit Function
    End If

    Print #hmToHeader, gBuildWebHeaderDetail()
    'Print #hmToHeader, "attCode , NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, StartTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime"
    Print #hmToHeader, slStr
    Close #hmToHeader
    If Not gFTPFileToWebServer(sToFileHeader, slFileName) Then
        mMousePointer vbDefault
        frmProgressMsg.SetMessage 1, "Unable to update the Web Server." & vbCrLf & "Web site not updated"
        Exit Function
    End If
    If Not gSendCmdToWebServer("ImportHeaders.dll", slFileName) Then
        mMousePointer vbDefault
        frmProgressMsg.SetMessage 1, "FAIL: Unable to instruct Web Server to Import..."
        Exit Function
    End If
    Unload frmProgressMsg
    UpdateWebSite = True
    Exit Function
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmAgmnt-UpdateWebSite"
End Function

Private Sub mSort(optSort As OptionButton, cboCtrl As ComboBox, iCode() As Integer)
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim iFound As Integer
    Dim iTest As Integer
    
    cboCtrl.Clear
    '9/17/11: Reload Stations if required
    '         Note: Changes my this user will not cause reloading of the station table
    gPopStations
    If optSort.Value = True Then
        'cboCtrl.AddItem "[New]|-1"
        'lbcLookup1.AddItem "[New]"
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            'If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                iFound = False
                For iTest = 0 To UBound(iCode) - 1 Step 1
                    If iCode(iTest) = tgStationInfo(iLoop).iCode Then
                        iFound = True
                        Exit For
                    End If
                Next iTest
                If Not iFound Then
                    If tgStationInfo(iLoop).lMultiCastGroupID <= 0 Then
                        cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    Else
                        cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket) & " Multicast"
                    End If
                    cboCtrl.ItemData(cboCtrl.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            'End If
        Next iLoop
    Else
        'cboCtrl.AddItem "[New]|-1"
        'lbcLookup1.AddItem "[New]"
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            'If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                iFound = False
                For iTest = 0 To UBound(iCode) - 1 Step 1
                    If iCode(iTest) = tgStationInfo(iLoop).iCode Then
                        iFound = True
                        Exit For
                    End If
                Next iTest
                If Not iFound Then
                    If tgStationInfo(iLoop).lMultiCastGroupID <= 0 Then
                        cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sMarket) & ", " & Trim$(tgStationInfo(iLoop).sCallLetters)
                    Else
                        cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sMarket) & ", " & Trim$(tgStationInfo(iLoop).sCallLetters) & " Multicast"
                    End If
                    cboCtrl.ItemData(cboCtrl.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            'End If
        Next iLoop
    End If
End Sub


Private Function mAdjCPTT(iNewRec As Integer, sOnAir As String, sOffAir As String, sDropDate As String) As Integer
    
    'This module is for adding/deleting CPTT records
    Dim sLLD As String 'Last Log date
    Dim lSvCPTTStart As Long
    Dim lSvCPTTEnd As Long
    Dim lCPTTStart As Long
    Dim lCPTTEnd As Long
    Dim lSDate As Long
    Dim lEDate As Long
    Dim lDate As Long
    Dim iCycle As Integer
    Dim sTime As String
    Dim iWkDay As Integer
    Dim sMsg As String
    Dim rst_TestWk As ADODB.Recordset
    '7701
    'Dim ilAgreeType As Integer
    Dim ilExport As Integer
    Dim slStr As String
    Dim temp_rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim slServiceAgreement As String
    
    On Error GoTo ErrHand
    
    'If rbcExportType(0).Value = True Then
    '    ilAgreeType = 0
    'ElseIf rbcExportType(1).Value = True Then
    '    ilAgreeType = 1
    'ElseIf rbcExportType(3).Value = True Then
    '    ilAgreeType = 1
    'Else
    '    ilAgreeType = 2
    'End If
'7701 ilAgreeType no longer used
'    If rbcExportType(0).Value = True Then
'        ilAgreeType = 0 'Manual
'    ElseIf rbcExportType(1).Value = True Then
'        '7701
'        If ckcExportTo(0).Value = vbChecked Then
'            ilAgreeType = 1 'Network Web
'        Else
'            With cboLogDelivery
'                If .Index > -1 Then
'                    Select Case .ItemData(.Index)
'                        Case Vendors.Cumulus
'                            ilAgreeType = 1
'                        Case Vendors.NetworkConnect
'                            ilAgreeType = 3
'                        Case Vendors.cBs
'                            ilAgreeType = 4
'                    End Select
'                End If
'            End With
'        End If
''        If ckcExportTo(0).Value = vbChecked Then
''            ilAgreeType = 1 'Network Web
''        ElseIf ckcExportTo(1).Value = vbChecked Then
''            ilAgreeType = 1 'Cumulus
''        ElseIf ckcExportTo(2).Value = vbChecked Then
''            ilAgreeType = 2 'Univision
''        ElseIf ckcExportTo(3).Value = vbChecked Then
''            ilAgreeType = 3 'Marketron
''        '6592
''        ElseIf ckcExportTo(4).Value = vbChecked Then
''            ilAgreeType = 4 'CBS
''        End If
'    Else
'        ilAgreeType = -1
'    End If
    sTime = Format("12:00AM", "hh:mm:ss")

    sMsg = ""
    
    If DateValue(gAdjYear(sOnAir)) = DateValue("1/1/1970") Then
        mAdjCPTT = True
        Exit Function
    End If
    
    '6/7/19
    ilExportType = 0
    slServiceAgreement = "N"
    slStr = "Select attExportType, attServiceAgreement from att where attCode = " & lmAttCode
    Set temp_rst = gSQLSelectCall(slStr)
    If temp_rst.EOF = False Then
        ilExportType = temp_rst!attExportType
        slServiceAgreement = temp_rst!attServiceAgreement
    End If
    
    If iNewRec Then
        SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
        SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
        SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            mAdjCPTT = True
            Exit Function
        End If
        If IsNull(rst!vpfLLD) Then
            mAdjCPTT = True
            Exit Function
        End If
        If Not gIsDate(rst!vpfLLD) Then
            'sLLD = "1/1/1970"
            'iWkDay = vbMonday  'Monday
            mAdjCPTT = True
            Exit Function
        Else
            sLLD = Format$(rst!vpfLLD, sgShowDateForm)
            If Trim$(sLLD) = "" Then
                mAdjCPTT = True
                Exit Function
            End If
            'iWkDay = Weekday(Format$(DateValue(gAdjYear(sLLD)) + 1, "m/d/yyyy"))
        End If
        iCycle = 7
        If DateValue(gAdjYear(sOnAir)) <= DateValue(gAdjYear(sLLD)) Then
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sLLD)) Then
                'Add CPTT between OnAir And DropDate
                lSDate = DateValue(gAdjYear(sOnAir))
                lEDate = DateValue(gAdjYear(sDropDate))
            Else
                'Add CPTT between OnAir and LLD
                lSDate = DateValue(gAdjYear(sOnAir))
                lEDate = DateValue(gAdjYear(sLLD))
            End If
            lSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(lSDate, "m/d/yy"))))
            If lSDate <= lEDate Then
                sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                For lDate = lSDate To lEDate Step iCycle
                    'D.S. 10/25/04
                    'Before we add the new cptt recs we have to clean out any old for that week
                    SQLQuery = "DELETE FROM cptt WHERE"
                    SQLQuery = SQLQuery & " cpttVefCode = " & imVefCode
                    SQLQuery = SQLQuery & " And cpttShfCode = " & imShttCode
                    SQLQuery = SQLQuery & " And cpttStartDate = " & "'" & Format$(lDate, sgSQLDateForm) & "'"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                        mAdjCPTT = False
                        Exit Function
                    End If
                    
                    SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                    SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                    SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode)"
                    SQLQuery = SQLQuery & " VALUES "
                    SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imShttCode & ", " & imVefCode & ", "
                    SQLQuery = SQLQuery & "'" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format(lDate, sgSQLDateForm) & "', "
                    If slServiceAgreement = "Y" Then
                        SQLQuery = SQLQuery & "" & 1 & ", " & igUstCode & ")"
                    Else
                        SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ")"
                    End If
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                        mAdjCPTT = False
                        Exit Function
                    End If
                Next lDate
                gFileChgdUpdate "cptt.mkd", True
            End If
        End If
    Else
        If (DateValue(gAdjYear(sOnAir)) = DateValue(gAdjYear(smSvOnAirDate))) And (DateValue(gAdjYear(sOffAir)) = DateValue(gAdjYear(smSvOffAirDate))) And (DateValue(gAdjYear(sDropDate)) = DateValue(gAdjYear(smSvDropDate))) Then  'Append
            mAdjCPTT = True
            Exit Function
        End If
        'Get the last log date
        SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
        SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
        SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            mAdjCPTT = True
            Exit Function
        End If
        If IsNull(rst!vpfLLD) Then
            mAdjCPTT = True
            Exit Function
        End If
        If Not gIsDate(rst!vpfLLD) Then
            'sLLD = "1/1/1970"
            'iWkDay = vbMonday
            mAdjCPTT = True
            Exit Function
        Else
            'set sLLD to last log date
            sLLD = Format$(rst!vpfLLD, sgShowDateForm)
            '1= Sun, 2= Mon, 3= Tues, 4= Wed, 5= Th, 6= Fri, 7= Sat
            'iWkDay = Weekday(Format$(DateValue(gAdjYear(sLLD)) + 1, "m/d/yyyy"))
        End If
        iCycle = 7  'rst!vpfLNoDaysCycle
        lSvCPTTStart = 0
        lSvCPTTEnd = 0
        lCPTTStart = 0
        lCPTTEnd = 0
        If DateValue(gAdjYear(smSvOnAirDate)) <= DateValue(gAdjYear(sLLD)) Then
            If DateValue(gAdjYear(smSvDropDate)) <= DateValue(gAdjYear(sLLD)) Then
                lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
                If DateValue(gAdjYear(smSvDropDate)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                    lSvCPTTEnd = DateValue(gAdjYear(smSvDropDate)) '- iCycle
                Else
                    lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
                End If
            Else
                lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
                If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                    lSvCPTTEnd = DateValue(gAdjYear(sLLD))
                Else
                    lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
                End If
            End If
        End If
        If DateValue(gAdjYear(sOnAir)) <= DateValue(gAdjYear(sLLD)) Then
            If DateValue(gAdjYear(sDropDate)) <= DateValue(gAdjYear(sLLD)) Then
                lCPTTStart = DateValue(gAdjYear(sOnAir))
                If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                    lCPTTEnd = DateValue(gAdjYear(sDropDate)) '- iCycle
                Else
                    lCPTTEnd = DateValue(gAdjYear(sOffAir))
                End If
            Else
                lCPTTStart = DateValue(gAdjYear(sOnAir))
                If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(sOffAir)) Then
                    lCPTTEnd = DateValue(gAdjYear(sLLD))
                Else
                    lCPTTEnd = DateValue(gAdjYear(sOffAir))
                End If
            End If
        End If
        If ((lCPTTStart < lSvCPTTStart) And (lCPTTStart > 0)) Or ((lCPTTStart > 0) And (lSvCPTTStart = 0)) Then
            'Create
            lSDate = lCPTTStart
            'Jim(7/6/99)- back date to week start, user will have to manually turn extra days off
'            If iCycle Mod 7 = 0 Then
'                Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
'                    lSDate = lSDate - 1 '+ 1
'                Loop
'            End If
            lSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(lSDate, "m/d/yy"))))
            If lSvCPTTStart > 0 Then
                lEDate = lSvCPTTStart - iCycle
            Else
                lEDate = lCPTTEnd
            End If
            If lSDate <= lEDate Then
                sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                For lDate = lSDate To lEDate Step iCycle
                    SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                    SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                    SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode)"
                    SQLQuery = SQLQuery & " VALUES "
                    SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imShttCode & ", " & imVefCode & ", "
                    SQLQuery = SQLQuery & "'" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format(lDate, sgSQLDateForm) & "', "
                    '6/7/19
                    If slServiceAgreement = "Y" Then
                        SQLQuery = SQLQuery & "" & 1 & ", " & igUstCode & ")"
                    Else
                        SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ")"
                    End If
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                        mAdjCPTT = False
                        Exit Function
                    End If
                Next lDate
                gFileChgdUpdate "cptt.mkd", True
            End If
        End If
        'Check to see if start date Advanced.  If so delete weeks
        If ((lSvCPTTStart < lCPTTStart) And (lSvCPTTStart > 0)) Or ((lSvCPTTStart > 0) And (lCPTTStart = 0)) Then
            'Remove
            lSDate = lSvCPTTStart
            lSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(lSDate, "m/d/yy"))))
            If lCPTTStart > 0 Then
                lEDate = lCPTTStart - iCycle
            Else
                lEDate = lSvCPTTEnd
            End If
            If lSDate <= lEDate Then
                'D.S. 10/25/04
                'igChangedNewErased values  1 = changed, 2 = new, 3 = erased
                'If they are changing an agreement and it's already been exported then don't delete the CPTTs
                If sMsg <> "" Then
                    sMsg = sMsg & Chr$(13) & Chr$(10) & "Deleted weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                Else
                    sMsg = "Deleted weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                End If
                Do
                    '9/25/06: Removed retaining of CPTT as user no longer allowed to set date prior to last posted date
                    'If Not (igChangedNewErased = 1 And gCheckIfSpotsHaveBeenExported(imVefCode, Format$(lSDate, sgSQLDateForm), ilAgreeType)) Then
                        SQLQuery = "DELETE FROM cptt WHERE (cpttAtfCode = " & lmAttCode '& " And cpttShfCode =" & imShttCode & " And cpttVefCode =" & imVefCode
                        SQLQuery = SQLQuery & " AND ((cpttStatus = 2) Or ((cpttStatus = 0) AND (cpttPostingStatus = 0)))"
                        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(lSDate, sgSQLDateForm) & "' And cpttStartDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            mMousePointer vbDefault
                            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                            mAdjCPTT = False
                            Exit Function
                        End If
                    
                        ''Doug (9/25/06): Remove Web Spots if using web.  For Martketron, i'm disallowing change if aet exist within date change
                        ''Delete the spots from the web first
                        '6/7/19:Moved to top of the routine
                        'slStr = "Select attExportType from att where attCode = " & lmAttCode
                        'Set temp_rst = gSQLSelectCall(slStr)
                        'If temp_rst.EOF = False Then
                            ''7701 test no longer needed
                           '' If (temp_rst!attExportType = 1) And ((temp_rst!attExportToWeb = "Y") Or (gIfNullInteger(temp_rst!vatWvtIdCodeLog) = Vendors.Cumulus)) And gHasWebAccess Then
                           'If temp_rst!attExportType = 1 And gHasWebAccess Then
                            If ilExportType = 1 And gHasWebAccess Then
                                ilRet = gWebDeleteSpots(lmAttCode, Format$(lSDate, sgSQLDateForm), Format$(lSDate + 6, sgSQLDateForm))
                            End If
                        'End If
                    
                        '9/25/06: Added removal of ast
                        SQLQuery = "DELETE FROM ast WHERE (astAtfCode = " & lmAttCode '& " And astShfCode =" & imShttCode & " And astVefCode =" & imVefCode
                        SQLQuery = SQLQuery & " And astAirDate >= '" & Format$(lSDate, sgSQLDateForm) & "' And astAirDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            mMousePointer vbDefault
                            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                            mAdjCPTT = False
                            Exit Function
                        End If
                    
                    'End If
                    lSDate = lSDate + 7
                Loop While lSDate <= lEDate
                gFileChgdUpdate "cptt.mkd", True
            End If
        End If
        
        'Check to see if end date advanced.  If so add weeks
        If (lSvCPTTEnd < lCPTTEnd) And (lSvCPTTEnd > 0) Then
            'Create
            lSDate = lSvCPTTEnd  'iCycle
            'Jim(7/6/99)- back date to week start, user will have to manually turn extra days off
            If iCycle Mod 7 = 0 Then
                lSDate = lSDate + 1
                'Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> vbMonday
                    lSDate = lSDate - 1 '+ 1
                Loop
            Else
                lSDate = lSDate + 1
            End If
            lEDate = lCPTTEnd
            If lSDate <= lEDate Then
                'D.S. 09/19/02 Test if the first week already exists before inserting.
                'If so don't insert it increase the start date by one week. Covers the
                'case where the entered drop date is not a Sunday causing it to insert
                'starting one week to early making a dupicate week.
                SQLQuery = "Select cpttAtfCode from cptt "
                SQLQuery = SQLQuery & "Where (cpttAtfCode =  " & lmAttCode & " "
                SQLQuery = SQLQuery & "And cpttShfCode = " & imShttCode & " "
                SQLQuery = SQLQuery & "And cpttVefCode = " & imVefCode & " "
                SQLQuery = SQLQuery & "And cpttStartdate = '" & Format(lSDate, sgSQLDateForm) & "' )"
                Set rst_TestWk = gSQLSelectCall(SQLQuery)
                If Not rst_TestWk.EOF Then
                    lSDate = lSDate + 7 'Add a week to the start date
                End If
                Set rst_TestWk = Nothing
                'End 09/19/02
                
                If sMsg <> "" Then
                    sMsg = sMsg & Chr$(13) & Chr$(10) & "Added weeks: " & Format$(lSDate, "m/d/yyyy") & "-" & Format$(lEDate, "m/d/yyyy")
                Else
                    sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                End If
                
                'Generate a CPTT record one/week
                For lDate = lSDate To lEDate Step 7 'iCycle
                    SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                    SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                    SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode)"
                    SQLQuery = SQLQuery & " VALUES "
                    SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imShttCode & ", " & imVefCode & ", "
                    SQLQuery = SQLQuery & "'" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format(lDate, sgSQLDateForm) & "', "
                    '6/7/19
                    If slServiceAgreement = "Y" Then
                        SQLQuery = SQLQuery & "" & 1 & ", " & igUstCode & ")"
                    Else
                        SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ")"
                    End If
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                        mAdjCPTT = False
                        Exit Function
                    End If
                Next lDate
                gFileChgdUpdate "cptt.mkd", True
            End If
        End If
        'Check to see if end date reduced.  If so, remove weeks
        If (lCPTTEnd < lSvCPTTEnd) And (lCPTTEnd > 0) Then
            'Remove
            'lSDate = lCPTTEnd + iCycle
            lSDate = lCPTTEnd  'iCycle
            'Advance to next week as dates are for last week to air
            If iCycle Mod 7 = 0 Then
                lSDate = lSDate + 1
                'Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> vbMonday
                    lSDate = lSDate + 1
                Loop
            Else
                lSDate = lSDate + 1
            End If
            lEDate = lSvCPTTEnd
            If lSDate <= lEDate Then
                'D.S. 12/06/04  We need to clean-up any week where logs have been generated and that week has
                'never been exported.
                If sMsg <> "" Then
                    sMsg = sMsg & Chr$(13) & Chr$(10) & "Deleted weeks: " & Format$(lSDate, "m/d/yyyy") & "-" & Format$(lEDate, "m/d/yyyy")
                Else
                    sMsg = "Deleted weeks: " & Format$(lSDate, "m/d/yyyy") & "-" & Format$(lEDate, "m/d/yyyy")
                End If
                Do
                    '9/25/06: Removed retaining of CPTT as user no longer allowed to set date prior to last posted date
                    'ilExport = gCheckIfSpotsHaveBeenExported(imVefCode, Format$(lSDate, sgSQLDateForm), ilAgreeType)
                    ''D.S. 10/25/04
                    ''igChangedNewErased values  1 = changed, 2 = new, 3 = erased
                    ''If they are changing an agreement and it's already been exported then don't delete the CPTTs
                    'If (igChangedNewErased = 1) And (ilExport = False) Then
                        SQLQuery = "DELETE FROM cptt WHERE (cpttAtfCode = " & lmAttCode '& " And cpttShfCode =" & imShttCode & " And cpttVefCode =" & imVefCode
                        SQLQuery = SQLQuery & " And cpttStartDate >='" & Format$(lSDate, sgSQLDateForm) & "' And cpttStartDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            mMousePointer vbDefault
                            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                            mAdjCPTT = False
                            Exit Function
                        End If
                    'End If
                    
                    ''Doug (9/25/06): Remove Web Spots if using web.  For Martketron, i'm disallowing change if aet exist within date change
                    ''Delete the spots from the web first
                    ''dan adjusted what was returned 11/30/15
                    '6/7/19: Moved to top of this routine
                    'slStr = "Select attExportType from att where attCode = " & lmAttCode
                    'Set temp_rst = gSQLSelectCall(slStr)
                    'If temp_rst.EOF = False Then
                        '7701
                       ' If (temp_rst!attExportType = 1) And ((temp_rst!attExportToWeb = "Y") Or (gIfNullInteger(temp_rst!vatWvtIdCodeLog) = Vendors.Cumulus)) And gHasWebAccess Then
                        'If temp_rst!attExportType = 1 And gHasWebAccess Then
                        If ilExportType = 1 And gHasWebAccess Then
                            ilRet = gWebDeleteSpots(lmAttCode, Format$(lSDate, sgSQLDateForm), Format$(lSDate + 6, sgSQLDateForm))
                        End If
                    'End If
                    
                    ''D.S. 10/25/04
                    'If ilExport = False Then
                        SQLQuery = "DELETE FROM ast WHERE (astAtfCode = " & lmAttCode '& " And astShfCode =" & imShttCode & " And astVefCode =" & imVefCode
                        SQLQuery = SQLQuery & " And astAirDate >= '" & Format$(lSDate, sgSQLDateForm) & "' And astAirDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            mMousePointer vbDefault
                            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjCPTT"
                            mAdjCPTT = False
                            Exit Function
                        End If
                    'End If
                    lSDate = lSDate + 7
                Loop While lSDate <= lEDate
                gFileChgdUpdate "cptt.mkd", True
            End If
        End If
    End If
    If sMsg <> "" Then
        gMsgBox sMsg, vbOKOnly
    End If
    mAdjCPTT = True
    Exit Function
ErrHand:
    mMousePointer vbDefault
    gHandleError "", "Agreement-mAdjCptt"
    mAdjCPTT = False
End Function

Private Function mCPDelete(sOnAir As String, sOffAir As String, sDropDate As String) As Integer
'Test if CP will be deleted, and if so, ask user if Ok to proceed
    Dim sLLD As String 'Last Log date
    Dim lSvCPTTStart As Long
    Dim lSvCPTTEnd As Long
    Dim lCPTTStart As Long
    Dim lCPTTEnd As Long
    Dim lSDate As Long
    Dim lEDate As Long
    Dim lDate As Long
    Dim iCycle As Integer
    Dim sTime As String
    Dim iWkDay As Integer
    Dim iRet As Integer
    Dim sMsg As String
    
    On Error GoTo ErrHand
    sTime = Format("12:00AM", "hh:mm:ss")
    sMsg = ""
    
    If DateValue(gAdjYear(sOnAir)) = DateValue("1/1/1970") Then
        mCPDelete = True
        Exit Function
    End If
    
    If (DateValue(gAdjYear(sOnAir)) = DateValue(gAdjYear(smSvOnAirDate))) And (DateValue(gAdjYear(sOffAir)) = DateValue(gAdjYear(smSvOffAirDate))) And (DateValue(gAdjYear(sDropDate)) = DateValue(gAdjYear(smSvDropDate))) Then  'Append
        mCPDelete = True
        Exit Function
    End If
    SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mCPDelete = True
        Exit Function
    End If
    If IsNull(rst!vpfLLD) Then
        mCPDelete = True
        Exit Function
    End If
    If Not gIsDate(rst!vpfLLD) Then
        mCPDelete = True
        Exit Function
    Else
        sLLD = Format$(rst!vpfLLD, sgShowDateForm)
        iWkDay = Weekday(Format$(DateValue(gAdjYear(sLLD)) + 1, "m/d/yyyy"))
    End If
    iCycle = rst!vpfLNoDaysCycle
    lSvCPTTStart = 0
    lSvCPTTEnd = 0
    lCPTTStart = 0
    lCPTTEnd = 0
    If DateValue(gAdjYear(smSvOnAirDate)) <= DateValue(gAdjYear(sLLD)) Then
        If DateValue(gAdjYear(smSvDropDate)) <= DateValue(gAdjYear(sLLD)) Then
            lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
            If DateValue(gAdjYear(smSvDropDate)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                lSvCPTTEnd = DateValue(gAdjYear(smSvDropDate)) '- iCycle
            Else
                lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
            End If
        Else
            lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
            If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                lSvCPTTEnd = DateValue(gAdjYear(sLLD))
            Else
                lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
            End If
        End If
    End If
    If DateValue(gAdjYear(sOnAir)) <= DateValue(gAdjYear(sLLD)) Then
        If DateValue(gAdjYear(sDropDate)) <= DateValue(gAdjYear(sLLD)) Then
            lCPTTStart = DateValue(gAdjYear(sOnAir))
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                lCPTTEnd = DateValue(gAdjYear(sDropDate)) '- iCycle
            Else
                lCPTTEnd = DateValue(gAdjYear(sOffAir))
            End If
        Else
            lCPTTStart = DateValue(gAdjYear(sOnAir))
            If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(sOffAir)) Then
                lCPTTEnd = DateValue(gAdjYear(sLLD))
            Else
                lCPTTEnd = DateValue(gAdjYear(sOffAir))
            End If
        End If
    End If
    'If ((lCPTTStart < lSvCPTTStart) And (lCPTTStart > 0)) Or ((lCPTTStart > 0) And (lSvCPTTStart = 0)) Then
    '    'Create
    '    lSDate = lCPTTStart
    '    'Jim(7/6/99)- back date to week start, user will have to manually turn extra days off
    '    If iCycle Mod 7 = 0 Then
    '        Do While WeekDay(Format$(lSDate, sgShowDateForm)) <> iWkDay
    '            lSDate = lSDate - 1 '+ 1
    '        Loop
    '    End If
    '    If lSvCPTTStart > 0 Then
    '        lEDate = lSvCPTTStart - iCycle
    '    Else
    '        lEDate = lCPTTEnd
    '    End If
    '    sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
    '    For lDate = lSDate To lEDate Step iCycle
    '        SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
    '        SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, cpttCycle, cpttAirTime, "
    '        SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode, cpttPrintStatus)"
    '        SQLQuery = SQLQuery & " VALUES "
    '        SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imShttCode & ", " & imVefCode & ", "
    '        SQLQuery = SQLQuery & "" & smCurDate & ", " & Format(lDate, "mm/dd/yyyy") & ", " & iCycle & ", " & sTime & ", "
    '        SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ", " & 0 & ")"
    '        cnn.Execute SQLQuery, rdExecDirect
    '    Next lDate
    'End If
    If ((lSvCPTTStart < lCPTTStart) And (lSvCPTTStart > 0)) Or ((lSvCPTTStart > 0) And (lCPTTStart = 0)) Then
        'Remove
        lSDate = lSvCPTTStart
        'Jim(7/6/99)- back date to week start, user will have to manually turn extra days off
        If iCycle Mod 7 = 0 Then
            Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                lSDate = lSDate - 1 '+ 1
            Loop
        End If
        If lCPTTStart > 0 Then
            lEDate = lCPTTStart - iCycle
        Else
            lEDate = lSvCPTTEnd
        End If
        If lSDate <= lEDate Then
            If sMsg <> "" Then
                sMsg = sMsg & Chr$(13) & Chr$(10) & "and weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
            Else
                sMsg = "The following weeks will be Deleted: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
            End If
        End If
    End If
    
    'If (lSvCPTTEnd < lCPTTEnd) And (lSvCPTTEnd > 0) Then
    '    'Create
    '    lSDate = lSvCPTTEnd  'iCycle
    '    'Jim(7/6/99)- back date to week start, user will have to manually turn extra days off
    '    If iCycle Mod 7 = 0 Then
    '        lSDate = lSDate + 1
    '        Do While WeekDay(Format$(lSDate, sgShowDateForm)) <> iWkDay
    '            lSDate = lSDate - 1 '+ 1
    '        Loop
    '    Else
    '        lSDate = lSDate + 1
    '    End If
    '    lEDate = lCPTTEnd
    '    If sMsg <> "" Then
    '        sMsg = sMsg & Chr$(13) & Chr$(10) & "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
    '    Else
    '        sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
    '    End If
    '    'Generate a CPTT record one/week
    '    For lDate = lSDate To lEDate Step 7 'iCycle
    '        SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
    '        SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, cpttCycle, cpttAirTime, "
    '        SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode, cpttPrintStatus)"
    '        SQLQuery = SQLQuery & " VALUES "
    '        SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imShttCode & ", " & imVefCode & ", "
    '        SQLQuery = SQLQuery & "" & smCurDate & ", " & Format(lDate, "mm/dd/yyyy") & ", " & iCycle & ", " & sTime & ", "
    '        SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ", " & 0 & ")"
    '        cnn.Execute SQLQuery, rdExecDirect
    '    Next lDate
    'End If
    If (lCPTTEnd < lSvCPTTEnd) And (lCPTTEnd > 0) Then
        'Remove
        lSDate = lCPTTEnd  'iCycle
        'Advance to next week as dates are for last week to air
        If iCycle Mod 7 = 0 Then
            lSDate = lSDate + 1
            Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                lSDate = lSDate + 1
            Loop
        Else
            lSDate = lSDate + 1
        End If
        lEDate = lSvCPTTEnd
        If lSDate <= lEDate Then
            If sMsg <> "" Then
                sMsg = sMsg & Chr$(13) & Chr$(10) & "and weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
            Else
                sMsg = "The following weeks will be Deleted: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
            End If
        End If
    End If
        
    If sMsg <> "" Then
        iRet = gMsgBox(sMsg, vbOKCancel, "Delete Weeks")
        If iRet = vbOK Then
            mCPDelete = True
        Else
            mCPDelete = False
        End If
    Else
        mCPDelete = True
    End If
    Exit Function
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmAgmnt -UpdateWebSite"
    mCPDelete = False
    Exit Function
End Function
'6/28/18: Add ShowMsg
'Private Function mDateOverlap(lAttCode As Long, lOnAir As Long, lOffAir As Long, lDropDate As Long) As Integer
Private Function mDateOverlap(lAttCode As Long, lOnAir As Long, lOffAir As Long, lDropDate As Long, blShowMsg As Boolean) As Integer
'Test if Agreement date overlap- if so disallow agreement being saved
    Dim sOnAir As String
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sEndDate As String
    Dim lEndDate As Long
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim slLastWebPostedDate As String
    Dim slTempDate As String
    '7701
    Dim slattExportToMarketron As String
    Dim slattWebInterface As String
    
    If lDropDate < lOffAir Then
        lEndDate = lDropDate
    Else
        lEndDate = lOffAir
    End If
    ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
    On Error GoTo ErrHand
    '7701
'    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attExportType, attExportToWeb, attWebInterface, attExportToUnivision, attExportToMarketron FROM att"
    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attExportType, attExportToWeb, attWebInterface, attexporttoUnivision FROM att"
    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode & " AND attVefCode = " & imVefCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        '7701
        slattExportToMarketron = "N"
        slattWebInterface = ""
        If gIsVendorWithAgreement(rst!attCode, Vendors.stratus) Then
            slattWebInterface = "C"
        End If
        If gIsVendorWithAgreement(rst!attCode, Vendors.NetworkConnect) Then
            slattExportToMarketron = "Y"
        End If
'
'        If gIfNullInteger(rst!vatWvtIdCodeLog) = vendors.Stratus Then
'            slattWebInterface = "C"
'        ElseIf gIfNullInteger(rst!vatWvtIdCodeLog) = Vendors.NetworkConnect Then
'            slattExportToMarketron = "Y"
'        End If
        'Test dates
        If (lAttCode <> rst!attCode) Then
            'sOnAir = Format$(rst!attOnAir, sgShowDateForm)
            sOnAir = Format$(rst!attOnAir, "mm/dd/yyyy")
            'sOffAir = Format$(rst!attOffAir, sgShowDateForm)
            sOffAir = Format$(rst!attOffAir, "mm/dd/yyyy")
            'sDropDate = Format$(rst!attDropDate, sgShowDateForm)
            sDropDate = Format$(rst!attDropDate, "mm/dd/yyyy")
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                sEndDate = sDropDate
            Else
                sEndDate = sOffAir
            End If
            If (lEndDate >= DateValue(gAdjYear(sOnAir))) And (DateValue(gAdjYear(sEndDate)) >= lOnAir) Then
                'Dates overlap, get last posted date for agreement
                slLastWebPostedDate = "1/1/1990"
                'If rbcExportType(1).Value = True And gUsingWeb Then
                'If rst!attExportType = 1 And gUsingWeb Then
                '7701
                If ((rst!attExportToWeb = "Y") Or (slattWebInterface = "C")) And gUsingWeb Then
               ' If ((rst!attExportToWeb = "Y") Or (rst!attWebInterface = "C")) And gUsingWeb Then
                    'Obtain the last Posted date
                    SQLQuery = "SELECT max(weblPostDay) FROM webl WHERE"
                    SQLQuery = SQLQuery & " weblType = 1 And weblAttCode = " & rst!attCode
                    Set rst_Webl = gSQLSelectCall(SQLQuery)
                    If Not rst_Webl.EOF Then
                        If IsDate(rst_Webl(0).Value) Then
                            slLastWebPostedDate = Format$(rst_Webl(0).Value, sgShowDateForm)
                        End If
                    End If
                    'Doug (9/25/06)- Add test to see if submitted date is newer then slLastWebPostedDate
                    '      If so, update slLastWebPostedDate with that date
                    '      Don't show error message if unable to access web
                    'ilRet = gRemoteExecSql("Select Max(PledgeStartDate) from spots WITH (INDEX(IX_Spots_Headers)) Where where attCode = " & "'" & rst!attCode & "'", "MaxPostDate.txt", "WebImports", True, True, 30)
                    ilRet = gRemoteExecSql("Select Max(PledgeStartDate) as MaxPostDate from spots where PostedFlag = 1 AND attCode = " & "'" & rst!attCode & "'", "MaxPostDate.txt", "WebImports", True, True, 30)
                    slTempDate = gRemoteMaxPostDayResults("MaxPostDate.txt", "WebImports")
                    If slTempDate <> "" Then
                        If DateValue(slTempDate) > DateValue(slLastWebPostedDate) Then
                            slLastWebPostedDate = slTempDate
                        End If
                    End If
                End If
                'ElseIf rst!attExportType = 2 Then   'Univision: use latest aet date as last posted date event thuo it is not
                If rst!attExportToUnivision = "Y" Then
                    SQLQuery = " Select Max(aetFeedDate) FROM Aet"
                    SQLQuery = SQLQuery & " Where aetatfCode = " & rst!attCode
                    Set rst_Aet = gSQLSelectCall(SQLQuery)
                    If Not rst_Aet.EOF Then
                        'D.S. 10/11/06
                        'If rst_Aet(0).Value <> Null Then
                        If IsDate(rst_Aet(0).Value) Then
                            slLastWebPostedDate = Format$(rst_Aet(0).Value, sgShowDateForm)
                        End If
                    End If
                End If
                'dan M added Marketron posting 11/18/10
                '7701
                If slattExportToMarketron = "Y" Then
                    SQLQuery = "SELECT max(weblpostday) FROM webl WHERE webltype = 3 and weblattcode = " & rst!attCode
                    Set rst_Webl = gSQLSelectCall(SQLQuery)
                    If Not rst_Webl.EOF Then
                        If IsDate(rst_Webl(0).Value) Then
                            If DateDiff("d", slLastWebPostedDate, rst_Webl(0).Value) > 0 Then
                                slLastWebPostedDate = Format$(rst_Webl(0).Value, sgShowDateForm)
                            End If
                        End If
                    End If
                End If
                'Get lastest posted date as user will not be allowed to drop prior to that date
                SQLQuery = "SELECT max(astFeedDate) FROM ast WHERE"
                SQLQuery = SQLQuery & " astAtfCode = " & rst!attCode
                SQLQuery = SQLQuery & " AND astCPStatus = 1"
                Set rst_Ast = gSQLSelectCall(SQLQuery)
                If Not rst_Ast.EOF Then
                    If IsDate(rst_Ast(0).Value) Then
                        If (slLastWebPostedDate = "") Or (DateValue(Format$(rst_Ast(0).Value, sgShowDateForm)) > DateValue(slLastWebPostedDate)) Then
                            slLastWebPostedDate = Format$(rst_Ast(0).Value, sgShowDateForm)
                        End If
                    End If
                End If
                '6/28/18
                If blShowMsg Then
                    If imSource = 0 Then
                        If lOnAir <= DateValue(slLastWebPostedDate) Then
                            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
                            gMsgBox "On air dates overlap with previously defined Agreement that was Posted through " & slLastWebPostedDate & ", Can't Save", vbOKOnly, "Save"
                            mDateOverlap = True
                            Exit Function
                        End If
                    Else
                        If (lOnAir < DateValue(gAdjYear(smSvOnAirDate))) And (lOnAir <= DateValue(slLastWebPostedDate)) Then
                            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
                            gMsgBox "On air dates overlap with previously defined Agreement that was Posted through " & slLastWebPostedDate & ", Can't Save", vbOKOnly, "Save"
                            mDateOverlap = True
                            Exit Function
                        End If
                        If (lOffAir > DateValue(gAdjYear(smSvOffAirDate))) And (DateValue("1/1/1990") <> DateValue(slLastWebPostedDate)) Then
                            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
                            gMsgBox "Off air dates overlap with previously defined Agreement that was Posted through " & slLastWebPostedDate & ", Can't Save", vbOKOnly, "Save"
                            mDateOverlap = True
                            Exit Function
                        End If
                        If (lDropDate > DateValue(gAdjYear(smSvDropDate))) And (DateValue("1/1/1990") <> DateValue(slLastWebPostedDate)) Then
                            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
                            gMsgBox "Drop dates overlap with previously defined Agreement that was Posted through " & slLastWebPostedDate & ", Can't Save", vbOKOnly, "Save"
                            mDateOverlap = True
                            Exit Function
                        End If
                    End If
                End If
                ilUpper = UBound(tmOverlapInfo)
                tmOverlapInfo(ilUpper).lAttCode = rst!attCode
                tmOverlapInfo(ilUpper).lOnAirDate = DateValue(gAdjYear(sOnAir))
                tmOverlapInfo(ilUpper).lOffAirDate = DateValue(gAdjYear(sOffAir))
                tmOverlapInfo(ilUpper).lDropDate = DateValue(gAdjYear(sDropDate))
                ReDim Preserve tmOverlapInfo(0 To ilUpper + 1) As AGMNTOVERLAPINFO
                'gMsgBox "On/Off dates overlap with previously defined Agreement, Can't Save", vbOKOnly, "Save"
                'mDateOverlap = True
                'Exit Function
            End If
        End If
        rst.MoveNext
    Wend
    If UBound(tmOverlapInfo) > LBound(tmOverlapInfo) Then
        '6/28/18
        If (UBound(tmOverlapInfo) >= LBound(tmOverlapInfo)) + 2 And (lEndDate < DateValue("12/31/2069")) Then
            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
            gMsgBox "Dates overlap multi-Agreements and is not TFN, Can't Save", vbOKOnly, "Save"
            mDateOverlap = True
            Exit Function
        End If
        If imSource = 0 Then    'New
            If blShowMsg Then
                ilRet = gMsgBox("Agreement Dates Overlap, Continue with Save by Terminate Overlapped Agreement?", vbYesNo)
                If ilRet = vbNo Then
                    mDateOverlap = True
                    Exit Function
                End If
            End If
        Else    'Changed
            ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
            gMsgBox "On/Off dates overlap with previously defined Agreement, Can't Save", vbOKOnly, "Save"
            mDateOverlap = True
            Exit Function
        End If
    End If
    mDateOverlap = False
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mDateOverlap"
    mDateOverlap = True
    Exit Function
End Function

Private Function mTestDaypart(lbInSave As Boolean)
    Dim iLoop As Integer
    Dim iPack As Integer
    Dim sTime As String
    Dim iDay As Integer
    Dim iIndex As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilError As Integer
    Dim llRowIndex As Long
    Dim llSTime As Long
    Dim llETime As Long
    Dim slPdDayFed As String
    Dim llPrgSTime As Long
    Dim llPrgETime As Long
    Dim ilRet As Integer
    
    grdPledge.Redraw = False
    'Test if fields defined
    ilError = 0
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        grdPledge.Row = llRow
        For llCol = MONFDINDEX To ESTIMATETIMEINDEX Step 1
            grdPledge.Col = llCol
            grdPledge.CellForeColor = vbBlack
        Next llCol
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            sTime = grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX)
            If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
                ilError = 1
                If Len(Trim$(sTime)) = 0 Then
                    grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX) = "Missing"
                End If
                grdPledge.Row = llRow
                grdPledge.Col = STARTTIMEFDINDEX
                grdPledge.CellForeColor = vbRed
            End If
            llSTime = gTimeToLong(sTime, False)
            sTime = grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX)
            If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then    'Time not valid.
                ilError = 1
                If Len(Trim$(sTime)) = 0 Then
                    grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX) = "Missing"
                End If
                grdPledge.Row = llRow
                grdPledge.Col = ENDTIMEFDINDEX
                grdPledge.CellForeColor = vbRed
            End If
            llETime = gTimeToLong(sTime, True)
            If llETime < llSTime Then
                ilError = 1
                grdPledge.Row = llRow
                grdPledge.Col = ENDTIMEFDINDEX
                grdPledge.CellForeColor = vbRed
            End If
            'Test that the feed times are not outside of the program times
            If lbInSave Then
                llPrgSTime = gTimeToLong(sgVehProgStartTime, False)
                llPrgETime = gTimeToLong(sgVehProgEndTime, True)
                If llPrgSTime < llPrgETime Then
                    If (llSTime < llPrgSTime) Or (llETime > llPrgETime) Then
                        ilError = 2
                        grdPledge.Row = llRow
                        grdPledge.Col = STARTTIMEFDINDEX
                        grdPledge.CellForeColor = vbRed
                    End If
                ElseIf llPrgSTime > llPrgETime Then
                    If (llSTime < llPrgSTime) Or (llETime > 86400) Then
                        If (llSTime < 0) Or (llETime > llPrgETime) Then
                            ilError = 2
                            grdPledge.Row = llRow
                            grdPledge.Col = STARTTIMEFDINDEX
                            grdPledge.CellForeColor = vbRed
                        End If
                    End If
                End If
            End If
            slStr = grdPledge.TextMatrix(llRow, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tmStatusTypes(iIndex).iPledged = 1 Then
                    sTime = grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX)
                    If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
                        ilError = 1
                        If Len(Trim$(sTime)) = 0 Then
                            grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = "Missing"
                        End If
                        grdPledge.Row = llRow
                        grdPledge.Col = STARTTIMEPDINDEX
                        grdPledge.CellForeColor = vbRed
                    End If
                    llSTime = gTimeToLong(sTime, False)
                    sTime = grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX)
                    If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then    'Time not valid.
                        ilError = 1
                        If Len(Trim$(sTime)) = 0 Then
                            grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = "Missing"
                        End If
                        grdPledge.Row = llRow
                        grdPledge.Col = ENDTIMEPDINDEX
                        grdPledge.CellForeColor = vbRed
                    End If
                    llETime = gTimeToLong(sTime, True)
                    If llETime < llSTime Then
                        ilError = 1
                        grdPledge.Row = llRow
                        grdPledge.Col = ENDTIMEPDINDEX
                        grdPledge.CellForeColor = vbRed
                    End If
                    If mPdPriorFd(llRow) Then
                        slPdDayFed = grdPledge.TextMatrix(llRow, DAYFEDINDEX)
                        If (slPdDayFed <> "B") And (slPdDayFed <> "A") Then
                            ilError = 2
                            If Len(Trim$(slPdDayFed)) = 0 Then
                                grdPledge.TextMatrix(llRow, DAYFEDINDEX) = "Missing"
                            End If
                            grdPledge.Row = llRow
                            grdPledge.Col = DAYFEDINDEX
                            grdPledge.CellForeColor = vbRed
                        End If
                    End If
                End If
            End If
        End If
    Next llRow
    If ilError > 0 Then
        'If (IsAgmntDirty = True) Or (ilError = 1) Then
        If (ilError = 1) Then
            grdPledge.Redraw = True
            If Not frcTab(2).Visible Then
                'gSendKeys "%P", True
                TabStrip1.Tabs(TABPLEDGE).Selected = True
            End If
            Beep
            mMousePointer vbDefault
            'If (pbcPledgeSTab.Enabled) And (pbcPledgeSTab.Visible) Then
            '    pbcPledgeSTab.SetFocus
            'End If
            mTestDaypart = False
            gMsgBox "Error in Times.  Errors are Indicated in Red."
            Exit Function
        Else
            ilRet = MsgBox("Feed and/or Pledge Time outside of Program Time, Continue with Save.", vbYesNo, "Time Error")
            If ilRet = vbNo Then
                mTestDaypart = False
                mMousePointer vbDefault
                If Not frcTab(2).Visible Then
                    'gSendKeys "%P", True
                    TabStrip1.Tabs(TABPLEDGE).Selected = True
                End If
                'If (pbcPledgeSTab.Enabled) And (pbcPledgeSTab.Visible) Then
                '    pbcPledgeSTab.SetFocus
                'End If
                Exit Function
            End If
            mTestDaypart = True
        End If
    Else
        mTestDaypart = True
        Exit Function
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mAdjustDate                     *
'*                                                     *
'*             Created:7/11/01       By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust dates greater than the  *
'*                      year 2070 to the year 19XX     *
'*                                                     *
'*                      Exp. 2099 becomes 1999         *
'*                           2082 becomes 1982         *
'*                                                     *
'*******************************************************
Public Sub mAdjustDates()

    Dim llLoop As Long
    Dim ilYear As Integer
    Dim slDate As String
            
    On Error GoTo ErrHand
    If gMsgBox("Adjust Start and End Dates Greater than Year 2070 To 19XX?", vbYesNo) = vbYes Then
        mMousePointer vbHourglass
        For llLoop = 0 To UBound(tmAdjustDates) - 1 Step 1
            'If the length <> 10 then the date can't be >= 2070. Exp. 02/03/2099
            If Len(tmAdjustDates(llLoop).sAttAgreeStart) = 10 Then
                If Year(tmAdjustDates(llLoop).sAttAgreeStart) >= 2070 Then
                    slDate = Left$(tmAdjustDates(llLoop).sAttAgreeStart, 6)
                    slDate = slDate & "19"
                    slDate = slDate & right$(tmAdjustDates(llLoop).sAttAgreeStart, 2)
                    SQLQuery = "UPDATE att SET "
                    '5589 Dan 11/07/12 Dick says remove here
                    SQLQuery = SQLQuery & "attAgreeStart = '" & Format$(slDate, sgSQLDateForm) & "' "
                    'SQLQuery = SQLQuery & "attAgreeStart = '" & Format$(slDate, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                    SQLQuery = SQLQuery & " WHERE attCode = " & tmAdjustDates(llLoop).lAttCode
                    'cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjustDates"
                        Exit Sub
                    End If
                    'cnn.CommitTrans
                End If
            End If
            
            'If the length <> 10 then the date can't be >= 2070 Exp. 02/03/2099
            If Len(tmAdjustDates(llLoop).sAttAgreeEnd) = 10 Then
                If Year(tmAdjustDates(llLoop).sAttAgreeEnd) >= 2070 Then
                    slDate = Left$(tmAdjustDates(llLoop).sAttAgreeEnd, 6)
                    slDate = slDate & "19"
                    slDate = slDate & right$(tmAdjustDates(llLoop).sAttAgreeEnd, 2)
                    SQLQuery = "UPDATE att SET "
                    '5589 Dan 11/07/12 Dick says remove here
                    SQLQuery = SQLQuery & "attAgreeEnd = '" & Format$(slDate, sgSQLDateForm) & "' "
                    'SQLQuery = SQLQuery & "attAgreeEnd = '" & Format$(slDate, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                    SQLQuery = SQLQuery & " WHERE attCode = " & tmAdjustDates(llLoop).lAttCode
                    'cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        mMousePointer vbDefault
                        gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjustDates"
                        Exit Sub
                    End If
                    'cnn.CommitTrans
                End If
            End If
        Next llLoop
        mMousePointer vbDefault
        cmdAdjustDates.Visible = False
    End If
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmAgmnt-general-mAdjustDates"
End Sub


Public Sub mGetDayParts()
    
    Dim ilLoop As Integer
    Dim ilGroupNo As Integer
    Dim slVefType As String * 1
    Dim ilRdfIdx As Integer
    Dim ilTestIdx As Integer
    Dim ilFound As Integer
    Dim ilVefCode As Integer
    Dim rst_DayParts As ADODB.Recordset
    Dim ilVehArray() As Integer
    
    On Error GoTo ErrHand
    
    'Get the user selected vehicle Type and GroupNo
    For ilLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
        If tgVehicleInfo(ilLoop).iCode = imVefCode Then
            slVefType = tgVehicleInfo(ilLoop).sVehType
            ilGroupNo = tgVehicleInfo(ilLoop).iVpfSAGroupNo
            ilVefCode = tgVehicleInfo(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
    
    ReDim ilVehArray(0 To 0) As Integer
       
    'Conventional or Sports Vehicle
    '11/4/09-  Show Log and Conventional vehicle.  Let client pick which they want agreements to be used for
    'Temporarily include only for Special user until testing is complete
    'If (slVefType = "C") Or (slVefType = "G") Then
    'If (slVefType = "C") Or (slVefType = "G") Or ((Len(sgSpecialPassword) = 4) And (slVefType = "L")) Then
    If (slVefType = "C") Or (slVefType = "G") Then
        ilVehArray(UBound(ilVehArray)) = ilVefCode  'tgVehicleInfo(ilLoop).iCode
        If imVefCombo > 0 Then
            ReDim Preserve ilVehArray(0 To 2)
            ilVehArray(1) = imVefCombo
        Else
            ReDim Preserve ilVehArray(0 To 1)
        End If
        ilLoop = ilLoop
    End If

    'Airing Vehicle - Get the Selling Vehicles that map to the Airing Vehicles
    If slVefType = "A" Then
        For ilLoop = LBound(tgSellingVehicleInfo) To UBound(tgSellingVehicleInfo) Step 1
            If tgSellingVehicleInfo(ilLoop).iVpfSAGroupNo = ilGroupNo Then
                ilVehArray(UBound(ilVehArray)) = tgSellingVehicleInfo(ilLoop).iCode
                ReDim Preserve ilVehArray(0 To (UBound(ilVehArray) + 1))
            End If
        Next ilLoop
    End If
    
    If slVefType = "L" Then
        For ilLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
            If tgVehicleInfo(ilLoop).iVefCode = ilVefCode Then
                ilVehArray(UBound(ilVehArray)) = tgVehicleInfo(ilLoop).iCode
                ReDim Preserve ilVehArray(0 To (UBound(ilVehArray) + 1))
            End If
        Next ilLoop
    End If
    
    ReDim tgRdfCodes(0 To 0) As Integer
    'Get the rdf code from the Rif table for every vehicle in the vehicle array
    For ilLoop = LBound(ilVehArray) To UBound(ilVehArray) - 1 Step 1
        SQLQuery = "SELECT rifRdfCode"
        SQLQuery = SQLQuery & " FROM RIF_Rate_Card_Items"
        SQLQuery = SQLQuery + " WHERE (rifVefCode = " & ilVehArray(ilLoop) '& " OR rifVefCode = " & imVefCombo & ""
        SQLQuery = SQLQuery + " AND rifRcfCode = " & tgLatestRateCard(0).iLatestRCFCode & ")"
                
        Set rst_DayParts = gSQLSelectCall(SQLQuery)
        While Not rst_DayParts.EOF
            ilFound = False
            'Check to see if the Rdf code is already in the array.
            'We don't save duplicate Rdf codes
            For ilTestIdx = 0 To UBound(tgRdfCodes) Step 1
                If tgRdfCodes(ilTestIdx) = rst_DayParts!rifRdfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilTestIdx
            
            'Save the unique Rdf codes
            If Not ilFound Then
                tgRdfCodes(UBound(tgRdfCodes)) = rst_DayParts!rifRdfCode
                ReDim Preserve tgRdfCodes(0 To (UBound(tgRdfCodes) + 1))
            End If
            
            rst_DayParts.MoveNext
        Wend
    Next ilLoop
    Set rst_DayParts = Nothing
    Exit Sub
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mGetDayParts"
End Sub

Public Function mTestForDPConflict() As Integer

    Dim ilDayIdx As Integer
    Dim ilHourIdx As Integer
    Dim ilDatIdx As Integer
    Dim llDatFdSTime As Long
    Dim llDatFdETime As Long
    Dim llDatPdSTime As Long
    Dim llDatPdETime As Long
    Dim ilFound As Integer
               
    'Initialiaze Fed/Sold daypart arrary
    For ilDayIdx = 0 To 6 Step 1
        ReDim tmFDDayTime(ilDayIdx).lSTime(0 To 0)
        ReDim tmFDDayTime(ilDayIdx).lETime(0 To 0)
    Next ilDayIdx
    
    'Test the Fed/Sold Area
    For ilDatIdx = 0 To (UBound(tgDat) - 1) Step 1
        llDatFdSTime = gTimeToLong(tgDat(ilDatIdx).sFdSTime, False)
        llDatFdETime = gTimeToLong(tgDat(ilDatIdx).sFdETime, True)
            
        For ilDayIdx = 0 To 6 Step 1
            If tgDat(ilDatIdx).iFdDay(ilDayIdx) = 1 Then
                ilFound = False
                For ilHourIdx = 0 To UBound(tmFDDayTime(ilDayIdx).lSTime) - 1 Step 1
                    ilFound = False
                    If ((llDatFdSTime >= tmFDDayTime(ilDayIdx).lSTime(ilHourIdx)) And (llDatFdSTime < tmFDDayTime(ilDayIdx).lETime(ilHourIdx))) Then
                        ilFound = True
                        mTestForDPConflict = 1
                        Exit Function
                    End If
                    If ((llDatFdETime >= tmFDDayTime(ilDayIdx).lSTime(ilHourIdx)) And (llDatFdETime < tmFDDayTime(ilDayIdx).lETime(ilHourIdx))) Then
                        ilFound = True
                        mTestForDPConflict = 1
                        Exit Function
                    End If
                Next ilHourIdx
                If Not ilFound Then
                    tmFDDayTime(ilDayIdx).lSTime(UBound(tmFDDayTime(ilDayIdx).lSTime)) = llDatFdSTime
                    tmFDDayTime(ilDayIdx).lETime(UBound(tmFDDayTime(ilDayIdx).lSTime)) = llDatFdETime
                    ReDim Preserve tmFDDayTime(ilDayIdx).lSTime(0 To UBound(tmFDDayTime(ilDayIdx).lSTime) + 1)
                    ReDim Preserve tmFDDayTime(ilDayIdx).lETime(0 To UBound(tmFDDayTime(ilDayIdx).lETime) + 1)
                End If
            End If
        Next ilDayIdx
    Next ilDatIdx
    
    'Initialiaze Pledged daypart arrary
    For ilDayIdx = 0 To 6 Step 1
        ReDim tmPDDayTime(ilDayIdx).lSTime(0 To 0)
        ReDim tmPDDayTime(ilDayIdx).lETime(0 To 0)
    Next ilDayIdx
    
    'Test the Pledge Area
    For ilDatIdx = 0 To (UBound(tgDat) - 1) Step 1
        '9/28/06: TTP 2030- Don't check pledge info if Not Aired
        If tgStatusTypes(tgDat(ilDatIdx).iFdStatus).iPledged <= 1 Then
            llDatPdSTime = gTimeToLong(tgDat(ilDatIdx).sPdSTime, False)
            llDatPdETime = gTimeToLong(tgDat(ilDatIdx).sPdETime, True)
                
            For ilDayIdx = 0 To 6 Step 1
                If tgDat(ilDatIdx).iPdDay(ilDayIdx) = 1 Then
                    ilFound = False
                    For ilHourIdx = 0 To UBound(tmPDDayTime(ilDayIdx).lSTime) - 1 Step 1
                        ilFound = False
                        If ((llDatPdSTime >= tmPDDayTime(ilDayIdx).lSTime(ilHourIdx)) And (llDatPdSTime < tmPDDayTime(ilDayIdx).lETime(ilHourIdx))) Then
                            ilFound = True
                            mTestForDPConflict = 2
                            Exit Function
                        End If
                        If ((llDatPdETime >= tmPDDayTime(ilDayIdx).lSTime(ilHourIdx)) And (llDatPdETime < tmPDDayTime(ilDayIdx).lETime(ilHourIdx))) Then
                            ilFound = True
                            mTestForDPConflict = 2
                            Exit Function
                        End If
                    Next ilHourIdx
                    If Not ilFound Then
                        tmPDDayTime(ilDayIdx).lSTime(UBound(tmPDDayTime(ilDayIdx).lSTime)) = llDatPdSTime
                        tmPDDayTime(ilDayIdx).lETime(UBound(tmPDDayTime(ilDayIdx).lSTime)) = llDatPdETime
                        ReDim Preserve tmPDDayTime(ilDayIdx).lSTime(0 To UBound(tmPDDayTime(ilDayIdx).lSTime) + 1)
                        ReDim Preserve tmPDDayTime(ilDayIdx).lETime(0 To UBound(tmPDDayTime(ilDayIdx).lETime) + 1)
                    End If
                End If
            Next ilDayIdx
        End If
    Next ilDatIdx
    
    
    mTestForDPConflict = 0
End Function

Public Sub mVerifyPWExists()

    Dim rst_pw As ADODB.Recordset
    Dim slTemp As String
    
    If imShttCode = 0 Then
        Exit Sub
    End If
    
    SQLQuery = "SELECT attWebPW"
    SQLQuery = SQLQuery & " FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
    Set rst_pw = gSQLSelectCall(SQLQuery)
    
    If lmAttCode <> 0 Then
        If rst_pw.EOF = False Then
            slTemp = Trim$(rst_pw!attWebPW)
            If slTemp = "" Then
                SQLQuery = "SELECT shttWebPW, shttWebEmail"
                SQLQuery = SQLQuery & " FROM shtt"
                SQLQuery = SQLQuery + " WHERE (shttCode = " & imShttCode & ")"
                Set rst_pw = gSQLSelectCall(SQLQuery)
                
                If rst_pw.EOF = False Then
                    slTemp = Trim$(rst_pw!shttWebPW)
                    If slTemp = "" Then
'                        gMsgBox "First, a valid password must be defined at the Station level."
'                        Exit Sub
                        txtLogPassword.Text = Trim$(rst_pw!shttWebPW)
                        txtEmailAddr.Text = Trim$(rst_pw!shttWebEmail)
                        'SQLQuery = "UPDATE att SET "
                        'SQLQuery = SQLQuery + "attExportType = " & 0 & ""
                        'SQLQuery = SQLQuery + " WHERE attCode = " & lmAttCode & ""
                        'cnn.Execute SQLQuery, rdExecDirect
                        'rbcExportType(0).Value = True
                        'gMsgBox "Export type has been reset to Manual"
                    Else
                        txtLogPassword.Text = Trim$(rst_pw!shttWebPW)
                        txtEmailAddr.Text = Trim$(rst_pw!shttWebEmail)
                    End If
                End If
            End If
        End If
    Else
        SQLQuery = "SELECT shttWebPW, shttWebEmail"
        SQLQuery = SQLQuery & " FROM shtt"
        SQLQuery = SQLQuery + " WHERE (shttCode = " & imShttCode & ")"
        Set rst_pw = gSQLSelectCall(SQLQuery)
        
        If rst_pw.EOF = False Then
            slTemp = Trim$(rst_pw!shttWebPW)
            If slTemp = "" Then
'                gMsgBox "First, a valid password must be defined at the Station level."
'                Exit Sub
                txtLogPassword.Text = Trim$(rst_pw!shttWebPW)
                txtEmailAddr.Text = Trim$(rst_pw!shttWebEmail)
                'SQLQuery = "UPDATE att SET "
                'SQLQuery = SQLQuery + "attExportType = " & 0 & ""
                'SQLQuery = SQLQuery + " WHERE attCode = " & lmAttCode & ""
                'cnn.Execute SQLQuery, rdExecDirect
                'rbcExportType(0).Value = True
                'gMsgBox "Export type has been reset to Manual"
            Else
                txtLogPassword.Text = Trim$(rst_pw!shttWebPW)
                txtEmailAddr.Text = Trim$(rst_pw!shttWebEmail)
            End If
        End If
    End If
End Sub

Private Sub mAdjustDatTimes(lOffset As Long)

    Dim llLoop As Long
    Dim llPdTime As Long

    For llLoop = 0 To UBound(tgDat) - 1 Step 1
        llPdTime = gTimeToCurrency(tgDat(llLoop).sPdSTime, False)
        llPdTime = llPdTime + lOffset
        tgDat(llLoop).sPdSTime = gLongToTime(llPdTime)
        llPdTime = gTimeToCurrency(tgDat(llLoop).sPdETime, False)
        llPdTime = llPdTime + lOffset
        tgDat(llLoop).sPdETime = gLongToTime(llPdTime)
    Next llLoop


End Sub

Private Function mGetTimeOffSet(lTime1 As Long, lTime2 As Long) As Long

    'Pledged to air earlier then cd was supposed to
    If lTime1 > lTime2 Then
        mGetTimeOffSet = ((lTime1 - lTime2) * -1)
        Exit Function
    Else
        'Pledged to air later then cd was supposed to
        mGetTimeOffSet = lTime2 - lTime1
        Exit Function
    End If
    
    

End Function

Private Sub txtTime_Change()
    Dim slStr As String
    
    Select Case grdPledge.Col
        Case STARTTIMEFDINDEX, ENDTIMEFDINDEX, STARTTIMEPDINDEX, ENDTIMEPDINDEX
            slStr = Trim$(txtTime.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                grdPledge.CellForeColor = vbBlack
                slStr = gConvertTime(slStr)
                If Second(slStr) = 0 Then
                    slStr = Format$(slStr, sgShowTimeWOSecForm)
                Else
                    slStr = Format$(slStr, sgShowTimeWSecForm)
                End If
                If grdPledge.Text <> slStr Then
                    imFieldChgd = True
                    bmPledgeDataChgd = True
                End If
                grdPledge.Text = slStr
            End If
    End Select
End Sub

Private Sub txtTime_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mSetPledgeFromStatus(llRow As Long, llCol As Long)
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRowIndex As Long
    Dim ilLoop As Integer
    Dim ilDayOn As Integer
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llFdTime As Long
    Dim llPdTime As Long
    
    If (llCol <> STATUSINDEX) And (llCol <> STARTTIMEPDINDEX) And (llCol <> ENDTIMEPDINDEX) And (llCol <> -1) Then
        Exit Sub
    End If
    llSvRow = grdPledge.Row
    llSvCol = grdPledge.Col
    If (llCol = STATUSINDEX) Then
        slStr = txtDropdown.Text
    Else
        slStr = grdPledge.TextMatrix(llRow, STATUSINDEX)
    End If
    If (grdPledge.TextMatrix(llRow, STATUSINDEX) <> slStr) Or (llCol <> STATUSINDEX) Or (llCol = -1) Then
        imFieldChgd = True
        bmPledgeDataChgd = True
        llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
        If llRowIndex >= 0 Then
            ilIndex = lbcStatus.ItemData(llRowIndex)
            grdPledge.Row = llRow
            grdPledge.Col = AIRPLAYINDEX
            If (Trim$(grdPledge.TextMatrix(llRow, CODEINDEX)) <> "") Then
                grdPledge.CellBackColor = LIGHTYELLOW
            End If
            For ilLoop = MONPDINDEX To ESTIMATETIMEINDEX Step 1
                grdPledge.Col = ilLoop
                If (tmStatusTypes(ilIndex).iPledged <> 1) Then   'Delayed=1
                    grdPledge.CellBackColor = LIGHTYELLOW
                Else
                    grdPledge.CellBackColor = vbWhite
                    If ilLoop = ESTIMATETIMEINDEX Then
                        llFdTime = gTimeToLong(grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX), True) - gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX), False)
                        llPdTime = gTimeToLong(grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX), True) - gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX), False)
                        If llPdTime <= llFdTime Then
                            grdPledge.CellBackColor = LIGHTYELLOW
                        End If
                    End If
                End If
            Next ilLoop
            If (llCol = STARTTIMEPDINDEX) Or (llCol = ENDTIMEPDINDEX) Then
                grdPledge.Row = llSvRow
                grdPledge.Col = llSvCol
                Exit Sub
            End If
            If llCol <> -1 Then
                grdPledge.TextMatrix(llRow, STATUSINDEX) = lbcStatus.List(llRowIndex)
            End If
            If (tmStatusTypes(ilIndex).iPledged = 0) Then   'Live
                For ilLoop = MONFDINDEX To SUNFDINDEX Step 1
                    grdPledge.TextMatrix(llRow, ilLoop + MONPDINDEX - MONFDINDEX) = grdPledge.TextMatrix(llRow, ilLoop)
                Next ilLoop
                grdPledge.TextMatrix(llRow, DAYFEDINDEX) = ""
                grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX)
                grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX)
                grdPledge.TextMatrix(llRow, ESTIMATETIMEINDEX) = ""
            ElseIf (tmStatusTypes(ilIndex).iPledged = 1) Then   'Delayed
                ilDayOn = False
                For ilLoop = MONPDINDEX To SUNPDINDEX Step 1
                    If Trim$(grdPledge.TextMatrix(llRow, ilLoop)) <> "" Then
                        ilDayOn = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilDayOn Then
                    For ilLoop = MONFDINDEX To SUNFDINDEX Step 1
                        grdPledge.TextMatrix(llRow, ilLoop + MONPDINDEX - MONFDINDEX) = grdPledge.TextMatrix(llRow, ilLoop)
                    Next ilLoop
                End If
            ElseIf (tmStatusTypes(ilIndex).iPledged = 2) Then   'Not Aired
                For ilLoop = MONPDINDEX To SUNPDINDEX Step 1
                    grdPledge.TextMatrix(llRow, ilLoop) = ""
                Next ilLoop
                grdPledge.TextMatrix(llRow, DAYFEDINDEX) = ""
                grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = ""
                grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = ""
                grdPledge.TextMatrix(llRow, ESTIMATETIMEINDEX) = ""
            ElseIf (tmStatusTypes(ilIndex).iPledged = 3) Then   'No Pledged Times
                For ilLoop = MONPDINDEX To SUNPDINDEX Step 1
                    grdPledge.TextMatrix(llRow, ilLoop) = ""
                Next ilLoop
                grdPledge.TextMatrix(llRow, DAYFEDINDEX) = ""
                grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = ""
                grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = ""
                grdPledge.TextMatrix(llRow, ESTIMATETIMEINDEX) = ""
            End If
        End If
    End If
    grdPledge.Row = llSvRow
    grdPledge.Col = llSvCol
End Sub

Private Function mPledgeColAllowed(llCol As Long) As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim llSvCol As Long
    
    'If llCol = ESTIMATETIMEINDEX Then
    llSvCol = grdPledge.Col
        grdPledge.Col = llCol
        If grdPledge.CellBackColor = LIGHTYELLOW Then
            grdPledge.Col = llSvCol
            mPledgeColAllowed = False
            Exit Function
        End If
        grdPledge.Col = llSvCol
    'End If
    slStr = grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)
    llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        ilIndex = lbcStatus.ItemData(llRow)
        If (tmStatusTypes(ilIndex).iPledged = 0) Then   'Live
            If (llCol <= AIRPLAYINDEX) Or (llCol = EMBEDDEDORROSINDEX) Then
                mPledgeColAllowed = True
            Else
                mPledgeColAllowed = False
            End If
        ElseIf (tmStatusTypes(ilIndex).iPledged = 1) Then   'Delayed
            If llCol = DAYFEDINDEX Then
                If mPdPriorFd(grdPledge.Row) Then
                    mPledgeColAllowed = True
                Else
                    mPledgeColAllowed = False
                End If
            Else
                mPledgeColAllowed = True
            End If
        ElseIf (tmStatusTypes(ilIndex).iPledged = 2) Then   'Not Aired
            If (llCol <= STATUSINDEX) Or (llCol = EMBEDDEDORROSINDEX) Then
                mPledgeColAllowed = True
            Else
                mPledgeColAllowed = False
            End If
        ElseIf (tmStatusTypes(ilIndex).iPledged = 3) Then   'No Pledged Times
            If (llCol <= STATUSINDEX) Or (llCol = EMBEDDEDORROSINDEX) Then
                mPledgeColAllowed = True
            Else
                mPledgeColAllowed = False
            End If
        End If
    Else
        If (llCol <= STATUSINDEX) Or (llCol = EMBEDDEDORROSINDEX) Then
            mPledgeColAllowed = True
        Else
            mPledgeColAllowed = False
        End If
    End If

End Function


'D.S.

'Validate the left and right side of the agreement for problems
Private Function mValidateRows() As Integer

    Dim llCol As Long
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilPledgeCnt As Integer
    Dim ilFeedCnt As Integer
    Dim slMsg As String
    Dim llRowIndex As Long
    Dim iIndex As Integer
    Dim slStatus As String
    Dim ilRet As Integer

    mValidateRows = True
    
    For ilLoop = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        grdPledge.Row = ilLoop
        grdPledge.Col = STATUSINDEX
        'If grdPledge.Text <> "" Then
        slStatus = Trim$(grdPledge.Text)
        If slStatus <> "" Then
           ilFound = False
            'Check the feed cols to see if any are checked
            ilFeedCnt = 0
            For llCol = MONFDINDEX To SUNFDINDEX Step 1
                grdPledge.Col = llCol
                If grdPledge.Text <> "" Then
                    ilFeedCnt = ilFeedCnt + 1
                    ilFound = True
                    'Exit For
                End If
            Next llCol
            
            If Not ilFound Then
                mValidateRows = False
                gMsgBox "All Rows with Feed Times must have at least one day selected", vbCritical
                Exit Function
            End If
            
            '9/28/06: TTP 2030- Don't check pledge info if Not Aired
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStatus)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tmStatusTypes(iIndex).iPledged <= 1 Then

                    ilFound = False
                    'Check the pledge cols to see if any are checked
                    ilPledgeCnt = 0
                    For llCol = MONPDINDEX To SUNPDINDEX Step 1
                        grdPledge.Col = llCol
                        If grdPledge.Text <> "" Then
                            ilPledgeCnt = ilPledgeCnt + 1
                            ilFound = True
                            'Exit For
                        End If
                    Next llCol
                
                    If Not ilFound Then
                        mValidateRows = False
                        gMsgBox "All Rows with Pledge Times must have at least one day selected", vbCritical
                        Exit Function
                    End If
                    If ilPledgeCnt <> ilFeedCnt Then
                        'slMsg = "     *** Save Failed ***" & vbCrLf & vbCrLf
                        'slMsg = slMsg & "Check Marks Must Balance Between the Feed/Sold and the Pledge Days."
                        'If ilPledgeCnt > ilFeedCnt Then
                        '    slMsg = slMsg & vbCrLf & vbCrLf & "On Line " & CStr(ilLoop - 1) & " there are Fewer Days on the Feed/Sold side than the Pledged Side, Please Correct."
                        'Else
                        '    slMsg = slMsg & vbCrLf & vbCrLf & "On Line " & CStr(ilLoop - 1) & " there are More Days on the Feed/Sold side than the Pledged Side, Please Correct."
                        'End If
                        'gMsgBox slMsg
                        slMsg = "Check Marks don't Balance Between the Feed/Sold and the Pledge Days."
                        If ilPledgeCnt > ilFeedCnt Then
                            slMsg = slMsg & vbCrLf & vbCrLf & "On Line " & CStr(ilLoop - 1) & " there are Fewer Days on the Feed/Sold side than the Pledged Side, Continue?"
                        Else
                            slMsg = slMsg & vbCrLf & vbCrLf & "On Line " & CStr(ilLoop - 1) & " there are More Days on the Feed/Sold side than the Pledged Side, Continue?"
                        End If
                        ilRet = MsgBox(slMsg, vbYesNo, "Days Not Balanced")
                        If ilRet = vbNo Then
                            mValidateRows = False
                            Exit Function
                        Else
                            mValidateRows = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next ilLoop

End Function

Private Function mPostMonthly() As Boolean

    Dim llCol As Long
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slFeedStartTime As String
    Dim slFeedEndTime As String
    Dim slPledgeStartTime As String
    Dim slPledgeEndTime As String
    Dim llFeedTotalSeconds As Long
    Dim llPledgeTotalSeconds As Long
    Dim llRowIndex As Long
    Dim iIndex As Integer
    Dim slStatus As String

    On Error GoTo Err_Handler
    
    mPostMonthly = False

    For ilLoop = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        grdPledge.Row = ilLoop
        slStatus = grdPledge.TextMatrix(grdPledge.Row, STATUSINDEX)
        
        If slStatus <> "9-Not Carried" And slStatus <> "" Then
            slPledgeStartTime = grdPledge.TextMatrix(grdPledge.Row, STARTTIMEPDINDEX)
            slPledgeEndTime = grdPledge.TextMatrix(grdPledge.Row, ENDTIMEPDINDEX)
            
            slFeedStartTime = grdPledge.TextMatrix(grdPledge.Row, STARTTIMEFDINDEX)
            slFeedEndTime = grdPledge.TextMatrix(grdPledge.Row, ENDTIMEFDINDEX)
            
            llFeedTotalSeconds = DateDiff("S", slFeedStartTime, slFeedEndTime)
            llPledgeTotalSeconds = DateDiff("S", slPledgeStartTime, slPledgeEndTime)
            
            If llFeedTotalSeconds < llPledgeTotalSeconds Then
                Exit Function
            End If
        End If
    Next ilLoop

    mPostMonthly = True
    
    Exit Function
    
Err_Handler:
    gHandleError "AffErrorLog.txt", "frmAgmnt-mPostMonthly"
End Function

Private Function mPopMulticast() As Integer

    'D.S. 12/15/05 Fill the multicast listbox on the preference screen
    'Any station(s) that are defined in the Multicast Group and are associated with
    'the station the Agreement screen is on will appear inthe listbox.
    'If the station gets * next to it then it's agreemeent multicast question is set to yes
    'Also, build array "tmAssocStnMulticastInfo" for later use to loop through and see
    'if the agreement are all the same.
    
    Dim ilUpper As Integer
    Dim llTemp As Long
    Dim rst_Shtt As ADODB.Recordset
    Dim slName As String
    Dim slStr As String
    '8/6/19: Changed market name from slStr to slMarkrt
    Dim slMarket As String
    '8/6/19: Added Oswner
    Dim slOwner As String
    Dim llOwner As Long
    Dim slSelectedMarket As String
    Dim slSelectedOwner As String
    
    Dim tmp_rst As ADODB.Recordset
    Dim ilFound As Integer
    Dim slAttOffAir As String
    Dim slAttOnAir As String
    Dim slOnAir As String
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llColor As Long
    Dim slOff As String
    Dim slOffAir As String
    Dim slDropDate As String
    
    On Error GoTo ErrHand
    'frcMulticast.Visible = False
    lacMulticast(0).Visible = False
    lacMulticast(1).Visible = False
    'lbcMulticast.Visible = False
    grdMulticast.Visible = False
    mPopMulticast = False
    lbcMulticast.Clear
    grdMulticast.Clear
    grdMulticast.Rows = 2
    grdMulticast.Redraw = False
    gGrid_Clear grdMulticast, True
    For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
        grdMulticast.Row = llRow
        For llCol = MCCALLLETTERSINDEX To MCDATERANGEINDEX Step 1
            grdMulticast.Col = llCol
            grdMulticast.CellBackColor = vbWhite
        Next llCol
    Next llRow
    If (imShttCode <= 0) Or (imVefCode <= 0) Then
        Exit Function
    End If
    llRow = grdMulticast.FixedRows
    slOnAir = txtOnAirDate.Text
    slOffAir = txtOffAirDate.Text
    If slOffAir = "" Then
        slOffAir = "12/31/2069"
    End If
    slDropDate = txtDropDate.Text
    If slDropDate = "" Then
        slDropDate = "12/31/2069"
    End If
    If DateValue(gAdjYear(slOffAir)) <= DateValue(gAdjYear(slDropDate)) Then
        slOff = slOffAir
    Else
        slOff = slDropDate
    End If
    If cboPSSort.Text <> "" And cboSSSort.Text <> "" And slOnAir <> "" Then
        ilUpper = 0
        If gIsMulticast(imShttCode) Then
        
            llTemp = gBinarySearchStationInfoByCode(imShttCode)
            If llTemp <> -1 Then
                slSelectedMarket = Trim$(tgStationInfoByCode(llTemp).sMarket)
            Else
                slSelectedMarket = "Market Missing"
            End If
            llTemp = gBinarySearchStationInfoByCode(imShttCode)
            If llTemp <> -1 Then
                llOwner = gBinarySearchOwner(tgStationInfoByCode(llTemp).lOwnerCode)
                If llOwner <> -1 Then
                    slSelectedOwner = Trim$(tgOwnerInfo(llOwner).sName)
                Else
                    slSelectedOwner = "Missing Owner"
                End If
            Else
                slSelectedOwner = "Missing Owner"
            End If
        
            'frcMulticast.Visible = True
            mSetMulticast
            'lbcMulticast.Clear
            lacMulticast(0).Caption = "Multicast " & gGetCallLettersByShttCode(imShttCode) & " with:"
            lacMulticast(1).Caption = lacMulticast(0).Caption
            'Get the multicast group ID for the current station, if it has one
            llTemp = gGetStaMulticastGroupID(imShttCode)
            If llTemp <> 0 Then
                'Select all of the stations with the same group ID
                SQLQuery = "SELECT shttCode, shttOwnerArttCode from shtt WHERE (shttMultiCastGroupID = " & llTemp & ")"
                Set rst_Shtt = gSQLSelectCall(SQLQuery)
                Do While Not rst_Shtt.EOF
                    'frcMulticast.Visible = True
                    'Get the market code for each station that has the above group ID, they may be different
                    If rst_Shtt!shttCode <> imShttCode Then
                        slName = gGetCallLettersByShttCode(rst_Shtt!shttCode)
                        llTemp = gGetStaMarketCode(rst_Shtt!shttCode)
                        slMarket = gGetStaMarketName(llTemp)
                        '8/6/19: add owner
                        llOwner = gBinarySearchOwner(rst_Shtt!shttOwnerArttCode)
                        If llOwner <> -1 Then
                            slOwner = Trim$(tgOwnerInfo(llOwner).sName)
                        Else
                            slOwner = "Missing Owner"
                        End If
                        'Find out if there is a current agreement for this station vehicle combination
                        SQLQuery = "SELECT attCode, attShfCode, attMulticast, attOnAir, attOffAir, attDropDate FROM att"
                        SQLQuery = SQLQuery & " WHERE (attShfCode = " & rst_Shtt!shttCode & " And attVefCode = " & imVefCode
                        'SQLQuery = SQLQuery & " AND (attOnAir <= " & "'" & Format$(Now(), sgSQLDateForm) & "'"
                        'SQLQuery = SQLQuery & " AND attOffAir >= " & "'" & Format$(Now(), sgSQLDateForm) & "'"
                        ''D.S. 5/15/09 Don't show cancel before starts
                        SQLQuery = SQLQuery & " AND attOffAir >= attOnAir"
                        'SQLQuery = SQLQuery & " AND attDropDate >= " & "'" & Format$(Now(), sgSQLDateForm) & "'" & ")"
                        SQLQuery = SQLQuery & " AND attOnAir <= " & "'" & Format$(slOff, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery & " AND (attOffAir >= " & "'" & Format$(slOnAir, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery & " AND attDropDate >= " & "'" & Format$(slOnAir, sgSQLDateForm) & "'" & ")" & ")"
                        Set tmp_rst = gSQLSelectCall(SQLQuery)
                        ilFound = False
                        Do While Not tmp_rst.EOF
                            ilFound = True
                            llColor = vbWhite
                            slAttOnAir = gAdjYear(tmp_rst!attOnAir)
                            If (DateValue(gAdjYear(tmp_rst!attOffAir)) <= DateValue(gAdjYear(tmp_rst!attDropDate))) Then
                                slAttOffAir = gAdjYear(tmp_rst!attOffAir)
                                If slAttOffAir = "12/31/2069" Then
                                    slAttOffAir = "TFN"
                                Else
                                    slAttOffAir = Format$(slAttOffAir, "m/d/yy")
                                    llColor = LIGHTYELLOW
                                End If
                            Else
                                slAttOffAir = gAdjYear(tmp_rst!attDropDate)
                                If slAttOffAir = "12/31/2069" Then
                                    slAttOffAir = "TFN"
                                Else
                                    slAttOffAir = Format$(slAttOffAir, "m/d/yy")
                                    llColor = LIGHTYELLOW
                                End If
                            End If
                            If llRow + 1 > grdMulticast.Rows Then
                                grdMulticast.AddItem ""
                            End If
                            grdMulticast.Row = llRow
                            For llCol = MCCALLLETTERSINDEX To MCDATERANGEINDEX Step 1
                                grdMulticast.Col = llCol
                                grdMulticast.CellBackColor = llColor
                            Next llCol
                            grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX) = slName
                            '8/6/19: changed name from slStr to Market
                            grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = slMarket
                            If slSelectedMarket <> slMarket Then
                                grdMulticast.Row = llRow
                                grdMulticast.Col = MCMARKETINDEX
                                grdMulticast.CellForeColor = vbRed
                            End If
                            '8/6/19: Add Owner
                            grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = slOwner
                            If slSelectedOwner <> slOwner Then
                                grdMulticast.Row = llRow
                                grdMulticast.Col = MCOWNERINDEX
                                grdMulticast.CellForeColor = vbRed
                            End If
                            grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0"
                            If Trim$(tmp_rst!attMulticast = "Y") Then
                                grdMulticast.TextMatrix(llRow, MCWITHINDEX) = "Multicast"
                                If smSvMulticast = "Y" Then
                                    grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1"
                                End If
                            Else
                                grdMulticast.TextMatrix(llRow, MCWITHINDEX) = "Not M'cast"
                            End If
                            grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX) = slAttOnAir & "-" & slAttOffAir
                            grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) = tmp_rst!attshfCode
                            grdMulticast.TextMatrix(llRow, MCATTCODEINDEX) = tmp_rst!attCode
                            llRow = llRow + 1
                            tmp_rst.MoveNext
                        Loop
                        If Not ilFound Then
                            If llRow + 1 > grdMulticast.Rows Then
                                grdMulticast.AddItem ""
                            End If
                            grdMulticast.Row = llRow
                            grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX) = slName
                            '8/6/19: changed name from slStr to Market
                            grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = slMarket
                            If slSelectedMarket <> slMarket Then
                                grdMulticast.Row = llRow
                                grdMulticast.Col = MCMARKETINDEX
                                grdMulticast.CellForeColor = vbRed
                            End If
                            '8/6/19: Add Owner
                            grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = slOwner
                            If slSelectedOwner <> slOwner Then
                                grdMulticast.Row = llRow
                                grdMulticast.Col = MCOWNERINDEX
                                grdMulticast.CellForeColor = vbRed
                            End If
                            grdMulticast.TextMatrix(llRow, MCWITHINDEX) = "Not M'cast"
                            grdMulticast.TextMatrix(llRow, MCDATERANGEINDEX) = "No Agreement"
                            grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) = rst_Shtt!shttCode
                            grdMulticast.TextMatrix(llRow, MCATTCODEINDEX) = 0
                            grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0"
                            llRow = llRow + 1
                        End If
                    End If
                    rst_Shtt.MoveNext
                Loop
            End If
            If smSvMulticast = "Y" Then
                For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
                    If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                        mMCPaintRowColor llRow
                    End If
                Next llRow
            End If
        'Else
        '    frcMulticast.Visible = False
        End If
    End If
    grdMulticast.Redraw = True

    mPopMulticast = True
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", ""
    grdMulticast.Redraw = True
    imInChg = False
End Function

Private Function mPdPriorFd(llRow As Long) As Integer
    Dim ilFirstFdDay As Integer
    Dim ilFirstPdDay As Integer
    Dim ilDay As Integer
    Dim slStr As String
    Dim llRowIndex As Long
    Dim ilIndex As Integer
    
    mPdPriorFd = False
    slStr = grdPledge.TextMatrix(llRow, STATUSINDEX)
    llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
    If llRowIndex >= 0 Then
        ilIndex = lbcStatus.ItemData(llRowIndex)
        If tmStatusTypes(ilIndex).iPledged <> 1 Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    ilFirstFdDay = -1
    For ilDay = MONFDINDEX To SUNFDINDEX Step 1
        If Trim$(grdPledge.TextMatrix(llRow, ilDay)) <> "" Then
            If ilFirstFdDay = -1 Then
                ilFirstFdDay = ilDay
                Exit For
            End If
        End If
    Next ilDay
    ilFirstPdDay = -1
    For ilDay = MONPDINDEX To SUNPDINDEX Step 1
        If Trim$(grdPledge.TextMatrix(llRow, ilDay)) <> "" Then
            If ilFirstPdDay = -1 Then
                ilFirstPdDay = ilDay
                Exit For
            End If
        End If
    Next ilDay
    If ilFirstPdDay - MONPDINDEX < ilFirstFdDay - MONFDINDEX Then
        mPdPriorFd = True
    End If
End Function
Private Sub mGetAffAE()
    'Replaced by Market Rep, attMktRepUstCode
    Dim slName As String
    
    On Error GoTo ErrHand
    
    cboAffAE.Clear
    SQLQuery = "SELECT * FROM artt WHERE arttType = 'R'"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        slName = Trim$(Trim$(rst!arttFirstName) & " " & rst!arttLastName)
        cboAffAE.AddItem Trim$(slName)
        cboAffAE.ItemData(cboAffAE.NewIndex) = rst!arttCode
        rst.MoveNext
    Wend
    cboAffAE.AddItem "[None]", 0
    cboAffAE.ItemData(cboAffAE.NewIndex) = 0
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mGetAffiliateAE"
End Sub

Private Sub mSetInitAffAE()
    'Replaced by Market Rep, attMktRepUstCode
    Dim llMktCode As Long
    Dim llMntCode As Long
    Dim ilRet As Integer
    Dim llArttCode As Long
    Dim ilLoop As Integer
        
    If (imShttCode = 0) Or (imVefCode = 0) Then
        Exit Sub
    End If
    On Error GoTo ErrHand
    llArttCode = -1
    SQLQuery = "Select shttMktCode, shttMntCode FROM shtt where shttCode = " & imShttCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        llMktCode = rst!shttMktCode
        llMntCode = rst!shttMntCode
        If llMktCode > 0 Then
            SQLQuery = "SELECT matArttCode FROM mat WHERE matMktCode = " & llMktCode
            SQLQuery = SQLQuery & " AND matVefCode = " & imVefCode
            SQLQuery = SQLQuery & " AND matMntCode = " & llMntCode
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                llArttCode = rst!matArttCode
            Else
                SQLQuery = "SELECT matArttCode FROM mat WHERE matMktCode = " & llMktCode
                SQLQuery = SQLQuery & " AND matVefCode = " & imVefCode
                SQLQuery = SQLQuery & " AND matMntCode = " & 0
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    llArttCode = rst!matArttCode
                Else
                    SQLQuery = "SELECT matArttCode FROM mat WHERE matMktCode = " & llMktCode
                    SQLQuery = SQLQuery & " AND matVefCode = " & 0
                    SQLQuery = SQLQuery & " AND matMntCode = " & 0
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        llArttCode = rst!matArttCode
                    End If
                End If
            End If
        End If
    End If
    If llArttCode <> -1 Then
        For ilLoop = 0 To cboAffAE.ListCount - 1 Step 1
            If cboAffAE.ItemData(ilLoop) = llArttCode Then
                cboAffAE.ListIndex = ilLoop
                cboAffAE.Text = Trim$(cboAffAE.List(ilLoop))
                Exit For
            End If
        Next ilLoop
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mSetInitAffAE"
End Sub

Private Function mSetUsedForAtt(ilShttCode As Integer, blTestDate As Boolean) As Integer
    Dim llOffDate As Long
    Dim llDropDate As Long
    Dim llNowDate As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand:
    
    llNowDate = DateValue(gAdjYear(Format(gNow(), sgShowDateForm)))
    If (Trim$(txtOffAirDate.Text) <> "") And (blTestDate) Then
        llOffDate = DateValue(gAdjYear(Trim$(txtOffAirDate.Text)))
    Else
        llOffDate = llNowDate + 1
    End If
    If (Trim$(txtDropDate.Text) <> "") And (blTestDate) Then
        llDropDate = DateValue(gAdjYear(Trim$(txtDropDate.Text)))
    Else
        llDropDate = llNowDate + 1
    End If
    If (llNowDate <= llOffDate) And (llNowDate <= llDropDate) Then
        slSQLQuery = "UPDATE shtt SET shttUsedForAtt = 'Y', shttAgreementExist = 'Y'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & ilShttCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSetUsedForAtt"
            mSetUsedForAtt = False
            Exit Function
        End If
        '11/26/16
        mUpdateShttTables ilShttCode, True, True, ""
    Else
        slSQLQuery = "UPDATE shtt SET shttAgreementExist = 'Y'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & ilShttCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSetUsedForAtt"
            mSetUsedForAtt = False
            Exit Function
        End If
        '11/26/16
        mUpdateShttTables ilShttCode, False, True, ""
    End If
    mSetUsedForAtt = True
    Exit Function
    
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mSetUsedForAtt"
    mSetUsedForAtt = False
End Function

Private Sub txtXDReceiverID_Change()
    imFieldChgd = True
End Sub

Private Sub txtXDReceiverID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtXDReceiverID_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub mPresetAgreement()
    Dim llRow As Long
    'DoEvents
    llRow = SendMessageByString(cboPSSort.hwnd, CB_FINDSTRING, -1, sgTCCallLetters)
    If llRow >= 0 Then
        cboPSSort.ListIndex = llRow
        If lgTCAttCode > 0 Then
            For llRow = 0 To cboSSSort.ListCount - 1 Step 1
                If lgTCAttCode = cboSSSort.ItemData(llRow) Then
                    cboSSSort.ListIndex = llRow
                    Exit For
                End If
            Next llRow
        End If
        If cboSSSort.ListCount <= 0 Then
            optExAll(1).Value = True
        End If
        If cboSSSort.Enabled Then
            cboSSSort.SetFocus
        Else
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        End If
    Else
        If cboPSSort.Enabled Then
            cboPSSort.SetFocus
        End If
    End If

End Sub

Private Sub mPopServiceRep()
    Dim slServiceRepName As String
    Dim ilServRepUstCode As Integer
    Dim ilUst As Integer
    
    On Error GoTo ErrHand
    
    If cbcServiceRep.ListIndex > 1 Then
        slServiceRepName = Trim$(cbcServiceRep.Text)
        ilServRepUstCode = cbcServiceRep.GetItemData(cbcServiceRep.ListIndex)
    Else
        slServiceRepName = ""
        ilServRepUstCode = -2
    End If


    cbcServiceRep.Clear
    cbcServiceRep.AddItem ("[Same as Station]")
    cbcServiceRep.SetItemData = 0
    SQLQuery = "SELECT UstName, ustReportName, ustCode FROM Ust INNER JOIN Dnt ON ustDntCode = dntCode WHERE dntType = " & "'S'" & " ORDER BY UstName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If Trim$(rst!ustReportName) <> "" Then
            cbcServiceRep.AddItem Trim$(rst!ustReportName)
        Else
            cbcServiceRep.AddItem Trim$(rst!ustname)
        End If
        cbcServiceRep.SetItemData = rst!ustCode
        rst.MoveNext
    Loop
    'If slServiceRepName <> "" Then
    If ilServRepUstCode > 0 Then
        For ilUst = 0 To cbcServiceRep.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcServiceRep.GetName(ilUst)), slServiceRepName, vbTextCompare) = 0 Then
            If cbcServiceRep.GetItemData(ilUst) = ilServRepUstCode Then
                cbcServiceRep.SetListIndex = ilUst
                Exit For
            End If
        Next ilUst
    End If
    
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mPopServiceRep"
End Sub
Private Sub mPopMarketRep()
    Dim slMarketRepName As String
    Dim ilMktRepUstCode As Integer
    Dim ilUst As Integer
    
    On Error GoTo ErrHand
    
    If cbcMarketRep.ListIndex > 1 Then
        slMarketRepName = Trim$(cbcMarketRep.Text)
        ilMktRepUstCode = cbcMarketRep.GetItemData(cbcMarketRep.ListIndex)
    Else
        slMarketRepName = ""
        ilMktRepUstCode = -2
    End If


    cbcMarketRep.Clear
    cbcMarketRep.AddItem ("[Same as Station]")
    cbcMarketRep.SetItemData = 0
    SQLQuery = "SELECT UstName, ustReportName, ustCode FROM Ust INNER JOIN Dnt ON ustDntCode = dntCode WHERE dntType = " & "'M'" & " ORDER BY UstName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If Trim$(rst!ustReportName) <> "" Then
            cbcMarketRep.AddItem Trim$(rst!ustReportName)
        Else
            cbcMarketRep.AddItem Trim$(rst!ustname)
        End If
        cbcMarketRep.SetItemData = rst!ustCode
        rst.MoveNext
    Loop
    'If slMarketRepName <> "" Then
    If ilMktRepUstCode > 0 Then
        For ilUst = 0 To cbcMarketRep.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcMarketRep.GetName(ilUst)), slMarketRepName, vbTextCompare) = 0 Then
            If cbcMarketRep.GetItemData(ilUst) = ilMktRepUstCode Then
                cbcMarketRep.SetListIndex = ilUst
                Exit For
            End If
        Next ilUst
    End If
    
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mPopMarketRep"
End Sub

Private Sub mGetShttInfo()
    Dim llLoop As Long
    On Error GoTo ErrHand
    txtPassword.Text = ""
    lacStationMarketRep.Caption = "Station Market Rep:"
    lacStationServiceRep.Caption = "Station Service Rep:"
    txtHistorialDate.Text = ""
    If imShttCode <= 0 Then
        Exit Sub
    End If
    smPassword = gGetStationPW(imShttCode)
    
    SQLQuery = "SELECT shttWebPW, shttWebEmail, shttMktRepUstCode, shttServRepUstCode, shttHistStartDate"
    SQLQuery = SQLQuery & " FROM shtt"
    SQLQuery = SQLQuery + " WHERE (shttCode = " & imShttCode & ")"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        smPassword = Trim$(rst_Shtt!shttWebPW)
        txtPassword.Text = smPassword
        If (DateValue(gAdjYear(rst_Shtt!shttHistStartDate)) = DateValue("1/1/1970")) Then  'Or (rst!attOnAir = "1/1/70") Then
            txtHistorialDate.Text = ""
        Else
            txtHistorialDate.Text = Format$(rst_Shtt!shttHistStartDate, sgShowDateForm)
        End If
        If rst_Shtt!shttMktRepUstCode > 0 Then
            For llLoop = 1 To cbcMarketRep.ListCount - 1 Step 1
                If cbcMarketRep.GetItemData(CInt(llLoop)) = rst_Shtt!shttMktRepUstCode Then
                    lacStationMarketRep.Caption = "Station Market Rep: " & cbcMarketRep.GetName(CLng(llLoop))
                    Exit For
                End If
            Next llLoop
        End If
        
        If rst_Shtt!shttServRepUstCode > 0 Then
            For llLoop = 1 To cbcServiceRep.ListCount - 1 Step 1
                If cbcServiceRep.GetItemData(CInt(llLoop)) = rst_Shtt!shttServRepUstCode Then
                    lacStationServiceRep.Caption = "Station Service Rep: " & cbcServiceRep.GetName(CLng(llLoop))
                    Exit For
                End If
            Next llLoop
        End If
    Else
        smPassword = ""
        txtPassword.Text = smPassword
    End If
    
    
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mGetShttInfo"
End Sub

Private Sub mPopContractPDF(blObtainSubfolder As Boolean)
    Dim llLoop As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim slContractPDF As String
    Dim ilPos As Integer
    Dim llPftCode As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    cbcContractPDF.Clear
    If sgContractPDFPath = "" Then
        Exit Sub
    End If
    If imVefCode <= 0 Then
        Exit Sub
    End If
    If imShttCode <= 0 Then
        Exit Sub
    End If
    cbcContractPDF.AddItem "[None]"
    cbcContractPDF.SetItemData = 0
    slContractPDF = ""
    llPftCode = 0
    If blObtainSubfolder Then
        smContractPDFSubFolder = ""
    End If
    If (lmAttCode > 0) And (blObtainSubfolder) Then
        SQLQuery = "SELECT * FROM pft"
        SQLQuery = SQLQuery + " WHERE (pftAttCode = " & lmAttCode & ")"
        Set rst_Pft = gSQLSelectCall(SQLQuery)
        If Not rst_Pft.EOF Then
            llPftCode = rst_Pft!PftCode
            slContractPDF = rst_Pft!pftPDFName
            ilPos = InStrRev(slContractPDF, "\")
            If ilPos > 0 Then
                smContractPDFSubFolder = Left$(slContractPDF, ilPos)
                slContractPDF = Mid$(slContractPDF, ilPos + 1)
            End If
        End If
    End If
    SQLQuery = "SELECT * FROM pft"
    SQLQuery = SQLQuery + " WHERE (pftShttCode = " & imShttCode & " AND pftVefCode = " & imVefCode & ")"
    Set rst_Pft = gSQLSelectCall(SQLQuery)
    Do While Not rst_Pft.EOF
        slStr = Trim$(rst_Pft!pftPDFName)
        ilPos = InStrRev(slStr, "\")
        If ilPos > 0 Then
            If smContractPDFSubFolder = Left$(slStr, ilPos) Then
                slStr = Mid$(slStr, ilPos + 1)
            Else
                slStr = ""
            End If
        End If
        If slStr <> "" Then
            ilFound = False
            For ilLoop = 0 To cbcContractPDF.ListCount - 1 Step 1
                If StrComp(UCase$(Trim$(slStr)), UCase$(Trim$(cbcContractPDF.GetName(ilLoop))), vbTextCompare) = 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                cbcContractPDF.AddItem slStr
                cbcContractPDF.SetItemData = rst_Pft!PftCode
            End If
        End If
        rst_Pft.MoveNext
    Loop
    
    slStr = sgContractPDFPath & smContractPDFSubFolder
    On Error GoTo ErrHand1
    lbcContractPDFList.Path = Left$(slStr, Len(slStr) - 1)
    On Error GoTo ErrHand
    lbcContractPDFList.Pattern = "*.PDF"
    For llLoop = 0 To lbcContractPDFList.ListCount - 1 Step 1
        slStr = Trim$(lbcContractPDFList.List(llLoop))
        SQLQuery = "SELECT * FROM pft"
        SQLQuery = SQLQuery + " WHERE (pftPDFName = '" & smContractPDFSubFolder & slStr & "')"
        Set rst_Pft = gSQLSelectCall(SQLQuery)
        If rst_Pft.EOF Then
            ilFound = False
            For ilLoop = 0 To cbcContractPDF.ListCount - 1 Step 1
                If StrComp(UCase$(Trim$(slStr)), UCase$(Trim$(cbcContractPDF.GetName(ilLoop))), vbTextCompare) = 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                cbcContractPDF.AddItem slStr
                cbcContractPDF.SetItemData = 0
            End If
        End If
    Next llLoop
    cbcContractPDF.SetListIndex = -1
    'If lmAttCode > 0 Then
    '    SQLQuery = "SELECT * FROM pft"
    '    SQLQuery = SQLQuery + " WHERE (pftAttCode = " & lmAttCode & ")"
    '    Set rst_Pft = gSQLSelectCall(SQLQuery)
    '    If Not rst_Pft.EOF Then
    If Trim$(smContractPDFSubFolder) = "" Then
        If sgContractPDFPath = "" Then
            lacPDFPath.Caption = "PDF Path: Not defined in Affliat.ini"
        Else
            lacPDFPath.Caption = "PDF Path: " '& sgContractPDFPath
        End If
        If slContractPDF <> "" Then
            For llLoop = 0 To cbcContractPDF.ListCount - 1 Step 1
                If cbcContractPDF.GetItemData(CInt(llLoop)) = llPftCode Then
                    cbcContractPDF.SetListIndex = llLoop
                    Exit For
                End If
            Next llLoop
        End If
    Else
        lacPDFPath.Caption = "PDF Path: " & Trim$(smContractPDFSubFolder)
    End If
    '    End If
    'End If
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mPopContractPDF"
    Exit Sub
ErrHand1:
    MsgBox "ContractPDF = " & sgContractPDFPath & " Drive\Path definition in the Affiliat.ini is invalid, please Fix"
    lacContractPDF.Enabled = False
    cbcContractPDF.Enabled False
    cmcBrowse.Enabled = False
    Exit Sub
End Sub

Private Function mSaveContractPDF(llAttCode As Long, ilShttCode As Integer, ilVefCode As Integer) As Integer
    Dim slPDFName As String

    On Error GoTo ErrHand
    mSaveContractPDF = True
    If cbcContractPDF.ListIndex < 0 Then
        Exit Function
    End If
    If (llAttCode <= 0) Or (ilShttCode <= 0) Or (ilVefCode <= 0) Then
        Exit Function
    End If
    If llAttCode > 0 Then
        SQLQuery = "DELETE FROM pft WHERE (pftAttCode = " & llAttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSaveContractPDF"
            mSaveContractPDF = False
            Exit Function
        End If
    End If
    If cbcContractPDF.ListIndex <= 0 Then
        Exit Function
    End If
    slPDFName = Trim$(cbcContractPDF.GetName(cbcContractPDF.ListIndex))
    If smContractPDFSubFolder <> "" Then
        slPDFName = smContractPDFSubFolder  '& slPDFName
    Else
    End If
    SQLQuery = "Insert Into pft ( "
    SQLQuery = SQLQuery & "pftAttCode, "
    SQLQuery = SQLQuery & "pftShttCode, "
    SQLQuery = SQLQuery & "pftVefCode, "
    SQLQuery = SQLQuery & "pftPDFName, "
    SQLQuery = SQLQuery & "pftDateEntered, "
    SQLQuery = SQLQuery & "pftUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & llAttCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slPDFName) & "', "
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        mMousePointer vbDefault
        gHandleError "AffErrorLog.txt", "AffAgmnt-mSaveContractPDF"
        mSaveContractPDF = False
        Exit Function
    End If
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mSaveContractPDF"
    mSaveContractPDF = False
End Function

Private Sub mPopPledge()
    Dim llRow As Long
    Dim iLoop As Integer
    Dim iDay As Integer
    ReDim iFdDay(0 To 6) As Integer
    ReDim iPdDay(0 To 6) As Integer
    Dim sStatus As String
    Dim llCol As Long
    Dim slStr As String
    Dim ilNext As Integer
    
    grdPledge.Redraw = False
    gGrid_Clear grdPledge, True
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        grdPledge.Row = llRow
        For llCol = MONFDINDEX To ESTIMATETIMEINDEX Step 1
            grdPledge.Col = llCol
            grdPledge.CellBackColor = vbWhite
            grdPledge.CellForeColor = vbBlack
        Next llCol
    Next llRow
    llRow = grdPledge.FixedRows
    If UBound(tgDat) > LBound(tgDat) Then
        For iLoop = 0 To UBound(tgDat) - 1 Step 1
            For iDay = 0 To 6 Step 1
                If tgDat(iLoop).iFdDay(iDay) = 1 Then
                    iFdDay(iDay) = True
                Else
                    iFdDay(iDay) = False
                End If
                If tgDat(iLoop).iPdDay(iDay) = 1 Then
                    iPdDay(iDay) = True
                Else
                    iPdDay(iDay) = False
                End If
            Next iDay
            'If Not ilAdjustCDStartTime And Not optTimeType(0).Value Then
                sStatus = Trim$(tmStatusTypes(tgDat(iLoop).iFdStatus).sName)
            'End If
            If llRow + 1 > grdPledge.Rows Then
                grdPledge.AddItem ""
            End If
            grdPledge.Row = llRow
            For llCol = MONFDINDEX To SUNFDINDEX Step 1
                grdPledge.Col = llCol
                grdPledge.CellBackColor = LIGHTYELLOW
                grdPledge.CellFontName = "Monotype Sorts"
                If iFdDay(llCol - MONFDINDEX) Then
                    grdPledge.Text = "4"
                Else
                    grdPledge.Text = ""
                End If
            Next llCol
            grdPledge.Col = STARTTIMEFDINDEX
            grdPledge.CellBackColor = LIGHTYELLOW
            grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX) = Trim$(tgDat(iLoop).sFdSTime)
            grdPledge.Col = ENDTIMEFDINDEX
            grdPledge.CellBackColor = LIGHTYELLOW
            grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX) = Trim$(tgDat(iLoop).sFdETime)
            grdPledge.TextMatrix(llRow, STATUSINDEX) = Trim$(sStatus)
            If tgDat(iLoop).iAirPlayNo > 0 Then
                grdPledge.TextMatrix(llRow, AIRPLAYINDEX) = tgDat(iLoop).iAirPlayNo
            Else
                grdPledge.TextMatrix(llRow, AIRPLAYINDEX) = ""
            End If
            For llCol = MONPDINDEX To SUNPDINDEX Step 1
                grdPledge.Col = llCol
                grdPledge.CellFontName = "Monotype Sorts"
                If iPdDay(llCol - MONPDINDEX) Then
                    grdPledge.Text = "4"
                Else
                    grdPledge.Text = ""
                End If
            Next llCol
            grdPledge.TextMatrix(llRow, DAYFEDINDEX) = Trim$(tgDat(iLoop).sPdDayFed)
            grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX) = Trim$(tgDat(iLoop).sPdSTime)
            grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX) = Trim$(tgDat(iLoop).sPdETime)
            grdPledge.Col = ESTIMATETIMEINDEX
            grdPledge.TextMatrix(llRow, ESTIMATETIMEINDEX) = ""
            If tgDat(iLoop).sEstimatedTime = "Y" Then
                grdPledge.CellBackColor = vbWhite
                grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX) = tgDat(iLoop).iFirstET
                slStr = ""
                ilNext = tgDat(iLoop).iFirstET
                Do While ilNext <> -1
                    If (Trim$(tmETAvailInfo(ilNext).sETTime) <> "") Or (Trim$(tmETAvailInfo(ilNext).sETDay) <> "") Then
                        If slStr = "" Then
                            slStr = tmETAvailInfo(ilNext).sETDay & ":" & tmETAvailInfo(ilNext).sETTime
                        Else
                            slStr = slStr & ", " & tmETAvailInfo(ilNext).sETDay & ":" & tmETAvailInfo(ilNext).sETTime
                        End If
                    End If
                    ilNext = tmETAvailInfo(ilNext).iNextET
                Loop
                grdPledge.TextMatrix(llRow, ESTIMATETIMEINDEX) = slStr
            Else
                grdPledge.CellBackColor = LIGHTYELLOW
                grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX) = ""
            End If
            grdPledge.TextMatrix(llRow, EMBEDDEDORROSINDEX) = Trim$(tgDat(iLoop).sEmbeddedOrROS)
            grdPledge.TextMatrix(llRow, CODEINDEX) = Trim$(Str$(tgDat(iLoop).lCode))
            mSetPledgeFromStatus llRow, -1
            llRow = llRow + 1
        Next iLoop
        If llRow >= grdPledge.Rows Then
            grdPledge.AddItem ""
            grdPledge.Row = llRow
            For llCol = MONFDINDEX To SUNFDINDEX Step 1
                grdPledge.Col = llCol
                grdPledge.CellFontName = "Monotype Sorts"
            Next llCol
            For llCol = MONPDINDEX To SUNPDINDEX Step 1
                grdPledge.Col = llCol
                grdPledge.CellFontName = "Monotype Sorts"
            Next llCol
        End If
        'For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        '    If grdPledge.TextMatrix(llRow, STATUSINDEX) = "" Then
        '        grdPledge.Row = llRow
        '        grdPledge.Col = ESTIMATETIMEINDEX
        '        grdPledge.CellBackColor = LIGHTYELLOW
        '    End If
        'Next llRow
    End If
    imLastPledgeSort = -1
    imLastPledgeColSorted = -1
    mPledgeSortCol STARTTIMEFDINDEX
    grdPledge.TopRow = grdPledge.FixedRows
    grdPledge.Redraw = True
End Sub

Private Function mLoadPledgeOk() As Integer
    mLoadPledgeOk = False
    If imShttCode <= 0 Then
        gMsgBox "Station must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        Exit Function
    End If
    If imVefCode <= 0 Then
        gMsgBox "Vehicle must be selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then
            If cboPSSort.Enabled Then
                cboPSSort.SetFocus
            End If
        Else
            If cboSSSort.Enabled Then
                cboSSSort.SetFocus
            End If
        End If
        Exit Function
    End If
    
    If txtOnAirDate.Text = "" Then
        'If ((optTimeType(0).Value) Or (optTimeType(1).Value)) And (Not imIgnoreTimeTypeChg) Then
        If (Not imIgnoreTimeTypeChg) Then
            Beep
            gMsgBox "On Air Date must be specified." & Chr$(13) & Chr$(10) & "Please enter.", vbOKOnly
        '    optTimeType(0).Value = False
        '    optTimeType(1).Value = False

        End If
        Exit Function
    Else
        If gIsDate(txtOnAirDate.Text) = False Then
            If Not imIgnoreTimeTypeChg Then
                Beep
                gMsgBox "On Air Date must be specified in the form mm/dd/yy.", vbOKOnly
            End If
            Exit Function
        End If
    End If
    mLoadPledgeOk = True
End Function

Private Sub mETEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llCol As Long
    
    
    If (grdET.Row >= grdET.FixedRows) And (grdET.Row < grdET.Rows) And (grdET.Col >= ETDAYINDEX) And (grdET.Col < grdET.Cols - 1) Then
        lmETEnableRow = grdET.Row
        lmETEnableCol = grdET.Col
        Select Case grdET.Col
            Case ETDAYINDEX
                'If grdET.Row - grdET.TopRow + grdET.FixedRows <= 4 Then
                '    cbcETDay.PopUpListDirection "B"
                    cbcETDay.Move frcET.Left + grdET.Left + imETColPos(grdET.Col) + 15, frcET.Top + grdET.Top + grdET.RowPos(grdET.Row) + 15, (3 * grdET.ColWidth(grdET.Col) / 2), grdET.RowHeight(grdET.Row) - 15
                'Else
                '    cbcETDay.PopUpListDirection "A"
                '    cbcETDay.Move frcET.Left + grdET.Left + imETColPos(grdET.Col) + 15, frcET.Top + grdET.Top + grdET.RowPos(grdET.Row) - grdET.Height + 15, (3 * grdET.ColWidth(grdET.Col) / 2), grdET.RowHeight(grdET.Row) - 15
                'End If
                If grdET.TextMatrix(lmETEnableRow, lmETEnableCol) = "" Then
                    '9/2/16: Add default estimate day if site set to allow
                    ''12/8/14: Don't default the day
                    ''If lmETEnableRow = grdET.FixedRows Then
                    ''    cbcETDay.Text = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX)
                    ''Else
                    ''    If grdET.TextMatrix(lmETEnableRow - 1, ETFDDAYINDEX) = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX) Then
                    ''        cbcETDay.Text = grdET.TextMatrix(lmETEnableRow - 1, lmETEnableCol)
                    ''    Else
                    ''        cbcETDay.Text = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX)
                    ''    End If
                    ''End If
                    'cbcETDay.Text = ""
                    If bmDefaultEstDay Then
                        If lmETEnableRow = grdET.FixedRows Then
                            cbcETDay.Text = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX)
                        Else
                            If grdET.TextMatrix(lmETEnableRow - 1, ETFDDAYINDEX) = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX) Then
                                cbcETDay.Text = grdET.TextMatrix(lmETEnableRow - 1, lmETEnableCol)
                            Else
                                cbcETDay.Text = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX)
                            End If
                        End If
                    Else
                        cbcETDay.Text = ""
                    End If
                Else
                    cbcETDay.Text = grdET.TextMatrix(lmETEnableRow, lmETEnableCol)
                End If
                grdET.TextMatrix(lmETEnableRow, lmETEnableCol) = cbcETDay.Text
                cbcETDay.SetDropDownWidth cbcETDay.Width
                cbcETDay.ZOrder
                cbcETDay.Visible = True
                cbcETDay.SetFocus
            Case ETTIMEINDEX
                txtET.Move grdET.Left + imETColPos(grdET.Col) + 15, grdET.Top + grdET.RowPos(grdET.Row) + 15, grdET.ColWidth(grdET.Col) - 30, grdET.RowHeight(grdET.Row) - 15
                If grdET.Text <> "Missing" Then
                    txtET.Text = grdET.Text
                Else
                    txtET.Text = ""
                End If
                If txtET.Height > grdET.RowHeight(grdET.Row) - 15 Then
                    txtET.FontName = "Arial"
                    txtET.Height = grdET.RowHeight(grdET.Row) - 15
                End If
                txtET.Visible = True
                If txtET.Enabled Then
                    txtET.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub mETSetShow()
    Dim ilAvailIndex As Integer
    
    If (lmETEnableRow >= grdET.FixedRows) And (lmETEnableRow < grdET.Rows) Then
        If Trim$(grdET.TextMatrix(lmETEnableRow, ETAVAILINFOINDEX)) <> "" Then
            ilAvailIndex = Val(grdET.TextMatrix(lmETEnableRow, ETAVAILINFOINDEX))
            Select Case lmETEnableCol
                Case ETDAYINDEX
                    'Remove time if day removed
                    If grdET.TextMatrix(lmETEnableRow, ETDAYINDEX) <> "" Then
                        tmETAvailInfo(ilAvailIndex).sFdDay = grdET.TextMatrix(lmETEnableRow, ETFDDAYINDEX)
                        tmETAvailInfo(ilAvailIndex).sETDay = grdET.TextMatrix(lmETEnableRow, ETDAYINDEX)
                    Else
                        tmETAvailInfo(ilAvailIndex).sETDay = ""
                        grdET.TextMatrix(lmETEnableRow, ETTIMEINDEX) = ""
                        tmETAvailInfo(ilAvailIndex).sETTime = grdET.TextMatrix(lmETEnableRow, ETTIMEINDEX)
                    End If
                Case ETTIMEINDEX
                    tmETAvailInfo(ilAvailIndex).sFdTime = grdET.TextMatrix(lmETEnableRow, ETFDTIMEINDEX)
                    tmETAvailInfo(ilAvailIndex).sETTime = grdET.TextMatrix(lmETEnableRow, ETTIMEINDEX)
            End Select
        End If
    End If
    lmETEnableRow = -1
    lmETEnableCol = -1
    cbcETDay.Visible = False
    txtET.Visible = False
End Sub


Private Sub mPopEstTimes()
    Dim ilDat As Integer
    Dim ilDay As Integer
    Dim llDPSTime As Long
    Dim llDPETime As Long
    Dim llTime As Long
    Dim ilNext As Integer
    Dim ilUpper As Integer
    Dim sSDate As String
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    
    If (grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = "") Then
        Exit Sub
    End If
    If (grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = "-1") Then
        'ReDim tlSvDat(0 To UBound(tgDat)) As DAT
        'For ilDat = 0 To UBound(tgDat) - 1 Step 1
        '    tlSvDat(ilDat) = tgDat(ilDat)
        'Next ilDat
        'sSDate = Format$(txtOnAirDate.Text, sgShowDateForm)
        'ReDim tgDat(0 To 0) As DAT
        'gGetAvails lmAttCode, imShttCode, imVefCode, imVefCombo, sSDate, True
        'ReDim tmAvailDat(0 To UBound(tgDat)) As DAT
        'For ilDat = 0 To UBound(tgDat) - 1 Step 1
        '    tmAvailDat(ilDat) = tgDat(ilDat)
        'Next ilDat
        'ReDim tgDat(0 To UBound(tlSvDat)) As DAT
        'For ilDat = 0 To UBound(tlSvDat) - 1 Step 1
        '    tgDat(ilDat) = tlSvDat(ilDat)
        'Next ilDat
        ReDim tmAvailDat(0 To 0) As DAT
        For ilDay = MONFDINDEX To SUNFDINDEX Step 1
            tmAvailDat(UBound(tmAvailDat)).sFdSTime = grdPledge.TextMatrix(lmEnableRow, STARTTIMEFDINDEX)
            tmAvailDat(UBound(tmAvailDat)).sFdETime = grdPledge.TextMatrix(lmEnableRow, ENDTIMEFDINDEX)
            If grdPledge.TextMatrix(lmEnableRow, ilDay) <> "" Then
                tmAvailDat(UBound(tmAvailDat)).iFdDay(ilDay - MONFDINDEX) = 1
            Else
                tmAvailDat(UBound(tmAvailDat)).iFdDay(ilDay - MONFDINDEX) = 0
            End If
            ReDim Preserve tmAvailDat(0 To UBound(tmAvailDat) + 1) As DAT
        Next ilDay
        llDPSTime = gTimeToLong(grdPledge.TextMatrix(lmEnableRow, STARTTIMEFDINDEX), False)
        llDPETime = gTimeToLong(grdPledge.TextMatrix(lmEnableRow, ENDTIMEFDINDEX), True)
        For ilDay = MONFDINDEX To SUNFDINDEX Step 1
            'If Trim$(grdPledge.TextMatrix(lmEnableRow, ilDay)) <> "" Then
                For ilDat = 0 To UBound(tmAvailDat) - 1 Step 1
                    If tmAvailDat(ilDat).iFdDay(ilDay - MONFDINDEX) > 0 Then
                        llTime = gTimeToLong(tmAvailDat(ilDat).sFdSTime, False)
                        If (llTime >= llDPSTime) And (llTime < llDPETime) Then
                            If (grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = "-1") Then
                                grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX) = UBound(tmETAvailInfo)
                            Else
                                ilNext = grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX)
                                Do While ilNext <> -1
                                    If tmETAvailInfo(ilNext).iNextET = -1 Then
                                        tmETAvailInfo(ilNext).iNextET = UBound(tmETAvailInfo)
                                        Exit Do
                                    End If
                                    ilNext = tmETAvailInfo(ilNext).iNextET
                                Loop
                            End If
                            ilUpper = UBound(tmETAvailInfo)
                            Select Case ilDay
                                Case MONFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Mo"
                                Case TUEFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Tu"
                                Case WEDFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "We"
                                Case THUFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Th"
                                Case FRIFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Fr"
                                Case SATFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Sa"
                                Case SUNFDINDEX
                                    tmETAvailInfo(ilUpper).sFdDay = "Su"
                            End Select
                            tmETAvailInfo(ilUpper).sFdTime = tmAvailDat(ilDat).sFdSTime
                            tmETAvailInfo(ilUpper).sETDay = ""
                            tmETAvailInfo(ilUpper).sETTime = ""
                            tmETAvailInfo(ilUpper).lEptCode = 0
                            tmETAvailInfo(ilUpper).iNextET = -1
                            ReDim Preserve tmETAvailInfo(0 To ilUpper + 1) As ETAVAILINFO
                        End If
                    End If
                Next ilDat
            'End If
        Next ilDay
    End If
    'gGrid_Clear grdET, False
    grdET.Rows = 3
    llRow = grdET.FixedRows
    grdET.Row = llRow
    grdET.TextMatrix(llRow, ETFDDAYINDEX) = ""
    grdET.TextMatrix(llRow, ETFDTIMEINDEX) = ""
    grdET.TextMatrix(llRow, ETDAYINDEX) = ""
    grdET.TextMatrix(llRow, ETTIMEINDEX) = ""
    grdET.TextMatrix(llRow, ETAVAILINFOINDEX) = ""
    ilNext = grdPledge.TextMatrix(lmEnableRow, ESTIMATEDFIRSTINDEX)
    Do While ilNext <> -1
        slStr = Trim$(tmETAvailInfo(ilNext).sFdDay)
        ilDay = Switch(slStr = "Mo", 0, slStr = "Tu", 1, slStr = "We", 2, slStr = "Th", 3, slStr = "Fr", 4, slStr = "Sa", 5, slStr = "Su", 6)
        ilDay = ilDay + MONFDINDEX
        If grdPledge.TextMatrix(lmEnableRow, ilDay) <> "" Then
            If llRow + 1 > grdET.Rows Then
                grdET.AddItem ""
            End If
            grdET.Row = llRow
            grdET.Col = ETFDDAYINDEX
            grdET.CellBackColor = LIGHTYELLOW
            grdET.Col = ETFDTIMEINDEX
            grdET.CellBackColor = LIGHTYELLOW
            grdET.TextMatrix(llRow, ETFDDAYINDEX) = Trim$(tmETAvailInfo(ilNext).sFdDay)
            grdET.TextMatrix(llRow, ETFDTIMEINDEX) = Trim$(tmETAvailInfo(ilNext).sFdTime)
            grdET.TextMatrix(llRow, ETDAYINDEX) = Trim$(tmETAvailInfo(ilNext).sETDay)
            grdET.TextMatrix(llRow, ETTIMEINDEX) = Trim$(tmETAvailInfo(ilNext).sETTime)
            grdET.TextMatrix(llRow, ETAVAILINFOINDEX) = ilNext
            llRow = llRow + 1
        Else
            tmETAvailInfo(ilNext).sETDay = ""
            tmETAvailInfo(ilNext).sETTime = ""
        End If
        ilNext = tmETAvailInfo(ilNext).iNextET
    Loop
    grdET.Height = grdET.RowHeight(0) * (llRow) + 15
    gGrid_IntegralHeight grdET
    grdET.Height = grdET.Height + 30
    frcET.Height = grdET.Height
End Sub

Private Function mETColOk() As Integer
    mETColOk = True
    If grdET.CellBackColor = LIGHTYELLOW Then
        mETColOk = False
        Exit Function
    End If
End Function

Private Function mSaveEstimatedInfo(llAttCode As Long, ilShttCode As Integer, ilVefCode As Integer) As Integer
    Dim ilDat As Integer
    Dim ilNextET As Integer
    Dim llCode As Long
    Dim ilSeqNo As Integer
    Dim ilETDefined As Integer
    
    On Error GoTo ErrHand
    mSaveEstimatedInfo = True
    If (llAttCode <= 0) Or (ilShttCode <= 0) Or (ilVefCode <= 0) Then
        Exit Function
    End If
    ilETDefined = False
    For ilDat = 0 To UBound(tgDat) - 1 Step 1
        If tgStatusTypes(tgDat(ilDat).iFdStatus).iPledged = 1 Then
            ilNextET = tgDat(ilDat).iFirstET
            Do While ilNextET <> -1
                If (Trim$(tmETAvailInfo(ilNextET).sETDay) <> "") And (Trim$(tmETAvailInfo(ilNextET).sETTime) <> "") Then
                    ilETDefined = True
                    Exit For
                End If
                ilNextET = tmETAvailInfo(ilNextET).iNextET
            Loop
        End If
    Next ilDat
    If Not ilETDefined Then
        Exit Function
    End If
    '8/12/16: Separated Pledge and Estimate
    If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
        If bmPledgeDataChgd Then
            If bmETDataChgd Then
                gMsgBox "Estimates may not be changed.  Spots have been posted." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "The Last Week that Spots were Posted was on " & smLastPostedDate
            End If
            Exit Function
        End If
        If Not bmETDataChgd Then
            Exit Function
        End If
    End If
    For ilDat = 0 To UBound(tgDat) - 1 Step 1
        If tgStatusTypes(tgDat(ilDat).iFdStatus).iPledged = 1 Then
            ilNextET = tgDat(ilDat).iFirstET
            ilSeqNo = 1
            Do While ilNextET <> -1
                'If tmETAvailInfo(ilNextET).lEptCode <= 0 Then
                '9/4/14: save only if day and time defined
                If (Trim$(tmETAvailInfo(ilNextET).sETDay) <> "") And (Trim$(tmETAvailInfo(ilNextET).sETTime) <> "") Then
                    SQLQuery = "Insert Into ept ( "
                    SQLQuery = SQLQuery & "eptCode, "
                    SQLQuery = SQLQuery & "eptDatCode, "
                    SQLQuery = SQLQuery & "eptSeqNo, "
                    SQLQuery = SQLQuery & "eptAttCode, "
                    SQLQuery = SQLQuery & "eptShttCode, "
                    SQLQuery = SQLQuery & "eptVefCode, "
                    SQLQuery = SQLQuery & "eptFdAvailDay, "
                    SQLQuery = SQLQuery & "eptFdAvailTime, "
                    SQLQuery = SQLQuery & "eptEstimatedDay, "
                    SQLQuery = SQLQuery & "eptEstimatedTime, "
                    SQLQuery = SQLQuery & "eptUnused "
                    SQLQuery = SQLQuery & ") "
                    SQLQuery = SQLQuery & "Values ( "
                    SQLQuery = SQLQuery & "Replace" & ", "
                    SQLQuery = SQLQuery & tgDat(ilDat).lCode & ", "
                    SQLQuery = SQLQuery & ilSeqNo & ", "
                    SQLQuery = SQLQuery & llAttCode & ", "
                    SQLQuery = SQLQuery & ilShttCode & ", "
                    SQLQuery = SQLQuery & ilVefCode & ", "
                    SQLQuery = SQLQuery & "'" & gFixQuote(tmETAvailInfo(ilNextET).sFdDay) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(tmETAvailInfo(ilNextET).sFdTime, sgSQLTimeForm) & "', "
                    If Trim$(tmETAvailInfo(ilNextET).sETTime) <> "" Then
                        SQLQuery = SQLQuery & "'" & gFixQuote(tmETAvailInfo(ilNextET).sETDay) & "', "
                        SQLQuery = SQLQuery & "'" & Format$(tmETAvailInfo(ilNextET).sETTime, sgSQLTimeForm) & "', "
                    Else
                        SQLQuery = SQLQuery & "'" & "" & "', "
                        SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
                    End If
                    SQLQuery = SQLQuery & "'" & "" & "' "
                    SQLQuery = SQLQuery & ") "
                    llCode = gInsertAndReturnCode(SQLQuery, "ept", "eptCode", "Replace")
                    If llCode > 0 Then
                        tmETAvailInfo(ilNextET).lEptCode = llCode
                    Else
                        mSaveEstimatedInfo = False
                        Exit Function
                    End If
                'Else
    
                '    SQLQuery = "Update ept Set "
                '    SQLQuery = SQLQuery & "eptDatCode = " & tgDat(ilDat).lCode & ", "
                '    SQLQuery = SQLQuery & "eptSeqNo = " & ilSeqNo & ", "
                '    SQLQuery = SQLQuery & "eptAttCode = " & lmAttCode & ", "
                '    SQLQuery = SQLQuery & "eptShttCode = " & imShttCode & ", "
                '    SQLQuery = SQLQuery & "eptVefCode = " & imVefCode & ", "
                '    SQLQuery = SQLQuery & "eptFdAvailDay = '" & gFixQuote(tmETAvailInfo(ilNextET).sFdDay) & "', "
                '    SQLQuery = SQLQuery & "eptFdAvailTime = '" & Format$(tmETAvailInfo(ilNextET).sFdTime, sgSQLTimeForm) & "', "
                '    If Trim$(tmETAvailInfo(ilNextET).sETTime) <> "" Then
                '        SQLQuery = SQLQuery & "eptEstimatedDay = '" & gFixQuote(tmETAvailInfo(ilNextET).sETDay) & "', "
                '        SQLQuery = SQLQuery & "eptEstimatedTime = '" & Format$(tmETAvailInfo(ilNextET).sETTime, sgSQLTimeForm) & "', "
                '    Else
                '        SQLQuery = SQLQuery & "eptEstimatedDay = '" & "" & "', "
                '        SQLQuery = SQLQuery & "eptEstimatedTime = '" & Format$("12AM", sgSQLTimeForm) & "', "
                '    End If
                '    SQLQuery = SQLQuery & "eptUnused = '" & "" & "' "
                '    SQLQuery = SQLQuery & " WHERE eptCode = " & tmETAvailInfo(ilNextET).lEptCode
                '    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '        GoSub ErrHand:
                '    End If
                'End If
                End If
                ilNextET = tmETAvailInfo(ilNextET).iNextET
            Loop
        End If
    Next ilDat
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mSaveEstimatedInfo"
    mSaveEstimatedInfo = False
End Function

Private Function mGetET(tlDat As DAT) As Integer
    Dim ilDat As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilUpper As Integer
    Dim ilIndex As Integer
    Dim ilNext As Integer
    
    On Error GoTo ErrHand
    tlDat.iFirstET = -1
    If tlDat.sEstimatedTime = "Y" Then
        For ilDay = 0 To 6 Step 1
            If tlDat.iFdDay(ilDay) = 1 Then
                ilUpper = UBound(tmETAvailInfo)
                If tlDat.iFirstET = -1 Then
                    tlDat.iFirstET = ilUpper
                Else
                    tmETAvailInfo(ilUpper - 1).iNextET = ilUpper
                End If
                Select Case ilDay
                    Case MONFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Mo"
                    Case TUEFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Tu"
                    Case WEDFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "We"
                    Case THUFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Th"
                    Case FRIFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Fr"
                    Case SATFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Sa"
                    Case SUNFDINDEX
                        tmETAvailInfo(ilUpper).sFdDay = "Su"
                End Select
                tmETAvailInfo(ilUpper).sFdTime = tlDat.sFdSTime
                tmETAvailInfo(ilUpper).sETDay = ""
                tmETAvailInfo(ilUpper).sETTime = ""
                tmETAvailInfo(ilUpper).lEptCode = 0
                tmETAvailInfo(ilUpper).iNextET = -1
                ReDim Preserve tmETAvailInfo(0 To ilUpper + 1) As ETAVAILINFO
            End If
        Next ilDay
        SQLQuery = "SELECT * FROM ept"
        SQLQuery = SQLQuery + " WHERE (eptDatCode = " & tlDat.lCode & ")"
        SQLQuery = SQLQuery & " ORDER BY eptDatCode, eptSeqNo"
        Set rst_ept = gSQLSelectCall(SQLQuery)
        Do While Not rst_ept.EOF
            'If tlDat.iFirstET = -1 Then
            '    tlDat.iFirstET = UBound(tmETAvailInfo)
            'Else
            '    tmETAvailInfo(UBound(tmETAvailInfo) - 1).iNextET = UBound(tmETAvailInfo)
            'End If
            'tmETAvailInfo(UBound(tmETAvailInfo)).sFdDay = rst_ept!eptFdAvailDay
            ilIndex = -1
            ilNext = tlDat.iFirstET
            Do While ilNext <> -1
                If tmETAvailInfo(ilNext).sFdDay = rst_ept!eptFdAvailDay Then
                    ilIndex = ilNext
                    Exit Do
                End If
                ilNext = tmETAvailInfo(ilNext).iNextET
            Loop
            If ilIndex <> -1 Then
                slStr = gConvertTime(rst_ept!eptFdAvailTime)
                If Second(slStr) = 0 Then
                    slStr = Format$(slStr, sgShowTimeWOSecForm)
                Else
                    slStr = Format$(slStr, sgShowTimeWSecForm)
                End If
                tmETAvailInfo(ilIndex).sFdTime = slStr
                If Trim$(rst_ept!eptEstimatedDay) <> "" Then
                    tmETAvailInfo(ilIndex).sETDay = rst_ept!eptEstimatedDay
                    slStr = gConvertTime(rst_ept!eptEstimatedTime)
                    If Second(slStr) = 0 Then
                        slStr = Format$(slStr, sgShowTimeWOSecForm)
                    Else
                        slStr = Format$(slStr, sgShowTimeWSecForm)
                    End If
                    tmETAvailInfo(ilIndex).sETTime = slStr
                Else
                    tmETAvailInfo(ilIndex).sETDay = ""
                    '9/4/14: Handle case wher Day as missing but time defined
                    If Trim$(rst_ept!eptEstimatedTime) <> "" Then
                        slStr = gConvertTime(rst_ept!eptEstimatedTime)
                        If Second(slStr) = 0 Then
                            slStr = Format$(slStr, sgShowTimeWOSecForm)
                        Else
                            slStr = Format$(slStr, sgShowTimeWSecForm)
                        End If
                        tmETAvailInfo(ilIndex).sETTime = slStr
                    Else
                        tmETAvailInfo(ilIndex).sETTime = ""
                    End If
                End If
                tmETAvailInfo(ilIndex).lEptCode = rst_ept!eptCode
            End If
            rst_ept.MoveNext
        Loop
    End If
    mGetET = True
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mGetET"
    mGetET = False
End Function
Private Sub mPledgeSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim ilFirstDay As Integer
    Dim ilDay As Integer
    
    grdPledge.Redraw = False
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            ilFirstDay = 0
            If (ilCol = STARTTIMEFDINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                For ilDay = MONFDINDEX To SUNFDINDEX Step 1
                    If Trim$(grdPledge.TextMatrix(llRow, ilDay)) <> "" Then
                        ilFirstDay = ilDay - MONFDINDEX
                        Exit For
                    End If
                Next ilDay
            ElseIf (ilCol = AIRPLAYINDEX) Then
                slSort = Trim$(grdPledge.TextMatrix(llRow, AIRPLAYINDEX))
                Do While Len(slSort) < 2
                    slSort = "0" & slSort
                Loop
                For ilDay = MONPDINDEX To SUNPDINDEX Step 1
                    If Trim$(grdPledge.TextMatrix(llRow, ilDay)) <> "" Then
                        ilFirstDay = ilDay - MONPDINDEX
                        Exit For
                    End If
                Next ilDay
            ElseIf (ilCol = STARTTIMEPDINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                For ilDay = MONPDINDEX To SUNPDINDEX Step 1
                    If Trim$(grdPledge.TextMatrix(llRow, ilDay)) <> "" Then
                        ilFirstDay = ilDay - MONPDINDEX
                        Exit For
                    End If
                Next ilDay
            Else
                slSort = UCase$(Trim$(grdPledge.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slSort = slSort + Trim$(Str$(ilFirstDay))
            slStr = grdPledge.TextMatrix(llRow, SORTINDEX)
            'ilPos = InStr(1, slStr, "|", vbTextCompare)
            'If ilPos > 1 Then
            '    slStr = Left$(slStr, ilPos - 1)
            'End If
            If (ilCol <> imLastPledgeColSorted) Or ((ilCol = imLastPledgeColSorted) And (imLastPledgeSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                grdPledge.TextMatrix(llRow, SORTINDEX) = Left$(slSort & slStr, 27) '& "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                grdPledge.TextMatrix(llRow, SORTINDEX) = Left$(slSort & slStr, 27) '& "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastPledgeColSorted Then
        imLastPledgeColSorted = SORTINDEX
    Else
        imLastPledgeColSorted = -1
    End If
    gGrid_SortByCol grdPledge, STATUSINDEX, SORTINDEX, imLastPledgeColSorted, imLastPledgeSort
    imLastPledgeColSorted = ilCol
    grdPledge.Redraw = True
End Sub

Private Function mCheckEstTimes() As Boolean
    Dim llRow As Long
    Dim slStr As String
    Dim ilNextET As Integer
    Dim llPdSTime As Long
    Dim llPdETime As Long
    Dim llETTime As Long
    Dim llRowIndex As Long
    Dim ilIndex As Integer
    
    mCheckEstTimes = True
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                ilIndex = lbcStatus.ItemData(llRowIndex)
                If tmStatusTypes(ilIndex).iPledged = 1 Then  'Delayed
                    If Trim$(grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX)) <> "" Then
                        llPdSTime = gTimeToLong(Trim$(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX)), False)
                        llPdETime = gTimeToLong(Trim$(grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX)), True)
                        
                        ilNextET = Val(grdPledge.TextMatrix(llRow, ESTIMATEDFIRSTINDEX))
                        Do While ilNextET <> -1
                            If (Trim$(tmETAvailInfo(ilNextET).sETDay) <> "") And (Trim$(tmETAvailInfo(ilNextET).sETTime) <> "") Then
                                llETTime = gTimeToLong(tmETAvailInfo(ilNextET).sETTime, False)
                                'If Trim$(tmETAvailInfo(ilNextET).sETTime) <> "" Then
                                    llETTime = gTimeToLong(tmETAvailInfo(ilNextET).sETTime, False)
                                'Else
                                '    llETTime = 0
                                'End If
                                If (llETTime < llPdSTime) Or (llETTime >= llPdETime) Then
                                    grdPledge.Row = llRow
                                    grdPledge.Col = ESTIMATETIMEINDEX
                                    grdPledge.CellForeColor = vbRed
                                    mCheckEstTimes = False
                                End If
                            End If
                            ilNextET = tmETAvailInfo(ilNextET).iNextET
                        Loop
                    End If
                End If
            End If
        End If
    Next llRow
    Exit Function
End Function

Private Function mCheckHistDate() As Integer
    Dim blUpdateDate As Boolean
    Dim slHistDate As String
    Dim slSQLQuery As String
    On Error GoTo ErrHand
    
    blUpdateDate = False
    If txtHistorialDate.Text <> "" Then
        If txtOnAirDate.Text <> "" Then
            If gDateValue(txtOnAirDate.Text) < gDateValue(txtHistorialDate.Text) Then
                blUpdateDate = True
            End If
        End If
    Else
        If txtOnAirDate.Text <> "" Then
            blUpdateDate = True
        End If
    End If
    If blUpdateDate Then
        slHistDate = txtOnAirDate.Text
        slSQLQuery = "UPDATE shtt SET shttHistStartDate = '" & Format(slHistDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & imShttCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mCheckHistDate"
            mCheckHistDate = False
            Exit Function
        End If
        '11/26/16
        mUpdateShttTables imShttCode, False, False, slHistDate
    End If
    mCheckHistDate = True
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mGetET"
    mCheckHistDate = False
End Function

Private Function mCheckAirPlays() As Integer
    Dim ilNoAirPlays As Integer
    Dim ilAirPlay As Integer
    Dim blAnyDefined As Boolean
    Dim blAirPlayFound As Boolean
    Dim llRowIndex As Long
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim blAllNotCarried As Boolean

    ilNoAirPlays = Val(edcNoAirPlays.Text)
    If ilNoAirPlays = 0 Then
        ilNoAirPlays = 1
    End If
    blAnyDefined = False
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            blAnyDefined = True
        End If
    Next llRow
    If (Not blAnyDefined) And (ilNoAirPlays = 1) Then
        mCheckAirPlays = 0
        Exit Function
    End If
    'Check if row missing air play number
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                ilIndex = lbcStatus.ItemData(llRowIndex)
                If (tmStatusTypes(ilIndex).iPledged <= 1) Then
                    If Trim$(grdPledge.TextMatrix(llRow, AIRPLAYINDEX)) = "" Then
                        mCheckAirPlays = 1
                        Exit Function
                    End If
                End If
            End If
        End If
    Next llRow
    'Test if air play number is to large
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                ilIndex = lbcStatus.ItemData(llRowIndex)
                If (tmStatusTypes(ilIndex).iPledged <= 1) Then
                    If Val(grdPledge.TextMatrix(llRow, AIRPLAYINDEX)) > ilNoAirPlays Then
                        mCheckAirPlays = 3
                        Exit Function
                    End If
                End If
            End If
        End If
    Next llRow
    For ilAirPlay = 1 To ilNoAirPlays Step 1
        blAirPlayFound = False
        For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
            slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
            If slStr <> "" Then
                'llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                'If llRowIndex >= 0 Then
                '    ilIndex = lbcStatus.ItemData(llRowIndex)
                '    If (tmStatusTypes(ilIndex).iPledged <= 1) Then
                        If ilAirPlay = Val(grdPledge.TextMatrix(llRow, AIRPLAYINDEX)) Then
                            blAirPlayFound = True
                            Exit For
                        End If
                '    End If
                'End If
            End If
        Next llRow
        If Not blAirPlayFound Then
            mCheckAirPlays = 2
            Exit Function
        End If
    Next ilAirPlay
    'Test if all pledges set to Not Carried
    blAllNotCarried = False
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                ilIndex = lbcStatus.ItemData(llRowIndex)
                If (tmStatusTypes(ilIndex).iPledged <= 1) Then
                    blAllNotCarried = True
                    Exit For
                End If
            End If
        End If
    Next llRow
    If Not blAllNotCarried Then
        mCheckAirPlays = 4
        Exit Function
    End If
    mCheckAirPlays = 0
End Function

Private Sub mShowTabs()

    If (imVefCode <= 0) Or (imShttCode <= 0) Then
        frcTab(0).Visible = False
        frcTab(1).Visible = False
        frcTab(2).Visible = False
        frcTab(3).Visible = False
        frcTab(4).Visible = False
        TabStrip1.Visible = False
        frcEvent.Visible = False
        grdMulticast.Visible = False
    Else
        TabStrip1.Visible = True
        If imTabIndex = TabStrip1.SelectedItem.Index Then
            Select Case TabStrip1.SelectedItem.Index
                Case TABMAIN  'Main
                    frcTab(0).Visible = True
                Case TABPERSONNEL  'Personnel
                    frcTab(4).Visible = True
                Case TABPLEDGE  'Pledge
                    If Not mLoadPledgeOk() Then
                        imTabIndex = -1
                        'TabStrip1.SetFocus
                        'gSendKeys "%M", True
                        TabStrip1.Tabs(TABMAIN).Selected = True
                        'txtOnAirDate.SetFocus
                        pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    If smPledgeByEvent <> "Y" Then
                        frcTab(2).Visible = True
                        frcEvent.Visible = False
                    Else
                        mLoadPledge True, -1
                        frcEvent.Visible = True
                        frcTab(2).Visible = False
                    End If
                Case TABDELIVERY  'Delivery
                    frcTab(1).Visible = True
                Case TABINTERFACE  'Interface
                    frcTab(3).Visible = True
            End Select
        'Else
        '    TabStrip1.SelectedItem.Index = imTabIndex
        End If
    End If

End Sub
Private Sub mMCPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdMulticast.Row = llRow
    For llCol = MCCALLLETTERSINDEX To MCDATERANGEINDEX Step 1
        grdMulticast.Col = llCol
        If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) <> "1" Then
            If grdMulticast.CellBackColor <> LIGHTYELLOW Then
                grdMulticast.CellBackColor = vbWhite
            Else
                grdMulticast.CellBackColor = LIGHTYELLOW
            End If
        Else
            grdMulticast.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub


Private Function mSavePledgeInfo(llAttCode As Long, ilShttCode As Integer, ilVefCode As Integer) As Integer
    Dim iLoop As Integer
    Dim sFdStTime As String
    Dim sFdEdTime As String
    Dim sPdStTime As String
    Dim sPdEdTime As String
    Dim llCode As Long
    
    On Error GoTo ErrHand
    mSavePledgeInfo = True
    If (llAttCode <= 0) Or (ilShttCode <= 0) Or (ilVefCode <= 0) Then
        Exit Function
    End If
    '2/18/16: Pledge not allowed to be changed
    If Not imOkToChange And smLastPostedDate <> "1/1/1970" Then
        '8/12/16: Separated Pledge and Estimate
        'Exit Function
        If bmPledgeDataChgd Then
            Exit Function
        End If
        If Not bmETDataChgd Then
            Exit Function
        End If
    End If
    If (imDatLoaded) And (sgUstPledge = "Y") And (smPledgeByEvent <> "Y") Then
        'Delete Dayparts or Avails
        SQLQuery = "DELETE FROM dat"
        SQLQuery = SQLQuery + " WHERE (datAtfCode = " & llAttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSavePledgeInfo"
            mSavePledgeInfo = False
            Exit Function
        End If
        'Since DAT are removed, remove the matching ept
        SQLQuery = "DELETE FROM ept"
        SQLQuery = SQLQuery + " WHERE (eptAttCode = " & llAttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSavePledgeInfo"
            mSavePledgeInfo = False
            Exit Function
        End If
        For iLoop = 0 To UBound(tgDat) - 1 Step 1
            sFdStTime = Format$(tgDat(iLoop).sFdSTime, sgShowTimeWSecForm)
            sFdEdTime = Format$(tgDat(iLoop).sFdETime, sgShowTimeWSecForm)
            If Len(Trim$(tgDat(iLoop).sPdSTime)) = 0 Or Asc(tgDat(iLoop).sPdSTime) = 0 Then
                sPdStTime = sFdStTime
            Else
                sPdStTime = Format$(tgDat(iLoop).sPdSTime, sgShowTimeWSecForm)
            End If
            If Len(Trim$(tgDat(iLoop).sPdETime)) = 0 Or Asc(tgDat(iLoop).sPdETime) = 0 Then
                sPdEdTime = sPdStTime
            Else
                sPdEdTime = Format$(tgDat(iLoop).sPdETime, sgShowTimeWSecForm)
            End If
            'To avoid duplicate key when two saves are done in a row.
            'always set lCode to zero (0).  This should not be required because of the 'Delete From'
            'If IsDirty = False Then
                tgDat(iLoop).lCode = 0
            'End If
            'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
            SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
            SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
            SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
            SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
            SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime, datEmbeddedOrROS)"
            'SQLQuery = SQLQuery & " VALUES (" & tgDat(iLoop).lCode & ", " & lmAttCode & ", " & imShttCode & ", " & imVefCode
            SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & llAttCode & ", " & ilShttCode & ", " & ilVefCode
            
            'If optTimeType(0).Value Or optTimeType(1).Value Or optTimeType(2).Value Then
            '    If optTimeType(0).Value Then       'Live Dayparts
            '        SQLQuery = SQLQuery & ",0,"
            '    ElseIf optTimeType(1).Value Then   'Avails
            '        SQLQuery = SQLQuery & ",1,"
            '    ElseIf optTimeType(2).Value Then   'CD/Tape Dayparts
            '        SQLQuery = SQLQuery & ",2,"
            '    End If
            'Else
            '    SQLQuery = SQLQuery & "," & tgDat(0).iDACode & ","
            'End If
            SQLQuery = SQLQuery & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(0) & ", " & tgDat(iLoop).iFdDay(1) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(2) & ", " & tgDat(iLoop).iFdDay(3) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(4) & ", " & tgDat(iLoop).iFdDay(5) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iFdDay(6) & ", "
            SQLQuery = SQLQuery & "'" & Format$(sFdStTime, sgSQLTimeForm) & "','" & Format$(sFdEdTime, sgSQLTimeForm) & "',"
            SQLQuery = SQLQuery & tgDat(iLoop).iFdStatus & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(0) & ", " & tgDat(iLoop).iPdDay(1) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(2) & ", " & tgDat(iLoop).iPdDay(3) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(4) & ", " & tgDat(iLoop).iPdDay(5) & ","
            SQLQuery = SQLQuery & tgDat(iLoop).iPdDay(6) & ", "
            
            If Asc(tgDat(iLoop).sPdDayFed) = 0 Then
                SQLQuery = SQLQuery & "'" & " " & "', "
            Else
                SQLQuery = SQLQuery & "'" & tgDat(iLoop).sPdDayFed & "', "
            End If
            
            
            
            SQLQuery = SQLQuery & "'" & Format$(sPdStTime, sgSQLTimeForm) & "','" & Format$(sPdEdTime, sgSQLTimeForm) & "',"
            SQLQuery = SQLQuery & tgDat(iLoop).iAirPlayNo & ", "
            SQLQuery = SQLQuery & "'" & tgDat(iLoop).sEstimatedTime & "', '" & tgDat(iLoop).sEmbeddedOrROS & "')"
            ''cnn.Execute SQLQuery, rdExecDirect
            'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '    GoSub ErrHand:
            'End If
            llCode = gInsertAndReturnCode(SQLQuery, "dat", "datCode", "Replace")
            If llCode > 0 Then
                tgDat(iLoop).lCode = llCode
            Else
                mSavePledgeInfo = False
                Exit Function
            End If
            '5/22/07:  No Required as DAT is deleted (never updated)
            'SQLQuery = "SELECT MAX(datCode) from dat"
            'Set rst = gSQLSelectCall(SQLQuery)
            'If Not rst.EOF Then
            '    tgDat(iLoop).lCode = rst(0).Value
            'End If
        Next iLoop
    End If
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mSavePledgeInfo"
    mSavePledgeInfo = False
End Function

Private Function mSplitAgreement(blAddAgreementOnly As Boolean, llAttCode As Long, slTerminateDate As String, slOffAir As String, CurTime As String, ilNoAirPlays As Integer, llNewATTCode As Long) As Boolean
    Dim slOnAir As String
    Dim llMCAttCode As Long
    Dim ilMCShttCode As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mSplitAgreement = True
    SQLQuery = "SELECT * FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & llAttCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF = True Then
        Exit Function
    End If
    If Not blAddAgreementOnly Then
        SQLQuery = "UPDATE att SET "
        SQLQuery = SQLQuery & "attOffAir = '" & Format$(slTerminateDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
        SQLQuery = SQLQuery & " WHERE attCode = " & llAttCode
        'cnn.BeginTrans
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSplitAgreement"
            mSplitAgreement = False
            Exit Function
        End If
    End If
    slOnAir = gObtainNextMonday(slTerminateDate)
    ilMCShttCode = rst!attshfCode
    SQLQuery = "Insert Into att ( "
    SQLQuery = SQLQuery & "attCode, "
    SQLQuery = SQLQuery & "attShfCode, "
    SQLQuery = SQLQuery & "attVefCode, "
    SQLQuery = SQLQuery & "attAgreeStart, "
    SQLQuery = SQLQuery & "attAgreeEnd, "
    SQLQuery = SQLQuery & "attOnAir, "
    SQLQuery = SQLQuery & "attOffAir, "
    SQLQuery = SQLQuery & "attSigned, "
    SQLQuery = SQLQuery & "attSignDate, "
    SQLQuery = SQLQuery & "attLoad, "
    SQLQuery = SQLQuery & "attTimeType, "
    SQLQuery = SQLQuery & "attComp, "
    SQLQuery = SQLQuery & "attStartTime, "
    SQLQuery = SQLQuery & "attBarCode, "
    SQLQuery = SQLQuery & "attDropDate, "
    SQLQuery = SQLQuery & "attCPType, "
    SQLQuery = SQLQuery & "attUsfCode, "
    SQLQuery = SQLQuery & "attEnterDate, "
    SQLQuery = SQLQuery & "attEnterTime, "
    SQLQuery = SQLQuery & "attNotice, "
    SQLQuery = SQLQuery & "attCarryCmml, "
    SQLQuery = SQLQuery & "attNoCDs, "
    SQLQuery = SQLQuery & "attSendTape, "
    SQLQuery = SQLQuery & "attACName, "
    SQLQuery = SQLQuery & "attACPhone, "
    SQLQuery = SQLQuery & "attGenLog, "
    SQLQuery = SQLQuery & "attGenCP, "
    SQLQuery = SQLQuery & "attPostingType, "
    SQLQuery = SQLQuery & "attPrintCP, "
    SQLQuery = SQLQuery & "attComments, "
    SQLQuery = SQLQuery & "attGenOther, "
    SQLQuery = SQLQuery & "attAgreementID, "
    SQLQuery = SQLQuery & "attWklyClear, "
    SQLQuery = SQLQuery & "attHrlyClear, "
    SQLQuery = SQLQuery & "attHrUsed, "
    SQLQuery = SQLQuery & "attExportType, "
    SQLQuery = SQLQuery & "attLogType, "
    SQLQuery = SQLQuery & "attPostType, "
    SQLQuery = SQLQuery & "attWebPW, "
    SQLQuery = SQLQuery & "attWebEmail, "
    SQLQuery = SQLQuery & "attSendLogEMail, "
    SQLQuery = SQLQuery & "attSuppressNotice, "
    SQLQuery = SQLQuery & "attLabelID, "
    SQLQuery = SQLQuery & "attLabelShipInfo, "
    SQLQuery = SQLQuery & "attMulticast, "
    SQLQuery = SQLQuery & "attRadarClearType, "
    SQLQuery = SQLQuery & "attArttCode, "
    SQLQuery = SQLQuery & "attStatus, "
    SQLQuery = SQLQuery & "attNCR, "
    SQLQuery = SQLQuery & "attFormerNCR, "
    SQLQuery = SQLQuery & "attForbidSplitLive, "
    SQLQuery = SQLQuery & "attXDReceiverID, "
    SQLQuery = SQLQuery & "attVoiceTracked, "
    SQLQuery = SQLQuery & "attMonthlyWebPost, "
    SQLQuery = SQLQuery & "attWebInterface, "
    SQLQuery = SQLQuery & "attContractPrinted, "
    SQLQuery = SQLQuery & "attMktRepUstCode, "
    SQLQuery = SQLQuery & "attServRepUstCode, "
    SQLQuery = SQLQuery & "attVehProgStartTime, "
    SQLQuery = SQLQuery & "attVehProgEndTime, "
    SQLQuery = SQLQuery & "attExportToWeb, "
    SQLQuery = SQLQuery & "attExportToUnivision, "
    SQLQuery = SQLQuery & "attExportToMarketron, "
    SQLQuery = SQLQuery & "attExportToCBS, "
    SQLQuery = SQLQuery & "attExportToClearCh, "
    SQLQuery = SQLQuery & "attContractPFTCode, "
    SQLQuery = SQLQuery & "attPledgeType, "
    SQLQuery = SQLQuery & "attNoAirPlays, "
    SQLQuery = SQLQuery & "attDesignVersion, "
    SQLQuery = SQLQuery & "attIDCReceiverID, "
    SQLQuery = SQLQuery & "attSentToXDSStatus, "
    SQLQuery = SQLQuery & "attAudioDelivery, "
    SQLQuery = SQLQuery & "attExportToJelli, "
    '3/23/15: Add Send Delays to XDS
    SQLQuery = SQLQuery & "attSendDelayToXDS, "
    SQLQuery = SQLQuery & "attServiceAgreement, "
    '4/3/19
    SQLQuery = SQLQuery & "attExcludeFillSpot, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeQ, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeR, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeT, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeM, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeS, "
    SQLQuery = SQLQuery & "attExcludeCntrTypeV, "
    
    SQLQuery = SQLQuery & "attUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & rst!attshfCode & ", "
    SQLQuery = SQLQuery & rst!attvefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(rst!attAgreeStart, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(rst!attAgreeEnd, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slOnAir, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slOffAir, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & rst!attSigned & ", "
    SQLQuery = SQLQuery & "'" & Format$(rst!attSignDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & rst!attLoad & ", "
    SQLQuery = SQLQuery & rst!attTimeType & ", "
    SQLQuery = SQLQuery & rst!attComp & ", "
    SQLQuery = SQLQuery & "'" & Format$(rst!attStartTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & rst!attBarCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(rst!attDropDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & rst!attCPType & ", "
    SQLQuery = SQLQuery & rst!attUsfCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(smCurDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(CurTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attNotice))) & "', "
    SQLQuery = SQLQuery & rst!attCarryCmml & ", "
    SQLQuery = SQLQuery & rst!attNoCDs & ", "
    SQLQuery = SQLQuery & rst!attSendTape & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attACName))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attACPhone))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attGenLog))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attGenCP))) & "', "
    SQLQuery = SQLQuery & rst!attPostingType & ", "
    SQLQuery = SQLQuery & rst!attPrintCP & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attComments))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attGenOther))) & "', "
    SQLQuery = SQLQuery & rst!attAgreementID & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attWklyClear))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attHrlyClear))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attHrUsed))) & "', "
    SQLQuery = SQLQuery & rst!attExportType & ", "
    SQLQuery = SQLQuery & rst!attLogType & ", "
    SQLQuery = SQLQuery & rst!attPostType & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attWebPW))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attWebEmail))) & "', "
    SQLQuery = SQLQuery & rst!attSendLogEmail & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attSuppressNotice))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attLabelID))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attLabelShipInfo))) & "', "
    SQLQuery = SQLQuery & "'" & "Y" & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attRadarClearType))) & "', "
    SQLQuery = SQLQuery & rst!attArttCode & ", "
    SQLQuery = SQLQuery & "'" & "C" & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attNCR))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attFormerNCR))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attForbidSplitLive))) & "', "
    SQLQuery = SQLQuery & rst!attXDReceiverId & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attVoiceTracked))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attMonthlyWebPost))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attWebInterface))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attContractPrinted))) & "', "
    SQLQuery = SQLQuery & rst!attMktRepUstCode & ", "
    SQLQuery = SQLQuery & rst!attServRepUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(rst!attVehProgStartTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(rst!attVehProgEndTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToWeb))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToUnivision))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToMarketron))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToCBS))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToClearCh))) & "', "
    SQLQuery = SQLQuery & rst!attContractPFTCode & ", "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attPledgeType))) & "', "
    SQLQuery = SQLQuery & ilNoAirPlays & ", "
    SQLQuery = SQLQuery & 2 & ", "
    SQLQuery = SQLQuery & "'" & rst!attIDCReceiverID & "', "
    SQLQuery = SQLQuery & "'" & "M" & "', "
    SQLQuery = SQLQuery & "'" & rst!attAudioDelivery & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExportToJelli))) & "', "
    '3/23/15: Add Send Delays to XDS
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attSendDelayToXDS))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attServiceAgreement))) & "', "
    '4/3/19
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeFillSpot))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeQ))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeR))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeT))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeM))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeS))) & "', "
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attExcludeCntrTypeV))) & "', "
    
    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(gStripIllegalChr(rst!attUnused))) & "' "
    SQLQuery = SQLQuery & ") "
    llMCAttCode = gInsertAndReturnCode(SQLQuery, "att", "attCode", "Replace")
    llNewATTCode = llMCAttCode
    If llMCAttCode <= 0 Then
        mSplitAgreement = False
        Exit Function
    '7701
    Else
        SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llMCAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & llAttCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSplitAgreement"
            mSplitAgreement = False
            Exit Function
        End If
    End If
    If Not mSavePledgeInfo(llMCAttCode, ilMCShttCode, imVefCode) Then
        mSplitAgreement = False
        Exit Function
    End If
    ilRet = mSetUsedForAtt(ilMCShttCode, False)
    'ilRet = mSaveContractPDF(llMCAttCode, ilMCShttCode, imVefCode)
    SQLQuery = "SELECT * FROM pft WHERE pftAttCode = " & llAttCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        SQLQuery = "Insert Into pft ( "
        SQLQuery = SQLQuery & "pftAttCode, "
        SQLQuery = SQLQuery & "pftShttCode, "
        SQLQuery = SQLQuery & "pftVefCode, "
        SQLQuery = SQLQuery & "pftPDFName, "
        SQLQuery = SQLQuery & "pftDateEntered, "
        SQLQuery = SQLQuery & "pftUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & llMCAttCode & ", "
        SQLQuery = SQLQuery & rst!pftShttCode & ", "
        SQLQuery = SQLQuery & rst!pftVefCode & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(rst!pftPDFName) & "', "
        SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            mMousePointer vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mSplitAgreement"
            mSplitAgreement = False
            Exit Function
        End If
    End If
    ilRet = mSaveEstimatedInfo(llMCAttCode, ilMCShttCode, imVefCode)
    ilRet = mSaveEventInfo(llMCAttCode, ilMCShttCode, imVefCode)
    
    '11/19/13: Create CPTT for multi-cast (ttp 5385)
    SQLQuery = "UPDATE cptt SET "
    SQLQuery = SQLQuery + "cpttAtfCode = " & llMCAttCode
    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & llAttCode
    SQLQuery = SQLQuery + " AND cpttStartDate >= " & "'" & Format$(slOnAir, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        mMousePointer vbDefault
        gHandleError "AffErrorLog.txt", "AffAgmnt-mSplitAgreement"
        mSplitAgreement = False
        Exit Function
    End If
    
    
    Exit Function

ErrHand:
    mMousePointer vbDefault
    Resume Next
    gHandleError "", "Agreement-mSplitAgreement"
    mSplitAgreement = False
End Function

Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdET, grdMulticast, ilMousepointer
    gSetMousePointer grdPledge, grdEvent, ilMousepointer
End Sub


Private Function mUpdateIDCExport(slStartDate As String, slEndDate As String) As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    mUpdateIDCExport = True
    If Trim$(smIDCReceiverID) = Trim$(txtIDCReceiverID.Text) Then
        Exit Function
    End If
    
    On Error GoTo ErrHand1
    If (Len(Trim$(smIDCReceiverID)) <> 0) And (Len(Trim$(txtIDCReceiverID.Text)) <> 0) Then
        If Val(smIDCReceiverID) = Val(txtIDCReceiverID.Text) Then
            Exit Function
        End If
    End If
    If gDateValue(slStartDate) > gDateValue(Format(gNow(), "m/d/yy")) + 28 Then
        Exit Function
    End If
    On Error GoTo ErrHand
'    '5882 no longer needed
'    'Reset all IDC Exports if copy Exported
'    SQLQuery = "SELECT Distinct crfRafCode From Crf_Copy_Rot_Header Where crfRafCode > 0 " & " and crfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "' and crfStartDate <= '" & Format(slEndDate, sgSQLDateForm) & "'"
'    Set rst_crf = gSQLSelectCall(SQLQuery)
'    Do While Not rst_crf.EOF
'        slSQLQuery = "Update ief_IDC_Enforced Set iefExportStatus = '" & "R" & "'"
'        slSQLQuery = slSQLQuery & " Where (iefExportStatus = 'E' And iefSplitRafCode = " & rst_crf!crfRafCode & ")"
'        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'        rst_crf.MoveNext
'    Loop
    Exit Function
ErrHand1:
    mUpdateIDCExport = False
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mUpdateIDCExport"
    mUpdateIDCExport = False
    Exit Function
End Function

Private Sub mPopulateEvents()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slDeclaration As String
    Dim llCellColor As Long
    Dim ilRet As Integer
    Dim llPet As Long
    Dim slDay As String
    Dim ilTimeAdj As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llDate As Long
    Dim llTime As Long
    Dim llSeasonGhfCode As Long
    Dim llLoop As Long
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String
    '12/12/14
    Dim slFed As String
    
    On Error GoTo ErrHand:
    If smPledgeByEvent <> "Y" Then
        Exit Sub
    End If
    If cbcSeason.ListIndex < 0 Then
        Exit Sub
    End If
    If (imShttCode <= 0) Or (imVefCode <= 0) Then
        Exit Sub
    End If
    mMousePointer vbHourglass
    mClearPledgeGrid
    llSeasonGhfCode = cbcSeason.GetItemData(cbcSeason.ListIndex)
    grdEvent.Redraw = False
    ilTimeAdj = gGetTimeAdj(imShttCode, imVefCode, slFed)
    If UBound(tmPetInfo) <= LBound(tmPetInfo) Then
        ilRet = gPopPet(lmAttCode, tmPetInfo())
    End If
    grdEvent.Row = 0
    For llCol = EVTEVENTNOINDEX To EVTUNDECIDEDINDEX Step 1
        grdEvent.Col = llCol
        grdEvent.CellBackColor = LIGHTBLUE
    Next llCol
    grdEvent.Row = 1
    For llCol = EVTEVENTNOINDEX To EVTUNDECIDEDINDEX Step 1
        grdEvent.Col = llCol
        grdEvent.CellBackColor = LIGHTBLUE
    Next llCol
    gGrid_Clear grdEvent, True
    For llRow = grdEvent.FixedRows To grdEvent.Rows - 1 Step 1
        grdEvent.Row = llRow
        For llCol = EVTEVENTNOINDEX To EVTUNDECIDEDINDEX Step 1
            grdEvent.Col = llCol
            grdEvent.CellBackColor = vbWhite
        Next llCol
    Next llRow
    slDeclaration = ""
    llRow = grdEvent.FixedRows
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & imVefCode & " AND gsfGhfCode = " & llSeasonGhfCode & ")"
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Gsf.EOF
        If llRow >= grdEvent.Rows Then
            grdEvent.AddItem ""
        End If
        llCellColor = vbWhite
        'slDeclaration = ""
        'If lmAttCode > 0 Then
        '    llPet = gBinarySearchPet(rst_Gsf!gsfCode)
        '    If llPet <> -1 Then
        '        slDeclaration = tmPetInfo(llPet).sDeclaredStatus
        '    End If
        'End If
        llPet = -1
        For llLoop = 0 To UBound(tmPetInfo) - 1 Step 1
            If tmPetInfo(llLoop).lgsfCode = rst_Gsf!gsfCode Then
                llPet = llLoop
                Exit For
            End If
        Next llLoop
        If llPet = -1 Then
            llPet = UBound(tmPetInfo)
            tmPetInfo(llPet).lCode = 0
            tmPetInfo(llPet).lGhfCode = rst_Gsf!gsfGhfCode
            tmPetInfo(llPet).lgsfCode = rst_Gsf!gsfCode
            If gDateValue(Format$(rst_Gsf!gsfAirDate, sgShowDateForm)) <= gDateValue(Format(gNow(), "m/d/yy")) Then
                tmPetInfo(llPet).sDeclaredStatus = "N"
            Else
                tmPetInfo(llPet).sDeclaredStatus = "U"
            End If
            tmPetInfo(llPet).sClearStatus = "U"
            tmPetInfo(llPet).sChanged = "Y"
            ReDim Preserve tmPetInfo(0 To llPet + 1) As PETINFO
        End If
        slDeclaration = tmPetInfo(llPet).sDeclaredStatus
        For llCol = EVTEVENTNOINDEX To EVTAIRTIMEINDEX Step 1
            grdEvent.Row = llRow
            grdEvent.Col = llCol
            grdEvent.CellBackColor = LIGHTYELLOW
        Next llCol
        For llCol = EVTCARRYINDEX To EVTUNDECIDEDINDEX Step 1
            grdEvent.Row = llRow
            grdEvent.Col = llCol
            grdEvent.CellBackColor = vbWhite
        Next llCol
        'Game Number
        grdEvent.Row = llRow
        grdEvent.Col = EVTEVENTNOINDEX
        grdEvent.CellAlignment = flexAlignRightCenter
        If rst_Gsf!gsfGameStatus = "F" Then
            grdEvent.CellForeColor = LIGHTGREEN
        ElseIf rst_Gsf!gsfGameStatus = "T" Then
            grdEvent.CellForeColor = ORANGE
        ElseIf rst_Gsf!gsfGameStatus = "P" Then
            grdEvent.CellForeColor = vbBlue
        ElseIf rst_Gsf!gsfGameStatus = "C" Then
            grdEvent.CellForeColor = vbRed
        End If
        grdEvent.TextMatrix(llRow, EVTEVENTNOINDEX) = rst_Gsf!gsfGameNo
        'Feed Source
        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
            gGetEventTitles imVefCode, slEventTitle1, slEventTitle2
            If rst_Gsf!gsfFeedSource = "V" Then
                grdEvent.TextMatrix(llRow, EVTFEEDSOURCEINDEX) = slEventTitle1  '"Visiting"
            ElseIf rst_Gsf!gsfFeedSource = "N" Then
                grdEvent.TextMatrix(llRow, EVTFEEDSOURCEINDEX) = "National"
            Else
                grdEvent.TextMatrix(llRow, EVTFEEDSOURCEINDEX) = slEventTitle2  '"Home"
            End If
        End If
        'Language
        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
            For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                If tgLangInfo(ilLang).iCode = rst_Gsf!gsfLangMnfCode Then
                    grdEvent.TextMatrix(llRow, EVTLANGUAGEINDEX) = Trim$(tgLangInfo(ilLang).sName)
                    Exit For
                End If
            Next ilLang
        End If
        'Visiting Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                grdEvent.TextMatrix(llRow, EVTVISITTEAMINDEX) = Trim$(tgTeamInfo(ilTeam).sName)
                Exit For
            End If
        Next ilTeam
        'Home Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                grdEvent.TextMatrix(llRow, EVTHOMETEAMINDEX) = Trim$(tgTeamInfo(ilTeam).sName)
                Exit For
            End If
        Next ilTeam
        'Air Date
        slDate = Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
        'Start Time
        slTime = Format$(rst_Gsf!gsfAirTime, sgShowTimeWSecForm)
        gAdjustEventTime ilTimeAdj, slDate, slTime
        Select Case Weekday(slDate)
            Case vbMonday
                slDay = "Mo, "
            Case vbTuesday
                slDay = "Tu, "
            Case vbWednesday
                slDay = "We, "
            Case vbThursday
                slDay = "Th, "
            Case vbFriday
                slDay = "Fr, "
            Case vbSaturday
                slDay = "Sa, "
            Case vbSunday
                slDay = "Su, "
        End Select

        grdEvent.TextMatrix(llRow, EVTAIRDATEINDEX) = slDay & slDate
        grdEvent.TextMatrix(llRow, EVTAIRTIMEINDEX) = slTime
        grdEvent.Row = llRow
        If slDeclaration = "Y" Then
            grdEvent.Col = EVTCARRYINDEX
            grdEvent.CellBackColor = LIGHTGREEN
        ElseIf slDeclaration = "N" Then
            grdEvent.Col = EVTNOTCARRIEDINDEX
            grdEvent.CellBackColor = vbRed
        Else
            grdEvent.Col = EVTUNDECIDEDINDEX
            grdEvent.CellBackColor = ORANGE
        End If
        grdEvent.TextMatrix(llRow, EVTPETINFOINDEX) = llPet
        llRow = llRow + 1
        rst_Gsf.MoveNext
    Loop
    rst_Gsf.Close
    mEventSortCol EVTAIRTIMEINDEX
    mEventSortCol EVTAIRDATEINDEX
    'If llRow = grdEvent.FixedRows + 1 Then
    '    grdEvent.Row = grdEvent.FixedRows
    '    grdEvent.RowSel = grdEvent.FixedRows
    '    grdEvent.Col = GAMENOINDEX
    '    grdEvent.ColSel = AIRTIMEINDEX
    '    lmRowSelected = grdEvent.FixedRows
    'Else
        grdEvent.Row = 0
        grdEvent.Col = EVTPETINFOINDEX
    'End If
    grdEvent.Redraw = True
    mMousePointer vbDefault
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mPopulateEvents"
End Sub

Private Sub mEventSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdEvent.FixedRows To grdEvent.Rows - 1 Step 1
        slStr = Trim$(grdEvent.TextMatrix(llRow, EVTEVENTNOINDEX))
        If slStr <> "" Then
            If ilCol = EVTAIRDATEINDEX Then
                slSort = Mid(Trim$(grdEvent.TextMatrix(llRow, EVTAIRDATEINDEX)), 5)
                slSort = gDateValue(slSort)
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = EVTAIRTIMEINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdEvent.TextMatrix(llRow, EVTAIRTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = EVTEVENTNOINDEX) Then
                slSort = Trim$(grdEvent.TextMatrix(llRow, EVTEVENTNOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdEvent.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = " "
                End If
            End If
            slStr = grdEvent.TextMatrix(llRow, EVTSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastEventColSorted) Or ((ilCol = imLastEventColSorted) And (imLastEventSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdEvent.TextMatrix(llRow, EVTSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdEvent.TextMatrix(llRow, EVTSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastEventColSorted Then
        imLastEventColSorted = EVTSORTINDEX
    Else
        imLastEventColSorted = -1
        imLastEventSort = -1
    End If
    gGrid_SortByCol grdEvent, EVTEVENTNOINDEX, EVTSORTINDEX, imLastEventColSorted, imLastEventSort
    imLastEventColSorted = ilCol
End Sub

Private Function mSaveEventInfo(llAttCode As Long, ilShttCode As Integer, ilVefCode As Integer) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llCode As Long
    Dim slDeclaration As String
    Dim llPet As Integer
    
    On Error GoTo ErrHand
    If smPledgeByEvent <> "Y" Then
        mSaveEventInfo = True
        Exit Function
    End If
    If (llAttCode <= 0) Or (ilShttCode <= 0) Or (ilVefCode <= 0) Then
        mSaveEventInfo = True
        Exit Function
    End If
    For llPet = 0 To UBound(tmPetInfo) - 1 Step 1
    
        slDeclaration = tmPetInfo(llPet).sDeclaredStatus
        If tmPetInfo(llPet).lCode <= 0 Then
            SQLQuery = "Insert Into pet ( "
            SQLQuery = SQLQuery & "petCode, "
            SQLQuery = SQLQuery & "petAttCode, "
            SQLQuery = SQLQuery & "petVefCode, "
            SQLQuery = SQLQuery & "petShttCode, "
            SQLQuery = SQLQuery & "petGsfCode, "
            SQLQuery = SQLQuery & "petDeclaredStatus, "
            SQLQuery = SQLQuery & "petClearStatus, "
            SQLQuery = SQLQuery & "petUstCode, "
            SQLQuery = SQLQuery & "petEnteredDate, "
            SQLQuery = SQLQuery & "petEnteredTime, "
            SQLQuery = SQLQuery & "petUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & "Replace" & ", "
            SQLQuery = SQLQuery & llAttCode & ", "
            SQLQuery = SQLQuery & ilVefCode & ", "
            SQLQuery = SQLQuery & ilShttCode & ", "
            SQLQuery = SQLQuery & tmPetInfo(llPet).lgsfCode & ", "
            SQLQuery = SQLQuery & "'" & slDeclaration & "', "
            SQLQuery = SQLQuery & "'" & gFixQuote("U") & "', "
            SQLQuery = SQLQuery & igUstCode & ", "
            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & gFixQuote("") & "' "
            SQLQuery = SQLQuery & ") "
            llCode = gInsertAndReturnCode(SQLQuery, "pet", "petCode", "Replace")
            If lmAttCode = llAttCode Then
                If llCode > 0 Then
                    tmPetInfo(llPet).lCode = llCode
                    tmPetInfo(llPet).sChanged = "N"
                Else
                    mSaveEventInfo = False
                    Exit Function
                End If
            End If
        Else
            If tmPetInfo(llPet).sChanged = "Y" Then
                SQLQuery = "Update pet Set "
                SQLQuery = SQLQuery & "petDeclaredStatus = '" & slDeclaration & "', "
                SQLQuery = SQLQuery & "petUstCode = " & igUstCode & ", "
                SQLQuery = SQLQuery & "petEnteredDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "petEnteredTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "petUnused = '" & gFixQuote("") & "' "
                SQLQuery = SQLQuery & " WHERE petCode = " & tmPetInfo(llPet).lCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    mMousePointer vbDefault
                    gHandleError "AffErrorLog.txt", "AffAgmnt-mSaveEventInfo"
                    mSaveEventInfo = False
                    Exit Function
                End If
                tmPetInfo(llPet).sChanged = "N"
            End If
        End If
    Next llPet
    mSaveEventInfo = True
    Exit Function

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mSaveEventInfo"
    mSaveEventInfo = False
End Function

Private Sub mSetMulticast()
    If TabStrip1.SelectedItem.Index = TABPLEDGE Then
        If gIsMulticast(imShttCode) Then
            lacMulticast(1).Left = lacMulticast(0).Left
            lacMulticast(1).Top = lacMulticast(0).Top
            lacMulticast(0).Visible = True
            lacMulticast(1).Visible = True
            grdMulticast.Visible = True
        Else
            lacMulticast(0).Visible = False
            lacMulticast(1).Visible = False
            grdMulticast.Visible = False
        End If
    Else
        lacMulticast(0).Visible = False
        lacMulticast(1).Visible = False
        grdMulticast.Visible = False
    End If
End Sub

Private Sub mGetEventStartDate()
    Dim slDate As String
    On Error GoTo ErrHand:
    If (imVefCode <= 0) Or (smPledgeByEvent <> "Y") Or (IsAgmntDirty) Then
        Exit Sub
    End If
    If txtOnAirDate.Text <> "" Then
        Exit Sub
    End If
    SQLQuery = "SELECT Min(gsfAirDate) FROM GSF_Game_Schd WHERE (gsfVefCode = " & imVefCode & " AND gsfAirDate > '" & Format(gNow(), sgSQLDateForm) & "')"
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    If Not rst_Gsf.EOF Then
        If Not IsNull(rst_Gsf(0).Value) Then
            slDate = gObtainPrevMonday(Format(rst_Gsf(0).Value, sgShowDateForm))
            If gDateValue(slDate) <= gDateValue(Format(gNow(), "m/d/yy")) Then
                slDate = DateAdd("d", 7, slDate)
            End If
            txtOnAirDate.Text = slDate
        End If
    End If
    Exit Sub

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "Agreement-mGetEventStartDate"
End Sub

Private Sub mGetPledgeBy()
    Dim ilVff As Integer
    smPledgeByEvent = "N"
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        Exit Sub
    End If
    If imVefCode <= 0 Then
        Exit Sub
    End If
    ilVff = gBinarySearchVff(imVefCode)
    If ilVff <> -1 Then
        smPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
        If smPledgeByEvent = "" Then
            smPledgeByEvent = "N"
        End If
    End If
End Sub

Private Sub mClearPledgeGrid()
    Dim llRow As Long
    Dim llCol As Long
    gGrid_Clear grdPledge, True
    For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
        grdPledge.Row = llRow
        For llCol = MONFDINDEX To ESTIMATETIMEINDEX Step 1
            grdPledge.Col = llCol
            grdPledge.CellBackColor = vbWhite
        Next llCol
    Next llRow
    ReDim tgDat(0 To 0) As DAT
End Sub

Private Sub mClearEventGrid()
    Dim llRow
    Dim llCol As Long
    cbcSeason.Clear
    gGrid_Clear grdEvent, True
    For llRow = grdEvent.FixedRows To grdEvent.Rows - 1 Step 1
        grdEvent.Row = llRow
        For llCol = EVTEVENTNOINDEX To EVTUNDECIDEDINDEX Step 1
            grdEvent.Col = llCol
            grdEvent.CellBackColor = vbWhite
        Next llCol
    Next llRow
End Sub

Private Function mVehicleAddTest(ilVefIndex As Integer) As Boolean
    Dim ilVff As Integer
    Dim ilVef As Integer
    Dim ilSetValue As Integer
    Dim rstATT As ADODB.Recordset
    
    On Error GoTo ErrHand:
    ilSetValue = True
    If tgVehicleInfo(ilVefIndex).sVehType = "L" Then
        'Check to see if any vehicle which belong to the Log vehicle is to be Merged
        ilSetValue = False
        For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilVefIndex).iCode = tgVehicleInfo(ilVef).iVefCode Then
                For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                    If tgVehicleInfo(ilVef).iCode = tgVffInfo(ilVff).iVefCode Then
                        If tgVffInfo(ilVff).sMergeAffiliate <> "S" Then
                            ilSetValue = True
                            Exit For
                        End If
                    End If
                Next ilVff
            End If
        Next ilVef
        If Not ilSetValue Then
            'Test if agreement exist for Log vehicle
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(tgVehicleInfo(ilVefIndex).iCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = tgVehicleInfo(ilVefIndex).iCode Then
                ilSetValue = True
            End If
        End If
    ElseIf ((tgVehicleInfo(ilVefIndex).sVehType = "C") Or (tgVehicleInfo(ilVefIndex).sVehType = "G") Or (tgVehicleInfo(ilVefIndex).sVehType = "A")) And (tgVehicleInfo(ilVefIndex).iVefCode > 0) Then
        'Check to see if the vehicle that references a Log vehicle is to have a separte agreement from the log vehicle
        ilSetValue = False
        For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
            If tgVehicleInfo(ilVefIndex).iCode = tgVffInfo(ilVff).iVefCode Then
                If tgVffInfo(ilVff).sMergeAffiliate = "S" Then
                    ilSetValue = True
                    Exit For
                End If
            End If
        Next ilVff
        If Not ilSetValue Then
            'Test if agreement exist for vehicle that references a Log vehicle
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(tgVehicleInfo(ilVefIndex).iCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = tgVehicleInfo(ilVefIndex).iCode Then
                ilSetValue = True
            End If
        End If
    End If
    mVehicleAddTest = ilSetValue
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mVehicleAddTest"
    mVehicleAddTest = False
End Function



Private Sub mPopSeason()
    Dim ilSeason As Integer
    
    On Error GoTo ErrHand:
    If imVefCode <= 0 Then
        Exit Sub
    End If
    If smPledgeByEvent <> "Y" Then
        Exit Sub
    End If
    cbcSeason.Clear
    SQLQuery = "SELECT * FROM GHF_Game_Header WHERE ghfVefCode = " & imVefCode & " ORDER BY ghfSeasonStartDate Desc"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcSeason.AddItem Trim$(rst!ghfSeasonName)
        cbcSeason.SetItemData = rst!ghfCode
        rst.MoveNext
    Loop
    SQLQuery = "SELECT vffSeasonGhfCode "
    SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
    SQLQuery = SQLQuery + " WHERE (vffVefCode = " & imVefCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        For ilSeason = 0 To cbcSeason.ListCount - 1 Step 1
            If cbcSeason.GetItemData(ilSeason) = rst!vffSeasonGhfCode Then
                cbcSeason.SetListIndex = ilSeason
                Exit For
            End If
        Next ilSeason
    Else
        If cbcSeason.ListCount > 0 Then
            cbcSeason.SetListIndex = 0
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mPopSeason"
    
End Sub
Private Sub mIDCShowGroup(blShow As Boolean)
    Dim ilLoop As Integer
    Dim myRs As ADODB.Recordset
    'only change if need to.  to only do sql when needed.  this works because 'clear' sets to false, so true is always new for
    'station, but not new for agreement within station
    If blShow <> lbcIDCGroup.Visible Then
        lbcIDCGroup.Visible = blShow
        For ilLoop = 0 To 2
            optIDCGroup(ilLoop).Visible = blShow
        Next ilLoop
        If blShow And imShttCode > 0 Then
            SQLQuery = "select shttCode from shtt where shttOnAddress1 <> '' AND shttCode <> '" & imShttCode & "' AND "
            SQLQuery = SQLQuery & " shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip + cast(shttmktcode as varchar(5)) = "
            SQLQuery = SQLQuery & "(select shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip + cast(shttmktcode as varchar(5)) from shtt where shttCode = " & imShttCode & ")"
            Set myRs = gSQLSelectCall(SQLQuery)
            If myRs.EOF Then
                optIDCGroup(2).Enabled = False
                myRs.Close
            Else
                optIDCGroup(2).Enabled = True
            End If
        End If
    End If
    Set myRs = Nothing
End Sub

Private Sub mSetEventTitles()
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String
    Dim ilPos As Integer
    
    If imVefCode = 0 Then
        grdEvent.TextMatrix(0, EVTVISITTEAMINDEX) = ""
        grdEvent.TextMatrix(1, EVTVISITTEAMINDEX) = ""
        grdEvent.TextMatrix(0, EVTHOMETEAMINDEX) = ""
        grdEvent.TextMatrix(1, EVTHOMETEAMINDEX) = ""
        Exit Sub
    End If
    
    gGetEventTitles imVefCode, slEventTitle1, slEventTitle2
    
    ilPos = InStr(1, slEventTitle1, " ", vbTextCompare)
    If ilPos > 0 Then
        grdEvent.TextMatrix(0, EVTVISITTEAMINDEX) = Left(slEventTitle1, ilPos)
        grdEvent.TextMatrix(1, EVTVISITTEAMINDEX) = Trim$(Mid(slEventTitle1, ilPos + 1))
    Else
        grdEvent.TextMatrix(0, EVTVISITTEAMINDEX) = slEventTitle1
        grdEvent.TextMatrix(1, EVTVISITTEAMINDEX) = ""
    End If

    ilPos = InStr(1, slEventTitle2, " ", vbTextCompare)
    If ilPos > 0 Then
        grdEvent.TextMatrix(0, EVTHOMETEAMINDEX) = Left(slEventTitle2, ilPos)
        grdEvent.TextMatrix(1, EVTHOMETEAMINDEX) = Trim$(Mid(slEventTitle2, ilPos + 1))
    Else
        grdEvent.TextMatrix(0, EVTHOMETEAMINDEX) = slEventTitle2
        grdEvent.TextMatrix(1, EVTHOMETEAMINDEX) = ""
    End If

End Sub
Private Function mTestEstimateTimes() As Integer
    Dim ilDat As Integer
    Dim ilNextET As Integer
    
    mTestEstimateTimes = True
    For ilDat = 0 To UBound(tgDat) - 1 Step 1
        If tgStatusTypes(tgDat(ilDat).iFdStatus).iPledged = 1 Then
            ilNextET = tgDat(ilDat).iFirstET
            Do While ilNextET <> -1
                If ((Trim$(tmETAvailInfo(ilNextET).sETTime) <> "") And (Trim$(tmETAvailInfo(ilNextET).sETDay) = "")) Then
                    gMsgBox "The Estimate day missing from Feed Time " & Trim$(tmETAvailInfo(ilNextET).sFdTime)
                    mTestEstimateTimes = False
                    Exit Function
                End If
                If ((Trim$(tmETAvailInfo(ilNextET).sETTime) = "") And (Trim$(tmETAvailInfo(ilNextET).sETDay) <> "")) Then
                    gMsgBox "The Estimate time missing from Feed Time " & Trim$(tmETAvailInfo(ilNextET).sFdTime)
                    mTestEstimateTimes = False
                    Exit Function
                End If
                ilNextET = tmETAvailInfo(ilNextET).iNextET
            Loop
        End If
    Next ilDat
End Function

Private Function mCheckSendDelays() As Integer
    Dim llRow As Long
    Dim llFdStartTime As Long
    Dim llFdEndTime As Long
    Dim llPdStartTime As Long
    Dim llPdEndTime As Long
    Dim slStr As String
    Dim llRowIndex As Long
    Dim ilIndex As Integer
    Dim blError As Boolean
    Dim slPdDayFed As String
    Dim ilRet As Integer
    '7701
    Dim blIsXDS As Boolean
    
    mCheckSendDelays = True
    blError = False
    ''7701
    blIsXDS = False
    If mMultiListIsData(Vendors.XDS_Break, lbcAudioDelivery) Or mMultiListIsData(Vendors.XDS_ISCI, lbcAudioDelivery) Then
        blIsXDS = True
    End If
'    With lbcAudioDelivery
'        If .ListIndex > -1 Then
'            If .ItemData(.ListIndex) = Vendors.XDS_Break Or .ItemData(.ListIndex) = Vendors.XDS_ISCI Then
'                blIsXDS = True
'            End If
'        End If
'    End With
    If bmSupportXDSDelay And ckcSendDelays.Value = vbChecked And blIsXDS Then
    'If bmSupportXDSDelay And ckcSendDelays.Value = vbChecked And (rbcAudio(0).Value Or rbcAudio(1).Value) Then
        For llRow = grdPledge.FixedRows To grdPledge.Rows - 1 Step 1
            slStr = Trim$(grdPledge.TextMatrix(llRow, STATUSINDEX))
            If slStr <> "" Then
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    ilIndex = lbcStatus.ItemData(llRowIndex)
                    If (tmStatusTypes(ilIndex).iPledged = 1) Then
                        llFdStartTime = gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEFDINDEX), False)
                        llFdEndTime = gTimeToLong(grdPledge.TextMatrix(llRow, ENDTIMEFDINDEX), False)
                        llPdStartTime = gTimeToLong(grdPledge.TextMatrix(llRow, STARTTIMEPDINDEX), False)
                        llPdEndTime = gTimeToLong(grdPledge.TextMatrix(llRow, ENDTIMEPDINDEX), False)
                        If (llFdEndTime - llFdStartTime) <> (llPdEndTime - llPdStartTime) Then
                            blError = True
                            grdPledge.Row = llRow
                            grdPledge.Col = STARTTIMEPDINDEX
                            grdPledge.CellForeColor = vbRed
                        End If
                        If mPdPriorFd(llRow) Then
                            slPdDayFed = grdPledge.TextMatrix(llRow, DAYFEDINDEX)
                            If slPdDayFed <> "A" Then
                                blError = True
                                grdPledge.Row = llRow
                                grdPledge.Col = DAYFEDINDEX
                                grdPledge.CellForeColor = vbRed
                            End If
                        End If
                    End If
                End If
            End If
        Next llRow
    End If
    If blError Then
        mMousePointer vbDefault
        ilRet = MsgBox("Feed and Pledge break length not matching and/or Before/After not set to After", vbOKOnly, "Time Error")
        grdPledge.Redraw = True
        If Not frcTab(2).Visible Then
            'gSendKeys "%P", True
            TabStrip1.Tabs(TABPLEDGE).Selected = True
        End If
        mCheckSendDelays = False
    End If
End Function
'8418
Private Sub mPopVendorList()
    ReDim tmVendorList(50)
    Dim ilCount As Integer
    Dim tlVendors() As VendorInfo
On Error GoTo ErrHand
    tlVendors = gGetAvailableVendors()
    SQLQuery = "Select wvtVendorID,wvtName,wvtDeliveryType From wvt_Vendor_Table"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        tmVendorList(ilCount).sName = Trim$(rst!wvtName)
        tmVendorList(ilCount).iCode = rst!wvtVendorId
        tmVendorList(ilCount).sType = rst!wvtDeliveryType
        tmVendorList(ilCount).iVersion = gVendorMinVersion(rst!wvtVendorId, tlVendors)
        ilCount = ilCount + 1
        rst.MoveNext
    Loop
    ReDim Preserve tmVendorList(ilCount)
    Erase tlVendors
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mPopVendorList"
End Sub
'8418
Private Sub mLoadAllowedVendorServices(ilShttCode As Integer)
    Dim ilStationVersion As Integer
    Dim c As Integer
    Dim ilShtt As Integer
    
    lbcLogDelivery.Clear
    lbcAudioDelivery.Clear
On Error GoTo ErrHand
    ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
    If ilShtt <> -1 Then
        If IsNumeric(tgStationInfoByCode(ilShtt).sWebNumber) Then
            ilStationVersion = tgStationInfoByCode(ilShtt).sWebNumber
        Else
            ilStationVersion = 1
        End If
    Else
        ilStationVersion = 1
    End If
    For c = 0 To UBound(tmVendorList) - 1
        If ilStationVersion >= tmVendorList(c).iVersion Then
            If tmVendorList(c).sType = "L" Then
                lbcLogDelivery.AddItem (tmVendorList(c).sName)
                lbcLogDelivery.ItemData(lbcLogDelivery.NewIndex) = tmVendorList(c).iCode
            Else
                lbcAudioDelivery.AddItem (tmVendorList(c).sName)
                lbcAudioDelivery.ItemData(lbcAudioDelivery.NewIndex) = tmVendorList(c).iCode
            End If
        End If
    Next c
    mEnableDeliveryOptions False
    mEnableDeliveryOptions True
    'don't show choices if can never select
    If lbcLogDelivery.ListCount = 0 Then
        rbcSendLogEMail(0).Visible = False
        rbcSendLogEMail(1).Visible = False
        lacSendLogEMail.Visible = False
    Else
        rbcSendLogEMail(0).Visible = True
        rbcSendLogEMail(1).Visible = True
        lacSendLogEMail.Visible = True
    End If
    If lbcAudioDelivery.ListCount = 0 Then
            lacIDCReceiverID.Visible = False
            txtIDCReceiverID.Visible = False
            lacXDReceiverID.Visible = False
            txtXDReceiverID.Visible = False
            optVoiceTracked(0).Visible = False
            optVoiceTracked(1).Visible = False
            lacVoiceTracked.Visible = False
    Else
        If Not mIsInListBox(lbcAudioDelivery, Vendors.iDc) Then
            lacIDCReceiverID.Visible = False
            txtIDCReceiverID.Visible = False
        Else
            lacIDCReceiverID.Visible = True
            txtIDCReceiverID.Visible = True
        End If
        If Not mIsInListBox(lbcAudioDelivery, Vendors.XDS_ISCI) And Not mIsInListBox(lbcAudioDelivery, Vendors.XDS_Break) Then
            lacXDReceiverID.Visible = False
            txtXDReceiverID.Visible = False
            optVoiceTracked(0).Visible = False
            optVoiceTracked(1).Visible = False
            lacVoiceTracked.Visible = False
        Else
            lacXDReceiverID.Visible = True
            txtXDReceiverID.Visible = True
            optVoiceTracked(0).Visible = True
            optVoiceTracked(1).Visible = True
            lacVoiceTracked.Visible = True
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mLoadAllowedVendorServices"
End Sub
'Private Sub mLoadDeliveryServices()
'    Dim ilRet As Integer
'    Dim blIsJelli As Boolean
'    Dim blIsMarketron As Boolean
'    Dim blIsIPump As Boolean
'    Dim ilXDS As Integer
'
'    ilRet = 0
'    ilXDS = 0
'    blIsJelli = False
'    blIsMarketron = False
'    blIsIPump = False
'On Error GoTo ErrHand
'    SQLQuery = "Select wvtVendorID,wvtName,wvtDeliveryType From wvt_Vendor_Table"
'    Set rst = gSQLSelectCall(SQLQuery)
'    With lbcLogDelivery
'        If Not rst.EOF Then
'            Do While Not rst.EOF
'                If rst!wvtDeliveryType = "L" Then
'                    .AddItem Trim$(rst!wvtName)
'                    .ItemData(.NewIndex) = rst!wvtVendorId
'                End If
'                rst.MoveNext
'            Loop
'            'for next time through
'            rst.MoveFirst
'        End If
'    End With
'    mEnableDeliveryOptions False
'    With lbcAudioDelivery
'         Do While Not rst.EOF
'            If rst!wvtDeliveryType = "A" Then
'                .AddItem Trim$(rst!wvtName)
'                .ItemData(.NewIndex) = rst!wvtVendorId
'            End If
'            rst.MoveNext
'        Loop
'    End With
'    mEnableDeliveryOptions True
'    'don't show choices if can never select
'    If lbcLogDelivery.ListCount = 0 Then
'        rbcSendLogEMail(0).Visible = False
'        rbcSendLogEMail(1).Visible = False
'        lacSendLogEMail.Visible = False
'    End If
'    If lbcAudioDelivery.ListCount = 0 Then
'            lacIDCReceiverID.Visible = False
'            txtIDCReceiverID.Visible = False
'            lacXDReceiverID.Visible = False
'            txtXDReceiverID.Visible = False
'            optVoiceTracked(0).Visible = False
'            optVoiceTracked(1).Visible = False
'            lacVoiceTracked.Visible = False
'    Else
'        If Not mIsInListBox(lbcAudioDelivery, Vendors.iDc) Then
'            lacIDCReceiverID.Visible = False
'            txtIDCReceiverID.Visible = False
'        End If
'        If Not mIsInListBox(lbcAudioDelivery, Vendors.XDS_ISCI) And Not mIsInListBox(lbcAudioDelivery, Vendors.XDS_Break) Then
'            lacXDReceiverID.Visible = False
'            txtXDReceiverID.Visible = False
'            optVoiceTracked(0).Visible = False
'            optVoiceTracked(1).Visible = False
'            lacVoiceTracked.Visible = False
'        End If
'    End If
'    Exit Sub
'ErrHand:
'    gHandleError "AffErrorLog.txt", "Agreement-mLoadDeliveryServices"
'
'End Sub
Private Function mIsInListBox(lbcBox As ListBox, ilType As Vendors)
    Dim blRet As Boolean
    Dim c As Integer
    
    blRet = False
    For c = 0 To lbcBox.ListCount - 1 Step 1
        If lbcBox.ItemData(c) = ilType Then
            blRet = True
            Exit For
        End If
    Next c
    mIsInListBox = blRet
End Function
Private Sub mEnableDeliveryOptions(blIsAudio As Boolean)
    Dim blEnabled As Boolean
    Dim blIDC As Boolean
    
    If blIsAudio Then
        'only xds has voice tracking and xds site id.  idc only allows idc box.
        '7701 both idc and xds?
        blEnabled = mMultiListIsData(Vendors.XDS_Break, lbcAudioDelivery)
        If Not blEnabled Then
            blEnabled = mMultiListIsData(Vendors.XDS_ISCI, lbcAudioDelivery)
        End If
        blIDC = mMultiListIsData(Vendors.iDc, lbcAudioDelivery)
        If Not blEnabled Then
            optVoiceTracked(1).Value = True
            ckcSendDelays.Value = vbUnchecked
            ckcSendNotCarried.Value = vbUnchecked
        End If
        lacXDReceiverID.Enabled = blEnabled
        txtXDReceiverID.Enabled = blEnabled
        lacVoiceTracked.Enabled = blEnabled
        optVoiceTracked(0).Enabled = blEnabled
        optVoiceTracked(1).Enabled = blEnabled
        If bmSupportXDSDelay Then
            ckcSendDelays.Enabled = blEnabled
        Else
            ckcSendDelays.Enabled = False
        End If
        ckcSendNotCarried.Enabled = blEnabled
        lacIDCReceiverID.Enabled = blIDC
        txtIDCReceiverID.Enabled = blIDC
    Else
        blEnabled = True
        If Not blEnabled Then
            rbcSendLogEMail(1).Value = True
        End If
        lacSendLogEMail.Enabled = blEnabled
        rbcSendLogEMail(0).Enabled = blEnabled
        rbcSendLogEMail(1).Enabled = blEnabled
    End If
End Sub
Private Function mSetMultilist(myLbc As ListBox, llItem As Long) As Boolean
    'if llItem is less than 1, clear selection.  Returns if found
    Dim blRet As Boolean
    Dim ilLoop As Integer
    '7902
    bmInMultiListChange = True
    blRet = False
    With myLbc
        If llItem < 1 Then
            For ilLoop = 0 To .ListCount - 1
                .Selected(ilLoop) = False
            Next
        Else
            For ilLoop = 0 To .ListCount - 1
                If .ItemData(ilLoop) = llItem Then
                    blRet = True
                    .Selected(ilLoop) = True
                    Exit For
                End If
            Next
        End If
        .ListIndex = -1
    End With
    bmInMultiListChange = False
    mSetMultilist = blRet
End Function
Private Function mRetrieveMultiListIntegers(myLbc As ListBox) As Integer()
    'returns array. use 'ubound -1' to get values
    Dim ilRet() As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    
    ilUpper = 0
    ReDim ilRet(ilUpper)
    With myLbc
        For ilLoop = 0 To .ListCount - 1
            If .Selected(ilLoop) Then
                ilRet(ilUpper) = .ItemData(ilLoop)
                ilUpper = ilUpper + 1
                ReDim Preserve ilRet(ilUpper)
            End If
        Next
    End With
    mRetrieveMultiListIntegers = ilRet
End Function
Private Function mRetrieveMultiListString(myLbc As ListBox) As String
    Dim slRet As String
    Dim ilLoop As Integer
    
    slRet = ""
    With myLbc
        For ilLoop = 0 To .ListCount - 1
            If .Selected(ilLoop) Then
                slRet = slRet & .List(ilLoop) & ","
            End If
        Next
    End With
    mRetrieveMultiListString = gLoseLastLetterIfComma(slRet)
End Function
Private Function mMultiListIsData(ilItemData As Integer, myLbc As ListBox) As Boolean
    Dim blRet As Boolean
    Dim ilLoop As Integer
    
    blRet = False
    With myLbc
        For ilLoop = 0 To .ListCount - 1
            If .Selected(ilLoop) Then
                If .ItemData(ilLoop) = ilItemData Then
                    blRet = True
                    Exit For
                End If
            End If
        Next
    End With
    
    
    mMultiListIsData = blRet
End Function
Private Function mIsDeliveryInconsistent() As Boolean
    Dim blRet As Boolean
    
    blRet = False
    If rbcExportType(0).Value And Len(mRetrieveMultiListString(lbcLogDelivery)) > 0 Then
        If MsgBox("Warning - You chose 'Manual' Affidavit Control on the Delivery tab, but you also selected log delivery interfaces.  Choosing 'Save' will deselect the log delivery interfaces and continue the save.  Choose 'cancel' to stop the save.", vbOKCancel + vbQuestion, "Manual Export Selected") = vbCancel Then
            blRet = True
        Else
            mSetMultilist lbcLogDelivery, -1
        End If
    End If
    mIsDeliveryInconsistent = blRet
End Function
Sub mUpdateShttTables(ilShttCode As Integer, blUsedForAtt As Boolean, blAgreementExist As Boolean, slHistDate As String)
    '11/26/17
    Dim ilIndex As Integer
    Dim blRepopRequired As Boolean
    Dim slCallLetters As String
    
    blRepopRequired = False
    ilIndex = gBinarySearchStationInfoByCode(ilShttCode)
    If ilIndex <> -1 Then
        If blUsedForAtt Then
            tgStationInfoByCode(ilIndex).sUsedForATT = "Y"
        End If
        If blAgreementExist Then
            tgStationInfoByCode(ilIndex).sAgreementExist = "Y"
        End If
        If slHistDate <> "" Then
            tgStationInfoByCode(ilIndex).lHistStartDate = gDateValue(slHistDate)
        End If
        slCallLetters = Trim$(tgStationInfoByCode(ilIndex).sCallLetters)
        ilIndex = gBinarySearchStation(slCallLetters)
        If ilIndex <> -1 Then
            If blUsedForAtt Then
                tgStationInfo(ilIndex).sUsedForATT = "Y"
            End If
            If blAgreementExist Then
                tgStationInfo(ilIndex).sAgreementExist = "Y"
            End If
            If slHistDate <> "" Then
                tgStationInfo(ilIndex).lHistStartDate = gDateValue(slHistDate)
            End If
        Else
            blRepopRequired = True
        End If
    Else
        blRepopRequired = True
    End If
    gFileChgdUpdate "shtt.mkd", blRepopRequired
End Sub
'9452
Private Sub mSetCpttAstStatus(llAtt As Long, ilVefCode As Integer, slOnAir As String)
    Dim slSDate As String
    Dim slLLD As String
    Dim slDate As String
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    'only here because user changed to allow to send 'not carried'.  Have to rerun asts in this case to get regionals set.
    If DateValue(gAdjYear(slOnAir)) = DateValue("1/1/1970") Then
        Exit Sub
    End If
    slSQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
    slSQLQuery = slSQLQuery + " FROM VPF_Vehicle_Options"
    slSQLQuery = slSQLQuery + " WHERE (vpfvefKCode =" & ilVefCode & ")"
    Set rst = gSQLSelectCall(slSQLQuery)
    If rst.EOF Then
        Exit Sub
    End If
    If IsNull(rst!vpfLLD) Then
        Exit Sub
    End If
    If Not gIsDate(rst!vpfLLD) Then
        Exit Sub
    Else
        slLLD = Format$(rst!vpfLLD, sgShowDateForm)
    End If
    If DateValue(gAdjYear(slOnAir)) <= DateValue(gAdjYear(slLLD)) Then 'smSvOnAirDate
        If sgNowDate = "" Then
            slDate = Now
        Else
            slDate = sgNowDate
        End If
        'slSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(sgNowDate, "m/d/yy"))))
        slSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(slDate, "m/d/yy"))))
        slSQLQuery = "UPDATE cptt SET "
        slSQLQuery = slSQLQuery + "cpttAstStatus = " & "'R'"
        slSQLQuery = slSQLQuery + " WHERE (cpttAtfCode = " & llAtt
        slSQLQuery = slSQLQuery + " AND cpttStartDate >= '" & Format(slSDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery + " AND cpttAstStatus = 'C'" & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmAgmnt-mSetCpttAstStatus"
        End If
    End If
End Sub
