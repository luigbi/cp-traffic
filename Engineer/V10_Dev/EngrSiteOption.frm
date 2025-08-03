VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EngrSiteOption 
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11565
   ControlBox      =   0   'False
   Icon            =   "EngrSiteOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11565
   Begin VB.Frame frcTab 
      Caption         =   "E-Mail"
      Height          =   4665
      Index           =   4
      Left            =   8385
      TabIndex        =   200
      Top             =   6165
      Visible         =   0   'False
      Width           =   10155
      Begin VB.Frame frcTLS 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   300
         Left            =   90
         TabIndex        =   210
         Top             =   2595
         Width           =   5265
         Begin VB.OptionButton rbcTLS 
            Caption         =   "False"
            Height          =   195
            Index           =   1
            Left            =   3390
            TabIndex        =   209
            Top             =   0
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton rbcTLS 
            Caption         =   "True"
            Height          =   195
            Index           =   0
            Left            =   2310
            TabIndex        =   212
            Top             =   -15
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Transport Layer Security:"
            Height          =   255
            Left            =   0
            TabIndex        =   217
            Top             =   0
            Width           =   2070
         End
      End
      Begin VB.Frame frcVerification 
         Caption         =   "Verification"
         Height          =   1185
         Left            =   90
         TabIndex        =   211
         Top             =   3375
         Width           =   9645
         Begin VB.TextBox edcResult 
            Height          =   285
            Left            =   3345
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "Please verify before saving information"
            Top             =   750
            Width           =   6015
         End
         Begin VB.CommandButton cmcVerify 
            Caption         =   "&Verify"
            Height          =   375
            Left            =   90
            TabIndex        =   215
            Top             =   705
            Width           =   1335
         End
         Begin VB.TextBox edcTo 
            Height          =   285
            Left            =   2130
            MaxLength       =   80
            TabIndex        =   213
            Top             =   270
            Width           =   7320
         End
         Begin VB.Label lacResult 
            Caption         =   "Result Message:"
            Height          =   255
            Left            =   1650
            TabIndex        =   216
            Top             =   765
            Width           =   1680
         End
         Begin VB.Label lacTo 
            Caption         =   "To Address:"
            Height          =   255
            Left            =   105
            TabIndex        =   214
            Top             =   315
            Width           =   1680
         End
      End
      Begin VB.TextBox edcPort 
         Height          =   285
         Left            =   2220
         MaxLength       =   3
         TabIndex        =   207
         Top             =   2070
         Width           =   825
      End
      Begin VB.TextBox edcPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2220
         MaxLength       =   80
         PasswordChar    =   "*"
         TabIndex        =   205
         Top             =   1485
         Width           =   7320
      End
      Begin VB.TextBox edcAccount 
         Height          =   285
         Left            =   2220
         MaxLength       =   80
         TabIndex        =   203
         Top             =   885
         Width           =   7320
      End
      Begin VB.TextBox edcHost 
         Height          =   285
         Left            =   2220
         MaxLength       =   80
         TabIndex        =   201
         Top             =   285
         Width           =   7320
      End
      Begin VB.Label lacAccountExample 
         Caption         =   "e.g., abcd@xyz.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2220
         TabIndex        =   219
         Top             =   1155
         Width           =   2280
      End
      Begin VB.Label lacHostExample 
         Caption         =   "e.g., smtp.att.yahoo.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2205
         TabIndex        =   218
         Top             =   555
         Width           =   2280
      End
      Begin VB.Label lacPort 
         Caption         =   "Port Number:"
         Height          =   255
         Left            =   90
         TabIndex        =   208
         Top             =   2055
         Width           =   1350
      End
      Begin VB.Label lacPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   90
         TabIndex        =   206
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Label lacAccount 
         Caption         =   "Account Name:"
         Height          =   255
         Left            =   90
         TabIndex        =   204
         Top             =   870
         Width           =   1680
      End
      Begin VB.Label laclHost 
         Caption         =   "SMTP Server:"
         Height          =   255
         Left            =   90
         TabIndex        =   202
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame frcNoUsed 
      Caption         =   "Not Used"
      Height          =   1605
      Left            =   8670
      TabIndex        =   165
      Top             =   5910
      Visible         =   0   'False
      Width           =   11085
      Begin VB.TextBox edcBkupClientImportPath 
         Height          =   285
         Left            =   2460
         MaxLength       =   100
         TabIndex        =   167
         Top             =   795
         Visible         =   0   'False
         Width           =   7350
      End
      Begin VB.TextBox edcBkupServerImportPath 
         Height          =   285
         Left            =   2460
         MaxLength       =   100
         TabIndex        =   166
         Top             =   315
         Visible         =   0   'False
         Width           =   7350
      End
      Begin VB.Label lacBkupClientImportPath 
         Caption         =   "Protection Client Import Path:"
         Height          =   255
         Left            =   60
         TabIndex        =   169
         Top             =   795
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lacBkupServerImportPath 
         Caption         =   "Protection Server Import Path:"
         Height          =   255
         Left            =   60
         TabIndex        =   168
         Top             =   315
         Visible         =   0   'False
         Width           =   2340
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Commercial Merge"
      Height          =   5055
      Index           =   3
      Left            =   240
      TabIndex        =   121
      Top             =   750
      Visible         =   0   'False
      Width           =   10785
      Begin VB.Frame frcTest 
         Caption         =   "Test System"
         Height          =   1500
         Index           =   0
         Left            =   90
         TabIndex        =   144
         Top             =   3390
         Width           =   10125
         Begin VB.OptionButton rbcMergeTest 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2475
            TabIndex        =   146
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton rbcMergeTest 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3525
            TabIndex        =   147
            Top             =   270
            Width           =   705
         End
         Begin VB.TextBox edcPriServerImportPathTest 
            Height          =   285
            Left            =   2490
            MaxLength       =   100
            TabIndex        =   149
            Top             =   630
            Width           =   7350
         End
         Begin VB.TextBox edcPriClientImportPathTest 
            Height          =   285
            Left            =   2490
            MaxLength       =   100
            TabIndex        =   151
            Top             =   1080
            Width           =   7350
         End
         Begin VB.Label lacMergeStopTest 
            Caption         =   "Stop Merge and Copy Task"
            Height          =   255
            Left            =   90
            TabIndex        =   145
            Top             =   270
            Width           =   2280
         End
         Begin VB.Label lacPriServerImportPathTest 
            Caption         =   "Primary Server Import Path:"
            Height          =   255
            Left            =   90
            TabIndex        =   148
            Top             =   630
            Width           =   2340
         End
         Begin VB.Label lacPriClientImportPathTest 
            Caption         =   "Primary Client Import Path:"
            Height          =   255
            Left            =   90
            TabIndex        =   150
            Top             =   1080
            Width           =   2340
         End
      End
      Begin VB.Frame frcProd 
         Caption         =   "Production System"
         Height          =   1500
         Left            =   90
         TabIndex        =   136
         Top             =   1740
         Width           =   10125
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2490
            TabIndex        =   138
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton rbcMerge 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3525
            TabIndex        =   139
            Top             =   270
            Width           =   705
         End
         Begin VB.TextBox edcPriClientImportPath 
            Height          =   285
            Left            =   2490
            MaxLength       =   100
            TabIndex        =   143
            Top             =   1080
            Width           =   7350
         End
         Begin VB.TextBox edcPriServerImportPath 
            Height          =   285
            Left            =   2490
            MaxLength       =   100
            TabIndex        =   141
            Top             =   630
            Width           =   7350
         End
         Begin VB.Label lacMergeStop 
            Caption         =   "Stop Merge Task"
            Height          =   255
            Left            =   90
            TabIndex        =   137
            Top             =   270
            Width           =   1845
         End
         Begin VB.Label lacPriClientImportPath 
            Caption         =   "Primary Client Import Path:"
            Height          =   255
            Left            =   90
            TabIndex        =   142
            Top             =   1080
            Width           =   2340
         End
         Begin VB.Label lacPriServerImportPath 
            Caption         =   "Primary Server Import Path:"
            Height          =   255
            Left            =   90
            TabIndex        =   140
            Top             =   630
            Width           =   2340
         End
      End
      Begin VB.TextBox edcTimeFormat 
         Height          =   285
         Left            =   6705
         MaxLength       =   20
         TabIndex        =   125
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox edcDateFormat 
         Height          =   285
         Left            =   2490
         MaxLength       =   20
         TabIndex        =   123
         Top             =   315
         Width           =   1545
      End
      Begin VB.TextBox edcMergeChkStartTime 
         Height          =   285
         Left            =   2490
         MaxLength       =   10
         TabIndex        =   131
         Top             =   1305
         Width           =   1275
      End
      Begin VB.TextBox edcMergeChkEnd 
         Height          =   285
         Left            =   5385
         MaxLength       =   10
         TabIndex        =   133
         Top             =   1305
         Width           =   1275
      End
      Begin VB.TextBox edcMergeChkInterval 
         Height          =   285
         Left            =   8640
         MaxLength       =   3
         TabIndex        =   135
         Top             =   1305
         Width           =   795
      End
      Begin VB.TextBox edcImportFileFormat 
         Height          =   285
         Left            =   2490
         MaxLength       =   20
         TabIndex        =   127
         Top             =   795
         Width           =   2475
      End
      Begin VB.TextBox edcImportExt 
         Height          =   285
         Left            =   6705
         MaxLength       =   3
         TabIndex        =   129
         Top             =   780
         Width           =   660
      End
      Begin VB.Label lacTimeFormat 
         Caption         =   "File Name Time Format:"
         Height          =   240
         Left            =   4530
         TabIndex        =   124
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label lacDateFormat 
         Caption         =   "File Name Date  Format:"
         Height          =   255
         Left            =   90
         TabIndex        =   122
         Top             =   300
         Width           =   2565
      End
      Begin VB.Label lacMergeChkStart 
         Caption         =   "Merge Check- Start Time:"
         Height          =   255
         Left            =   90
         TabIndex        =   130
         Top             =   1305
         Width           =   2250
      End
      Begin VB.Label lacMergeChkEnd 
         Caption         =   "End Time:"
         Height          =   255
         Left            =   4380
         TabIndex        =   132
         Top             =   1305
         Width           =   1605
      End
      Begin VB.Label lacMergeChkInterval 
         Caption         =   "Interval in Minutes:"
         Height          =   255
         Left            =   7170
         TabIndex        =   134
         Top             =   1335
         Width           =   1410
      End
      Begin VB.Label lacImportFileFormat 
         Caption         =   "File Name Format:"
         Height          =   255
         Left            =   90
         TabIndex        =   126
         Top             =   780
         Width           =   2565
      End
      Begin VB.Label lacImportExt 
         Caption         =   "Extension:"
         Height          =   240
         Left            =   5310
         TabIndex        =   128
         Top             =   780
         Width           =   1605
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "AutoSchd"
      Height          =   5610
      Index           =   2
      Left            =   9405
      TabIndex        =   102
      Top             =   5025
      Visible         =   0   'False
      Width           =   11100
      Begin VB.Frame frcSATest 
         Caption         =   "Test System"
         Height          =   1305
         Left            =   105
         TabIndex        =   185
         Top             =   4155
         Width           =   10860
         Begin VB.Frame frcPurgeTest 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   75
            TabIndex        =   192
            Top             =   915
            Width           =   10695
            Begin VB.TextBox edcPurgeTimeTest 
               Height          =   285
               Left            =   8970
               MaxLength       =   10
               TabIndex        =   196
               Top             =   0
               Width           =   1290
            End
            Begin VB.OptionButton rbcPurgeTest 
               Caption         =   "At"
               Height          =   195
               Index           =   2
               Left            =   8250
               TabIndex        =   195
               Top             =   0
               Width           =   690
            End
            Begin VB.OptionButton rbcPurgeTest 
               Caption         =   "After Schedule Completed"
               Height          =   195
               Index           =   0
               Left            =   3240
               TabIndex        =   194
               Top             =   0
               Width           =   2355
            End
            Begin VB.OptionButton rbcPurgeTest 
               Caption         =   "After Automation Completed"
               Height          =   195
               Index           =   1
               Left            =   5640
               TabIndex        =   193
               Top             =   0
               Width           =   2430
            End
            Begin VB.Label lacPurgeTest 
               Caption         =   "Purge 'Schedule' and 'Library'"
               Height          =   225
               Left            =   0
               TabIndex        =   197
               Top             =   0
               Width           =   2970
            End
         End
         Begin VB.Frame frcSchOrAutoTest 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   225
            Left            =   75
            TabIndex        =   188
            Top             =   240
            Width           =   10695
            Begin VB.OptionButton rbcSchOrAutoTest 
               Caption         =   "Generate them Independently"
               Height          =   195
               Index           =   2
               Left            =   7920
               TabIndex        =   191
               Top             =   0
               Width           =   2535
            End
            Begin VB.OptionButton rbcSchOrAutoTest 
               Caption         =   "Generate Schedule after Automation Completed"
               Height          =   195
               Index           =   1
               Left            =   3990
               TabIndex        =   190
               Top             =   0
               Width           =   3930
            End
            Begin VB.OptionButton rbcSchOrAutoTest 
               Caption         =   "Generate Automation after Schedule Completed"
               Height          =   195
               Index           =   0
               Left            =   -15
               TabIndex        =   189
               Top             =   0
               Width           =   3990
            End
         End
         Begin VB.TextBox edcSchdGenTimeTest 
            Height          =   285
            Left            =   2340
            MaxLength       =   10
            TabIndex        =   187
            Top             =   525
            Width           =   930
         End
         Begin VB.TextBox edcAutoGenTimeTest 
            Height          =   285
            Left            =   6540
            MaxLength       =   10
            TabIndex        =   186
            Top             =   525
            Width           =   930
         End
         Begin VB.Label lacSchdGenTimeTest 
            Caption         =   "Schedule Generation Time:"
            Height          =   255
            Left            =   60
            TabIndex        =   199
            Top             =   525
            Width           =   2130
         End
         Begin VB.Label lacAutoGenTimeTest 
            Caption         =   "Automation Generation Time:"
            Height          =   255
            Left            =   4095
            TabIndex        =   198
            Top             =   525
            Width           =   2400
         End
      End
      Begin VB.Frame frcSAProd 
         Caption         =   "Production System"
         Height          =   1290
         Left            =   90
         TabIndex        =   170
         Top             =   2850
         Width           =   10860
         Begin VB.TextBox edcAutoGenTime 
            Height          =   285
            Left            =   6540
            MaxLength       =   10
            TabIndex        =   183
            Top             =   525
            Width           =   930
         End
         Begin VB.TextBox edcSchdGenTime 
            Height          =   285
            Left            =   2340
            MaxLength       =   10
            TabIndex        =   181
            Top             =   525
            Width           =   930
         End
         Begin VB.Frame frcSchOrAuto 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   225
            Left            =   75
            TabIndex        =   177
            Top             =   240
            Width           =   10695
            Begin VB.OptionButton rbcSchOrAuto 
               Caption         =   "Generate Automation after Schedule Completed"
               Height          =   195
               Index           =   0
               Left            =   -15
               TabIndex        =   180
               Top             =   0
               Width           =   3990
            End
            Begin VB.OptionButton rbcSchOrAuto 
               Caption         =   "Generate Schedule after Automation Completed"
               Height          =   195
               Index           =   1
               Left            =   3990
               TabIndex        =   179
               Top             =   0
               Width           =   3930
            End
            Begin VB.OptionButton rbcSchOrAuto 
               Caption         =   "Generate them Independently"
               Height          =   195
               Index           =   2
               Left            =   7920
               TabIndex        =   178
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.Frame frcPurge 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   75
            TabIndex        =   171
            Top             =   915
            Width           =   10695
            Begin VB.OptionButton rbcPurge 
               Caption         =   "After Automation Completed"
               Height          =   195
               Index           =   1
               Left            =   5640
               TabIndex        =   175
               Top             =   0
               Width           =   2430
            End
            Begin VB.OptionButton rbcPurge 
               Caption         =   "After Schedule Completed"
               Height          =   195
               Index           =   0
               Left            =   3240
               TabIndex        =   174
               Top             =   0
               Width           =   2355
            End
            Begin VB.OptionButton rbcPurge 
               Caption         =   "At"
               Height          =   195
               Index           =   2
               Left            =   8250
               TabIndex        =   173
               Top             =   0
               Width           =   690
            End
            Begin VB.TextBox edcPurgeTime 
               Height          =   285
               Left            =   8970
               MaxLength       =   10
               TabIndex        =   172
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label lacPurge 
               Caption         =   "Purge 'Schedule' and 'Library'"
               Height          =   225
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Width           =   2970
            End
         End
         Begin VB.Label lacAutoGenTime 
            Caption         =   "Automation Generation Time:"
            Height          =   255
            Left            =   4095
            TabIndex        =   184
            Top             =   525
            Width           =   2400
         End
         Begin VB.Label lacSchdGenTime 
            Caption         =   "Schedule Generation Time:"
            Height          =   255
            Left            =   60
            TabIndex        =   182
            Top             =   525
            Width           =   2130
         End
      End
      Begin VB.Frame frcAutomation 
         Caption         =   "Automation Creation"
         Height          =   2670
         Left            =   6615
         TabIndex        =   115
         Top             =   120
         Width           =   3915
         Begin VB.TextBox edcAutoGrid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1365
            TabIndex        =   119
            Top             =   630
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.PictureBox pbcAutoArrow 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   105
            Picture         =   "EngrSiteOption.frx":030A
            ScaleHeight     =   165
            ScaleWidth      =   90
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox pbcAutoSTab 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   90
            Left            =   195
            ScaleHeight     =   90
            ScaleWidth      =   60
            TabIndex        =   117
            Top             =   270
            Width           =   60
         End
         Begin VB.PictureBox pbcAutoTab 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H8000000F&
            Height          =   90
            Left            =   120
            ScaleHeight     =   90
            ScaleWidth      =   60
            TabIndex        =   120
            Top             =   2490
            Width           =   60
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAuto 
            Height          =   2085
            Left            =   330
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   345
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   3678
            _Version        =   393216
            Rows            =   8
            Cols            =   3
            FixedCols       =   0
            ForeColorFixed  =   -2147483640
            BackColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorUnpopulated=   -2147483634
            AllowBigSelection=   0   'False
            ScrollBars      =   0
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame frcSchd 
         Caption         =   "Schedule Creation"
         Height          =   2685
         Left            =   90
         TabIndex        =   103
         Top             =   120
         Width           =   6315
         Begin VB.TextBox edcSchdGrid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1350
            TabIndex        =   107
            Top             =   645
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.PictureBox pbcSchdArrow 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   180
            Picture         =   "EngrSiteOption.frx":0614
            ScaleHeight     =   165
            ScaleWidth      =   90
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox pbcSchdSTab 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   90
            Left            =   270
            ScaleHeight     =   90
            ScaleWidth      =   60
            TabIndex        =   105
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox pbcSchdTab 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H8000000F&
            Height          =   90
            Left            =   120
            ScaleHeight     =   90
            ScaleWidth      =   60
            TabIndex        =   108
            Top             =   2430
            Width           =   60
         End
         Begin VB.TextBox edcMinEventID 
            Height          =   285
            Left            =   3735
            MaxLength       =   9
            TabIndex        =   110
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox edcMaxEventID 
            Height          =   285
            Left            =   3735
            MaxLength       =   9
            TabIndex        =   112
            Top             =   780
            Width           =   1005
         End
         Begin VB.TextBox edcCurrEventID 
            Height          =   285
            Left            =   3735
            MaxLength       =   9
            TabIndex        =   114
            Top             =   1230
            Width           =   1005
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSchd 
            Height          =   2085
            Left            =   330
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   3678
            _Version        =   393216
            Rows            =   8
            Cols            =   3
            FixedCols       =   0
            ForeColorFixed  =   -2147483640
            BackColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorUnpopulated=   -2147483634
            AllowBigSelection=   0   'False
            ScrollBars      =   0
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lacEventID 
            Caption         =   "Event ID-  Min:"
            Height          =   255
            Left            =   2460
            TabIndex        =   109
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label lacMaxEventID 
            Caption         =   "Max:"
            Height          =   255
            Left            =   3195
            TabIndex        =   111
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lacCurrEventID 
            Caption         =   "Current:"
            Height          =   255
            Left            =   3030
            TabIndex        =   113
            Top             =   1230
            Width           =   1005
         End
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Options"
      Height          =   5145
      Index           =   1
      Left            =   9915
      TabIndex        =   13
      Top             =   4050
      Visible         =   0   'False
      Width           =   10380
      Begin VB.CheckBox ckcMatchBNotT 
         Caption         =   "Buses match, Start/End Times either overlap or are the same"
         Height          =   285
         Left            =   510
         TabIndex        =   164
         Top             =   2520
         Width           =   4995
      End
      Begin VB.CheckBox ckcMatchANotT 
         Caption         =   "Audio Source match, Start/End Times don't match but do overlap"
         Height          =   285
         Left            =   510
         TabIndex        =   163
         Top             =   2220
         Width           =   5535
      End
      Begin VB.CheckBox ckcMatchATBNotI 
         Caption         =   "Audio Source match, Start/End Times match, Buses match, Item ID's don't match"
         Height          =   285
         Left            =   510
         TabIndex        =   162
         Top             =   1920
         Width           =   6180
      End
      Begin VB.CheckBox ckcMatchATNotB 
         Caption         =   "Audio Source match, Start/End Times match, Buses don't match"
         Height          =   285
         Left            =   510
         TabIndex        =   160
         Top             =   1620
         Width           =   5040
      End
      Begin VB.TextBox edcLengthTolerance 
         Height          =   285
         Left            =   10185
         MaxLength       =   7
         TabIndex        =   26
         Top             =   4920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox edcTimeTolerance 
         Height          =   285
         Left            =   10185
         MaxLength       =   7
         TabIndex        =   24
         Top             =   4425
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox edcAlertInterval 
         Height          =   285
         Left            =   2115
         MaxLength       =   5
         TabIndex        =   28
         Top             =   4230
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox edcChgInterval 
         Height          =   285
         Left            =   2115
         MaxLength       =   10
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox edcRetainActivityLog 
         Height          =   285
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   18
         Top             =   4680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox edcRetainSchd 
         Height          =   285
         Left            =   3435
         MaxLength       =   5
         TabIndex        =   15
         Top             =   345
         Width           =   750
      End
      Begin VB.Label lacConflicttest 
         Caption         =   "Conflict Test to be performed:"
         Height          =   255
         Left            =   90
         TabIndex        =   161
         Top             =   1395
         Width           =   2265
      End
      Begin VB.Label lacLengthTolerance 
         Caption         =   "'As Aired' Length Tolerance (mm:ss.t):"
         Height          =   255
         Left            =   7230
         TabIndex        =   25
         Top             =   4920
         Visible         =   0   'False
         Width           =   2910
      End
      Begin VB.Label lacTimeTolerance 
         Caption         =   "'As Aired' Time Tolerance (mm:ss.t):"
         Height          =   255
         Left            =   7230
         TabIndex        =   23
         Top             =   4425
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label lacAlertInterval 
         Caption         =   "Alert Interval in Minutes:"
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   4230
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lacActivityLogDays 
         Caption         =   "Days"
         Height          =   255
         Left            =   6420
         TabIndex        =   19
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label laclacChgsTime 
         Caption         =   "of 'Schedule/Automation' on Today's Date"
         Height          =   240
         Left            =   3615
         TabIndex        =   22
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lacChgs 
         Caption         =   "Allow Changes within:"
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   840
         Width           =   2340
      End
      Begin VB.Label lacActivityLog 
         Caption         =   "Retain 'Activity Log' File for:"
         Height          =   255
         Left            =   2820
         TabIndex        =   17
         Top             =   4680
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lacRetainSchdDays 
         Caption         =   "Days"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   345
         Width           =   615
      End
      Begin VB.Label lacRetainSchd 
         Caption         =   "Retain 'Schedule' and 'Library'  Files for:"
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   345
         Width           =   3375
      End
   End
   Begin VB.Frame frcIDtest 
      Caption         =   "Item ID Test"
      Height          =   5580
      Left            =   10170
      TabIndex        =   29
      Top             =   3630
      Visible         =   0   'False
      Width           =   11385
      Begin VB.TextBox edcSpotItemIDWindow 
         Height          =   285
         Left            =   6045
         MaxLength       =   5
         TabIndex        =   158
         Top             =   4680
         Width           =   825
      End
      Begin VB.Frame frcSecondary 
         Caption         =   "Protection Item ID System"
         Height          =   4425
         Left            =   5805
         TabIndex        =   66
         Top             =   180
         Width           =   5520
         Begin VB.CheckBox ckcDASSSecCheckSum 
            Caption         =   "Check Sum (CRC)"
            Height          =   240
            Left            =   3015
            TabIndex        =   97
            Top             =   3540
            Width           =   2160
         End
         Begin VB.TextBox edcDASSSecMgsErrType 
            Height          =   285
            Left            =   1860
            MaxLength       =   2
            TabIndex        =   96
            Top             =   3510
            Width           =   600
         End
         Begin VB.TextBox edcDASSSecLengthID 
            Height          =   285
            Left            =   4785
            MaxLength       =   2
            TabIndex        =   94
            Top             =   3075
            Width           =   600
         End
         Begin VB.TextBox edcDASSSecTitleID 
            Height          =   285
            Left            =   3030
            MaxLength       =   2
            TabIndex        =   92
            Top             =   3060
            Width           =   600
         End
         Begin VB.TextBox edcDASSSecConnectSeq 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   78
            Top             =   1635
            Width           =   1185
         End
         Begin VB.TextBox edcDASSSecMessEnd 
            Height          =   285
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   99
            Top             =   3975
            Width           =   825
         End
         Begin VB.TextBox edcDASSSecName 
            Height          =   285
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   68
            Top             =   300
            Width           =   3270
         End
         Begin VB.TextBox edcDASSSecDataBits 
            Height          =   285
            Left            =   1125
            MaxLength       =   2
            TabIndex        =   70
            Top             =   735
            Width           =   825
         End
         Begin VB.ComboBox cbcDASSSecParity 
            Height          =   315
            ItemData        =   "EngrSiteOption.frx":091E
            Left            =   2835
            List            =   "EngrSiteOption.frx":0920
            TabIndex        =   72
            Top             =   735
            Width           =   1065
         End
         Begin VB.ComboBox cbcDASSSecStopBit 
            Height          =   315
            ItemData        =   "EngrSiteOption.frx":0922
            Left            =   675
            List            =   "EngrSiteOption.frx":0924
            TabIndex        =   74
            Top             =   1170
            Width           =   1065
         End
         Begin VB.TextBox edcDASSSecBaud 
            Height          =   285
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   76
            Top             =   1170
            Width           =   825
         End
         Begin VB.TextBox edcDASSSecMachID 
            Height          =   285
            Left            =   4125
            MaxLength       =   2
            TabIndex        =   82
            Top             =   2085
            Width           =   510
         End
         Begin VB.TextBox edcDASSSecStartChar 
            Height          =   285
            Left            =   1455
            MaxLength       =   1
            TabIndex        =   80
            Top             =   2070
            Width           =   825
         End
         Begin VB.TextBox edcDASSSecReplyChar 
            Height          =   285
            Left            =   4425
            MaxLength       =   1
            TabIndex        =   101
            Top             =   3975
            Width           =   825
         End
         Begin VB.TextBox edcDASSSecMinID 
            Height          =   285
            Left            =   1500
            MaxLength       =   5
            TabIndex        =   84
            Top             =   2565
            Width           =   705
         End
         Begin VB.TextBox edcDASSSecMaxID 
            Height          =   285
            Left            =   3000
            MaxLength       =   5
            TabIndex        =   86
            Top             =   2565
            Width           =   690
         End
         Begin VB.TextBox edcDASSSecCurID 
            Height          =   285
            Left            =   4620
            MaxLength       =   5
            TabIndex        =   88
            Top             =   2565
            Width           =   705
         End
         Begin VB.TextBox edcDASSSecMessType 
            Height          =   285
            Left            =   1350
            MaxLength       =   2
            TabIndex        =   90
            Top             =   3060
            Width           =   825
         End
         Begin VB.Label lacDASSSecMgsErrType 
            Caption         =   "Message Error Type:"
            Height          =   255
            Left            =   90
            TabIndex        =   95
            Top             =   3510
            Width           =   1590
         End
         Begin VB.Label lacDASSSecLengthID 
            Caption         =   "Length ID:"
            Height          =   255
            Left            =   3900
            TabIndex        =   93
            Top             =   3075
            Width           =   795
         End
         Begin VB.Label lacDASSSecTitleID 
            Caption         =   "Title ID:"
            Height          =   255
            Left            =   2325
            TabIndex        =   91
            Top             =   3060
            Width           =   795
         End
         Begin VB.Label lacDASSSecConnectSeq 
            Caption         =   "Connect Sequence:"
            Height          =   255
            Left            =   90
            TabIndex        =   77
            Top             =   1635
            Width           =   1620
         End
         Begin VB.Label lacDASSSecMessEnd 
            Caption         =   "Message End Character:"
            Height          =   255
            Left            =   90
            TabIndex        =   98
            Top             =   3975
            Width           =   1950
         End
         Begin VB.Label lacDASSSecName 
            Caption         =   "Name:"
            Height          =   255
            Left            =   90
            TabIndex        =   67
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label lacDASSSecDataBits 
            Caption         =   "Data Bits:"
            Height          =   255
            Left            =   90
            TabIndex        =   69
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label lacDASSSecParity 
            Caption         =   "Parity:"
            Height          =   255
            Left            =   2250
            TabIndex        =   71
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label lacDASSSecStopBit 
            Caption         =   "Stop:"
            Height          =   255
            Left            =   90
            TabIndex        =   73
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label lacDASSSecBaud 
            Caption         =   "Baud:"
            Height          =   255
            Left            =   2010
            TabIndex        =   75
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label lacDASSSecMachID 
            Caption         =   "Machine ID:"
            Height          =   255
            Left            =   3090
            TabIndex        =   81
            Top             =   2085
            Width           =   1275
         End
         Begin VB.Label lacDASSSecStartChar 
            Caption         =   "Start Character:"
            Height          =   255
            Left            =   90
            TabIndex        =   79
            Top             =   2085
            Width           =   1905
         End
         Begin VB.Label lacDASSSecReplyChar 
            Caption         =   "Reply Character:"
            Height          =   255
            Left            =   3000
            TabIndex        =   100
            Top             =   3975
            Width           =   2040
         End
         Begin VB.Label lacDASSSecMinID 
            Caption         =   "Message ID-  Min:"
            Height          =   255
            Left            =   90
            TabIndex        =   83
            Top             =   2565
            Width           =   1440
         End
         Begin VB.Label lacDASSSecMaxID 
            Caption         =   "Max:"
            Height          =   255
            Left            =   2475
            TabIndex        =   85
            Top             =   2565
            Width           =   585
         End
         Begin VB.Label lacDASSSecCurID 
            Caption         =   "Current:"
            Height          =   255
            Left            =   3930
            TabIndex        =   87
            Top             =   2565
            Width           =   1005
         End
         Begin VB.Label lacDASSSecMessType 
            Caption         =   "Message Type:"
            Height          =   255
            Left            =   90
            TabIndex        =   89
            Top             =   3060
            Width           =   1440
         End
      End
      Begin VB.Frame frcPrimary 
         Caption         =   "Primary Item ID System"
         Height          =   4425
         Left            =   105
         TabIndex        =   30
         Top             =   180
         Width           =   5520
         Begin VB.CheckBox ckcDASSPriCheckSum 
            Caption         =   "Check Sum (CRC)"
            Height          =   240
            Left            =   3015
            TabIndex        =   61
            Top             =   3540
            Width           =   2160
         End
         Begin VB.TextBox edcDASSPriMgsErrType 
            Height          =   285
            Left            =   1860
            MaxLength       =   2
            TabIndex        =   60
            Top             =   3510
            Width           =   600
         End
         Begin VB.TextBox edcDASSPriLengthID 
            Height          =   285
            Left            =   4785
            MaxLength       =   2
            TabIndex        =   58
            Top             =   3075
            Width           =   600
         End
         Begin VB.TextBox edcDASSPriTitleID 
            Height          =   285
            Left            =   3030
            MaxLength       =   2
            TabIndex        =   56
            Top             =   3060
            Width           =   600
         End
         Begin VB.TextBox edcDASSPriConnectSeq 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   42
            Top             =   1635
            Width           =   1185
         End
         Begin VB.TextBox edcDASSPriMessEnd 
            Height          =   285
            Left            =   1935
            MaxLength       =   10
            TabIndex        =   63
            Top             =   3975
            Width           =   825
         End
         Begin VB.TextBox edcDASSPriMessType 
            Height          =   285
            Left            =   1350
            MaxLength       =   2
            TabIndex        =   54
            Top             =   3060
            Width           =   600
         End
         Begin VB.TextBox edcDASSPriCurID 
            Height          =   285
            Left            =   4515
            MaxLength       =   5
            TabIndex        =   52
            Top             =   2565
            Width           =   705
         End
         Begin VB.TextBox edcDASSPriMaxID 
            Height          =   285
            Left            =   2940
            MaxLength       =   5
            TabIndex        =   50
            Top             =   2565
            Width           =   705
         End
         Begin VB.TextBox edcDASSPriMinID 
            Height          =   285
            Left            =   1485
            MaxLength       =   5
            TabIndex        =   48
            Top             =   2565
            Width           =   705
         End
         Begin VB.TextBox edcDASSPriReplyChar 
            Height          =   285
            Left            =   4425
            MaxLength       =   1
            TabIndex        =   65
            Top             =   3975
            Width           =   825
         End
         Begin VB.TextBox edcDASSPriStartChar 
            Height          =   285
            Left            =   1455
            MaxLength       =   1
            TabIndex        =   44
            Top             =   2085
            Width           =   825
         End
         Begin VB.TextBox edcDASSPriMachID 
            Height          =   285
            Left            =   4110
            MaxLength       =   2
            TabIndex        =   46
            Top             =   2085
            Width           =   510
         End
         Begin VB.TextBox edcDASSPriBaud 
            Height          =   285
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   40
            Top             =   1170
            Width           =   825
         End
         Begin VB.ComboBox cbcDASSPriStopBit 
            Height          =   315
            ItemData        =   "EngrSiteOption.frx":0926
            Left            =   675
            List            =   "EngrSiteOption.frx":0928
            TabIndex        =   38
            Top             =   1170
            Width           =   1065
         End
         Begin VB.ComboBox cbcDASSPriParity 
            Height          =   315
            ItemData        =   "EngrSiteOption.frx":092A
            Left            =   2835
            List            =   "EngrSiteOption.frx":092C
            TabIndex        =   36
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox edcDASSPriDataBits 
            Height          =   285
            Left            =   1125
            MaxLength       =   2
            TabIndex        =   34
            Top             =   735
            Width           =   825
         End
         Begin VB.TextBox edcDASSPriName 
            Height          =   285
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   32
            Top             =   300
            Width           =   3270
         End
         Begin VB.Label acDASSPriMgsErrType 
            Caption         =   "Message Error Type:"
            Height          =   255
            Left            =   90
            TabIndex        =   59
            Top             =   3510
            Width           =   1590
         End
         Begin VB.Label lacDASSPriLengthID 
            Caption         =   "Length ID:"
            Height          =   255
            Left            =   3900
            TabIndex        =   57
            Top             =   3075
            Width           =   795
         End
         Begin VB.Label lacDASSPriTitleID 
            Caption         =   "Title ID:"
            Height          =   255
            Left            =   2235
            TabIndex        =   55
            Top             =   3060
            Width           =   795
         End
         Begin VB.Label lacDASSPriSConnectSeq 
            Caption         =   "Connect Sequence:"
            Height          =   255
            Left            =   90
            TabIndex        =   41
            Top             =   1635
            Width           =   1620
         End
         Begin VB.Label lacDASSPriMessEnd 
            Caption         =   "Message End Character:"
            Height          =   255
            Left            =   105
            TabIndex        =   62
            Top             =   3975
            Width           =   1935
         End
         Begin VB.Label lacDASSPriMessType 
            Caption         =   "Message Type:"
            Height          =   255
            Left            =   90
            TabIndex        =   53
            Top             =   3060
            Width           =   1440
         End
         Begin VB.Label lacDASSPriCurID 
            Caption         =   "Current:"
            Height          =   255
            Left            =   3855
            TabIndex        =   51
            Top             =   2565
            Width           =   1005
         End
         Begin VB.Label lacDASSPriMaxID 
            Caption         =   "Max:"
            Height          =   255
            Left            =   2415
            TabIndex        =   49
            Top             =   2565
            Width           =   585
         End
         Begin VB.Label lacDASSPriMinID 
            Caption         =   "Message ID-  Min:"
            Height          =   255
            Left            =   90
            TabIndex        =   47
            Top             =   2565
            Width           =   1425
         End
         Begin VB.Label lacDASSPriReplyChar 
            Caption         =   "Reply Character:"
            Height          =   255
            Left            =   3000
            TabIndex        =   64
            Top             =   3975
            Width           =   2040
         End
         Begin VB.Label lacDASSPriStartChar 
            Caption         =   "Start Character:"
            Height          =   255
            Left            =   90
            TabIndex        =   43
            Top             =   2085
            Width           =   1905
         End
         Begin VB.Label lacDASSPriMachID 
            Caption         =   "Machine ID:"
            Height          =   255
            Left            =   3075
            TabIndex        =   45
            Top             =   2085
            Width           =   1275
         End
         Begin VB.Label lacDASSPriBaud 
            Caption         =   "Baud:"
            Height          =   255
            Left            =   2010
            TabIndex        =   39
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label lacDASSPriStopBit 
            Caption         =   "Stop:"
            Height          =   255
            Left            =   90
            TabIndex        =   37
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label lacDASSPriParity 
            Caption         =   "Parity:"
            Height          =   255
            Left            =   2250
            TabIndex        =   35
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label lacDASSPriDataBits 
            Caption         =   "Data Bits:"
            Height          =   255
            Left            =   90
            TabIndex        =   33
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label lacDASSPriName 
            Caption         =   "Name:"
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Label lacSpotItemIDWindow 
         Caption         =   "Spot Item ID Length Check Window (in milliseconds, 1000 = plus/minus 1 second):"
         Height          =   255
         Left            =   135
         TabIndex        =   159
         Top             =   4680
         Width           =   6045
      End
      Begin VB.Label lacForm 
         Caption         =   $"EngrSiteOption.frx":092E
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   157
         Top             =   5310
         Width           =   11160
      End
      Begin VB.Label lacForm 
         Caption         =   "Request Seq:   Start Char     Mach ID    Message ID    Mess Type    Item ID    Title or Length ID    CRC    End Char"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   156
         Top             =   5055
         Width           =   11160
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "General"
      Height          =   3000
      Index           =   0
      Left            =   10380
      TabIndex        =   2
      Top             =   3240
      Width           =   9510
      Begin VB.TextBox edcAddress 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1605
         Width           =   5940
      End
      Begin VB.TextBox edcAddress 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1185
         Width           =   5940
      End
      Begin VB.TextBox edcFax 
         Height          =   285
         Left            =   6630
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2055
         Width           =   2475
      End
      Begin VB.TextBox edcPhone 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2055
         Width           =   2835
      End
      Begin VB.TextBox edcName 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   4
         Top             =   345
         Width           =   4530
      End
      Begin VB.TextBox edcAddress 
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   6
         Top             =   765
         Width           =   5940
      End
      Begin VB.Label lacFax 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   5385
         TabIndex        =   11
         Top             =   2055
         Width           =   1035
      End
      Begin VB.Label lacPhone 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   2055
         Width           =   1440
      End
      Begin VB.Label lacName 
         Caption         =   "Client Name:"
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   345
         Width           =   1275
      End
      Begin VB.Label lacAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   765
         Width           =   1380
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   570
      Top             =   6420
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6840
      FormDesignWidth =   11565
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5130
      TabIndex        =   153
      Top             =   6390
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3480
      TabIndex        =   152
      Top             =   6390
      Width           =   1335
   End
   Begin ComctlLib.TabStrip tabAuto 
      Height          =   6045
      Left            =   30
      TabIndex        =   1
      Top             =   255
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   10663
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Options"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Schedule && Automation Creation"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Commercial Merge"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&E-Mail"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6735
      TabIndex        =   154
      Top             =   6390
      Width           =   1335
   End
   Begin VB.Label lacScreen 
      Caption         =   "Site Option"
      Height          =   270
      Left            =   0
      TabIndex        =   155
      Top             =   0
      Width           =   2625
   End
End
Attribute VB_Name = "EngrSiteOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrSiteOption - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imBSMode As Integer
Private imInChg As Integer
Private imSoeCode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String

Private tmSOE As SOE
Private smITEStamp As String
Private tmITE() As ITE
Private smSGEStamp As String
Private tmSGE() As SGE
Private smSPEStamp As String
Private tmSPE() As SPE
Private tmSSE As SSE

Private imTabIndex As Integer

'Grid Controls
Private imSchdFromArrow As Integer
Private imShowGridBox As Integer
Private lmSchdEnableRow As Long         'Current or last row focus was on
Private lmSchdEnableCol As Long         'Current or last column focus was on
Private imAutoFromArrow As Integer
Private lmAutoEnableRow As Long         'Current or last row focus was on
Private lmAutoEnableCol As Long         'Current or last column focus was on


Const LOGDAYINDEX = 0
Const LEADDAYINDEX = 1
Const CODEINDEX = 2

Private Function mCheckFields(ilShowMsg As Integer) As Integer
    Dim slStr As String
    
    mCheckFields = True
    slStr = Trim$(edcName.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Client Names must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcChgInterval.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Allow Changes Within Schedule/Automation on Today must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcMinEventID.text)
    If (slStr = "") Or (Val(slStr) = 0) Then
        If ilShowMsg Then
            MsgBox "Schedule Min ID must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcMaxEventID.text)
    If (slStr = "") Or (Val(slStr) = 0) Then
        If ilShowMsg Then
            MsgBox "Schedule Max ID must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcMinEventID.text)
    If (slStr <> "") And (Val(slStr) <> 0) Then
        slStr = Trim$(edcMaxEventID.text)
        If (slStr <> "") And (Val(slStr) <> 0) Then
            If Val(slStr) < Val(Trim$(edcMinEventID)) Then
                If ilShowMsg Then
                    MsgBox "Schedule Min ID must be less then Max ID", vbCritical + vbOKOnly, "Save not Completed"
                    'edcName.SetFocus
                End If
                mCheckFields = False
            End If
        End If
    End If
    slStr = Trim$(edcCurrEventID.text)
    If (slStr = "") Or (Val(slStr) = 0) Then
        If ilShowMsg Then
            MsgBox "Schedule Current ID must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcMinEventID.text)
    If (slStr <> "") And (Val(slStr) <> 0) Then
        slStr = Trim$(edcMaxEventID.text)
        If (slStr <> "") And (Val(slStr) <> 0) Then
            slStr = Trim$(edcCurrEventID.text)
            If (slStr = "") Or (Val(slStr) = 0) Then
                If (Val(slStr) < Val(Trim$(edcMinEventID))) Or (Val(slStr) > Val(Trim$(edcMaxEventID))) Then
                    If ilShowMsg Then
                        MsgBox "Schedule Current ID must be between Min ID and Max ID", vbCritical + vbOKOnly, "Save not Completed"
                        'edcName.SetFocus
                    End If
                    mCheckFields = False
                End If
            End If
        End If
    End If
    If Val(edcMergeChkInterval.text) > 0 Then
        slStr = Trim$(edcImportFileFormat.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Merge File Format must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If InStr(1, slStr, "Date", vbTextCompare) > 0 Then
                slStr = Trim$(edcDateFormat.text)
                If slStr = "" Then
                    If ilShowMsg Then
                        MsgBox "Merge File Date Format must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                        'edcName.SetFocus
                    End If
                    mCheckFields = False
                End If
            End If
            slStr = Trim$(edcImportFileFormat.text)
            If InStr(1, slStr, "Time", vbTextCompare) > 0 Then
                slStr = Trim$(edcTimeFormat.text)
                If slStr = "" Then
                    If ilShowMsg Then
                        MsgBox "Merge File Time Format must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                        'edcName.SetFocus
                    End If
                    mCheckFields = False
                End If
            End If
        End If
        slStr = Trim$(edcImportExt.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Merge File Extension must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        End If
        slStr = Trim$(edcMergeChkStartTime.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Commercial Merge Check Start Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Commercial Merge Check Start Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
        slStr = Trim$(edcMergeChkEnd.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Commercial Merge Check End Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Commercial Merge Check End Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If
    If rbcPurge(2).Value = True Then
        slStr = Trim$(edcPurgeTime.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Schedule Purge Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Schedule Purge Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If
    If rbcSchOrAuto(0).Value Or rbcSchOrAuto(2).Value Then
        slStr = Trim$(edcSchdGenTime.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Schedule Generation Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Schedule Generation Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If
    If rbcSchOrAuto(1).Value Or rbcSchOrAuto(2).Value Then
        slStr = Trim$(edcAutoGenTime.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Automation Generation Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Automation Generation Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If

    If rbcPurgeTest(2).Value = True Then
        slStr = Trim$(edcPurgeTimeTest.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Test Schedule Purge Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Test Schedule Purge Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If
    If rbcSchOrAutoTest(0).Value Or rbcSchOrAutoTest(2).Value Then
        slStr = Trim$(edcSchdGenTimeTest.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Test Schedule Generation Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Test Schedule Generation Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If
    If rbcSchOrAutoTest(1).Value Or rbcSchOrAutoTest(2).Value Then
        slStr = Trim$(edcAutoGenTime.text)
        If slStr = "" Then
            If ilShowMsg Then
                MsgBox "Test Automation Generation Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
                'edcName.SetFocus
            End If
            mCheckFields = False
        Else
            If Not gIsTime(slStr) Then
                If ilShowMsg Then
                    MsgBox "Test Automation Generation Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
                End If
                mCheckFields = False
            End If
        End If
    End If

    slStr = Trim$(edcChgInterval.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Change Interval Time must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            'edcName.SetFocus
        End If
        mCheckFields = False
    Else
        If Not gIsTimeTenths(slStr) Then
            If ilShowMsg Then
                MsgBox "Change Interval Time not valid format", vbCritical + vbOKOnly, "Save not Completed"
            End If
            mCheckFields = False
        End If
    End If
End Function

Private Sub mInit()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    imInChg = True
    imTabIndex = 1
    imSoeCode = tgSOE.iCode
    mClearControls
    cbcDASSPriParity.AddItem "Even"
    cbcDASSPriParity.ItemData(cbcDASSPriParity.NewIndex) = 0
    cbcDASSPriParity.AddItem "Mark"
    cbcDASSPriParity.ItemData(cbcDASSPriParity.NewIndex) = 1
    cbcDASSPriParity.AddItem "None"
    cbcDASSPriParity.ItemData(cbcDASSPriParity.NewIndex) = 2
    cbcDASSPriParity.AddItem "Odd"
    cbcDASSPriParity.ItemData(cbcDASSPriParity.NewIndex) = 3
    cbcDASSPriParity.AddItem "Space"
    cbcDASSPriParity.ItemData(cbcDASSPriParity.NewIndex) = 4
    cbcDASSSecParity.AddItem "Even"
    cbcDASSSecParity.ItemData(cbcDASSSecParity.NewIndex) = 0
    cbcDASSSecParity.AddItem "Mark"
    cbcDASSSecParity.ItemData(cbcDASSSecParity.NewIndex) = 1
    cbcDASSSecParity.AddItem "None"
    cbcDASSSecParity.ItemData(cbcDASSSecParity.NewIndex) = 2
    cbcDASSSecParity.AddItem "Odd"
    cbcDASSSecParity.ItemData(cbcDASSSecParity.NewIndex) = 3
    cbcDASSSecParity.AddItem "Space"
    cbcDASSSecParity.ItemData(cbcDASSSecParity.NewIndex) = 4
    
    cbcDASSPriStopBit.AddItem "1"
    cbcDASSPriStopBit.ItemData(cbcDASSPriStopBit.NewIndex) = 1
    cbcDASSPriStopBit.AddItem "2"
    cbcDASSPriStopBit.ItemData(cbcDASSPriStopBit.NewIndex) = 2
    cbcDASSSecStopBit.AddItem "1"
    cbcDASSSecStopBit.ItemData(cbcDASSSecStopBit.NewIndex) = 1
    cbcDASSSecStopBit.AddItem "2"
    cbcDASSSecStopBit.ItemData(cbcDASSSecStopBit.NewIndex) = 2
    ilRet = gGetRec_SOE_SiteOption(tgSOE.iCode, "Site Option-mInit: Get Record", tmSOE)
    ilRet = gGetRecs_ITE_ItemTest(smITEStamp, tgSOE.iCode, "Site Option-mInit: Get ITE", tmITE())
    ilRet = gGetRecs_SGE_SiteGenSchd(smSGEStamp, tgSOE.iCode, "Site Option-mInit: Get SGE", tmSGE())
    ilRet = gGetRecs_SPE_SitePath(smSPEStamp, tgSOE.iCode, "Site Option-mInit: Get SPE", tmSPE())
    ilRet = gGetRec_SSE_Site_SMTP_Info(tgSOE.iCode, "Site Option-mInit: Get SSE", tmSSE)
    mMoveRecToCtrls
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(SITELIST) <> 2) Then
        igPasswordOk = False
    ElseIf (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Then
        igPasswordOk = True
    Else
'        Screen.MousePointer = vbDefault
'        EngrPass.Show vbModal
'        Screen.MousePointer = vbDefault
        igPasswordOk = True
    End If
    If Not igPasswordOk Then
        cmcDone.Enabled = False
        For ilLoop = frcTab.LBound To frcTab.UBound Step 1
            frcTab(ilLoop).Enabled = False
        Next ilLoop
    Else
        cmcDone.Enabled = True
        For ilLoop = frcTab.LBound To frcTab.UBound Step 1
            frcTab(ilLoop).Enabled = True
        Next ilLoop
    End If
    
    imInChg = False
    imFieldChgd = False
End Sub

Private Sub mMoveCtrlsToRec()
    Dim ilLoop As Integer
    Dim llRow As Long
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    tmSOE.iCode = imSoeCode
    tmSOE.sClientName = edcName.text
    tmSOE.sAddr1 = edcAddress(0).text
    tmSOE.sAddr2 = edcAddress(1).text
    tmSOE.sAddr3 = edcAddress(2).text
    tmSOE.sPhone = edcPhone.text
    tmSOE.sFax = edcFax.text
    tmSOE.iDaysRetainAsAir = Val(edcRetainSchd.text)
    tmSOE.iDaysRetainActive = Val(edcRetainActivityLog.text)
    If Trim$(edcChgInterval.text) = "" Then
        tmSOE.lChgInterval = gStrTimeInTenthToLong("00:00:00.0", False)
    Else
        tmSOE.lChgInterval = gStrTimeInTenthToLong(edcChgInterval.text, False)
    End If
    tmSOE.sMergeDateFormat = edcDateFormat.text
    tmSOE.sMergeTimeFormat = edcTimeFormat.text
    tmSOE.sMergeFileFormat = edcImportFileFormat.text
    tmSOE.sMergeFileExt = edcImportExt.text
    If Trim$(edcMergeChkStartTime.text) = "" Then
        tmSOE.sMergeStartTime = Format("00:00:00", sgShowTimeWSecForm)
    Else
        tmSOE.sMergeStartTime = gFormatTime(edcMergeChkStartTime.text)
    End If
    If Trim$(edcMergeChkEnd.text) = "" Then
        tmSOE.sMergeEndTime = Format("00:00:00", sgShowTimeWSecForm)
    Else
        tmSOE.sMergeEndTime = gFormatTime(edcMergeChkEnd.text)
    End If
    If rbcMerge(0).Value Then
        tmSOE.sMergeStopFlag = "Y"
    Else
        tmSOE.sMergeStopFlag = "N"
    End If
    tmSOE.iMergeChkInterval = Val(edcMergeChkInterval.text)
    tmSOE.iAlertInterval = Val(edcAlertInterval.text)
    tmSOE.lTimeTolerance = gStrLengthInTenthToLong(edcTimeTolerance.text)
    tmSOE.lLengthTolerance = gStrLengthInTenthToLong(edcLengthTolerance.text)
    tmSOE.sSchAutoGenSeq = "I"
    tmSOE.sSchAutoGenSeqTst = "I"
    tmSOE.lMinEventID = Val(edcMinEventID.text)
    tmSOE.lMaxEventID = Val(edcMaxEventID.text)
    tmSOE.lCurrEventID = Val(edcCurrEventID.text)
    If rbcSchOrAuto(0).Value Then
        tmSOE.sSchAutoGenSeq = "S"
    ElseIf rbcSchOrAuto(1).Value Then
        tmSOE.sSchAutoGenSeq = "A"
    Else
        tmSOE.sSchAutoGenSeq = "I"
    End If
    If rbcSchOrAutoTest(0).Value Then
        tmSOE.sSchAutoGenSeqTst = "S"
    ElseIf rbcSchOrAutoTest(1).Value Then
        tmSOE.sSchAutoGenSeqTst = "A"
    Else
        tmSOE.sSchAutoGenSeqTst = "I"
    End If
    If rbcMergeTest(0).Value Then
        tmSOE.sMergeStopFlagTst = "Y"
    Else
        tmSOE.sMergeStopFlagTst = "N"
    End If
    tmSOE.iVersion = imVersion + 1
    tmSOE.iOrigSoeCode = imSoeCode
    tmSOE.sCurrent = "Y"
    'tmSOE.sEnteredDate = smNowDate
    'tmSOE.sEnteredTime = smNowTime
    tmSOE.sEnteredDate = Format(Now, sgShowDateForm)
    tmSOE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmSOE.iUieCode = tgUIE.iCode
    tmSOE.iSpotItemIDWindow = Val(edcSpotItemIDWindow.text)
    If ckcMatchATNotB.Value = vbUnchecked Then
        tmSOE.sMatchATNotB = "N"
    Else
        tmSOE.sMatchATNotB = "Y"
    End If
    
    If ckcMatchATBNotI.Value = vbUnchecked Then
        tmSOE.sMatchATBNotI = "N"
    Else
        tmSOE.sMatchATBNotI = "Y"
    End If
    If ckcMatchANotT.Value = vbUnchecked Then
        tmSOE.sMatchANotT = "N"
    Else
        tmSOE.sMatchANotT = "Y"
    End If
    If ckcMatchBNotT.Value = vbUnchecked Then
        tmSOE.sMatchBNotT = "N"
    Else
        tmSOE.sMatchBNotT = "Y"
    End If

    tmSOE.sUnused = ""
    ReDim tmITE(0 To 2) As ITE
    ReDim tmSGE(0 To 4) As SGE
    ReDim tmSPE(0 To 6) As SPE
    For ilLoop = 0 To UBound(tmITE) - 1 Step 1
        tmITE(ilLoop).iCode = 0
        If ilLoop = 0 Then
            tmITE(ilLoop).sType = "P"
            tmITE(ilLoop).sName = edcDASSPriName.text
            tmITE(ilLoop).iDataBits = Val(edcDASSPriDataBits.text)
            If cbcDASSPriParity.ListIndex = 0 Then
                tmITE(ilLoop).sParity = "E"
            ElseIf cbcDASSPriParity.ListIndex = 1 Then
                tmITE(ilLoop).sParity = "M"
            ElseIf cbcDASSPriParity.ListIndex = 3 Then
                tmITE(ilLoop).sParity = "O"
            ElseIf cbcDASSPriParity.ListIndex = 4 Then
                tmITE(ilLoop).sParity = "S"
            Else
                tmITE(ilLoop).sParity = "N"
            End If
            If cbcDASSPriStopBit.ListIndex = 1 Then
                tmITE(ilLoop).sStopBit = "2"
            Else
                tmITE(ilLoop).sStopBit = "1"
            End If
            tmITE(ilLoop).iBaud = Val(edcDASSPriBaud.text)
            tmITE(ilLoop).sMachineID = edcDASSPriMachID.text
            tmITE(ilLoop).sStartCode = edcDASSPriStartChar.text
            tmITE(ilLoop).sReplyCode = edcDASSPriReplyChar.text
            tmITE(ilLoop).iMinMgsID = Val(edcDASSPriMinID.text)
            tmITE(ilLoop).iMaxMgsID = Val(edcDASSPriMaxID.text)
            tmITE(ilLoop).iCurrMgsID = Val(edcDASSPriCurID.text)
            tmITE(ilLoop).sMgsType = edcDASSPriMessType.text
            tmITE(ilLoop).sMgsEndCode = edcDASSPriMessEnd.text
            tmITE(ilLoop).sTitleID = edcDASSPriTitleID.text
            tmITE(ilLoop).sLengthID = edcDASSPriLengthID.text
            tmITE(ilLoop).sConnectSeq = edcDASSPriConnectSeq.text
            tmITE(ilLoop).sMgsErrType = edcDASSPriMgsErrType.text
            If ckcDASSPriCheckSum.Value = vbUnchecked Then
                tmITE(ilLoop).sCheckSum = "N"
            Else
                tmITE(ilLoop).sCheckSum = "Y"
            End If
            tmITE(ilLoop).sCmmdSeq = ""
            tmITE(ilLoop).sUnused = ""
        Else
            tmITE(ilLoop).sType = "S"
            tmITE(ilLoop).sName = edcDASSSecName.text
            tmITE(ilLoop).iDataBits = Val(edcDASSSecDataBits.text)
            If cbcDASSSecParity.ListIndex = 0 Then
                tmITE(ilLoop).sParity = "E"
            ElseIf cbcDASSSecParity.ListIndex = 1 Then
                tmITE(ilLoop).sParity = "M"
            ElseIf cbcDASSSecParity.ListIndex = 3 Then
                tmITE(ilLoop).sParity = "O"
            ElseIf cbcDASSSecParity.ListIndex = 4 Then
                tmITE(ilLoop).sParity = "S"
            Else
                tmITE(ilLoop).sParity = "N"
            End If
            If cbcDASSSecStopBit.ListIndex = 1 Then
                tmITE(ilLoop).sStopBit = "2"
            Else
                tmITE(ilLoop).sStopBit = "1"
            End If
            tmITE(ilLoop).iBaud = Val(edcDASSSecBaud.text)
            tmITE(ilLoop).sMachineID = edcDASSSecMachID.text
            tmITE(ilLoop).sStartCode = edcDASSSecStartChar.text
            tmITE(ilLoop).sReplyCode = edcDASSSecReplyChar.text
            tmITE(ilLoop).iMinMgsID = Val(edcDASSSecMinID.text)
            tmITE(ilLoop).iMaxMgsID = Val(edcDASSSecMaxID.text)
            tmITE(ilLoop).iCurrMgsID = Val(edcDASSSecCurID.text)
            tmITE(ilLoop).sMgsType = edcDASSSecMessType.text
            tmITE(ilLoop).sMgsEndCode = edcDASSSecMessEnd.text
            tmITE(ilLoop).sTitleID = edcDASSSecTitleID.text
            tmITE(ilLoop).sLengthID = edcDASSSecLengthID.text
            tmITE(ilLoop).sConnectSeq = edcDASSSecConnectSeq.text
            tmITE(ilLoop).sMgsErrType = edcDASSSecMgsErrType.text
            If ckcDASSSecCheckSum.Value = vbUnchecked Then
                tmITE(ilLoop).sCheckSum = "N"
            Else
                tmITE(ilLoop).sCheckSum = "Y"
            End If
            tmITE(ilLoop).sCmmdSeq = ""
            tmITE(ilLoop).sUnused = ""
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmSPE) - 1 Step 1
        tmSPE(ilLoop).iCode = 0
        tmSPE(ilLoop).sSubType = "P"
        If ilLoop = 0 Then
            tmSPE(ilLoop).sType = "SP"
            tmSPE(ilLoop).sPath = edcPriServerImportPath.text
        ElseIf ilLoop = 1 Then
            tmSPE(ilLoop).sType = "SB"
            tmSPE(ilLoop).sPath = edcBkupServerImportPath.text
        ElseIf ilLoop = 2 Then
            tmSPE(ilLoop).sType = "CP"
            tmSPE(ilLoop).sPath = edcPriClientImportPath.text
        ElseIf ilLoop = 3 Then
            tmSPE(ilLoop).sType = "CB"
            tmSPE(ilLoop).sPath = edcBkupClientImportPath.text
        ElseIf ilLoop = 4 Then
            tmSPE(ilLoop).sType = "SP"
            tmSPE(ilLoop).sSubType = "T"
            tmSPE(ilLoop).sPath = edcPriServerImportPathTest.text
        ElseIf ilLoop = 5 Then
            tmSPE(ilLoop).sType = "CP"
            tmSPE(ilLoop).sSubType = "T"
            tmSPE(ilLoop).sPath = edcPriClientImportPathTest.text
        End If
        tmSPE(ilLoop).sUnused = ""
    Next ilLoop
    For ilLoop = 0 To UBound(tmSGE) - 1 Step 1
        tmSGE(ilLoop).iCode = 0
        If ilLoop = 0 Then
            tmSGE(ilLoop).sType = "S"
            tmSGE(ilLoop).sSubType = "P"
            For llRow = grdSchd.FixedRows To grdSchd.FixedRows + 6 Step 1
                Select Case llRow - grdSchd.FixedRows
                    Case 0
                        tmSGE(ilLoop).iGenMo = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 1
                        tmSGE(ilLoop).iGenTu = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 2
                        tmSGE(ilLoop).iGenWe = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 3
                        tmSGE(ilLoop).iGenTh = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 4
                        tmSGE(ilLoop).iGenFr = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 5
                        tmSGE(ilLoop).iGenSa = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 6
                        tmSGE(ilLoop).iGenSu = Val(grdSchd.TextMatrix(llRow, 1))
                End Select
            Next llRow
            If rbcPurge(0).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = "Y"
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            ElseIf rbcPurge(1).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = ""
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            Else
                tmSGE(ilLoop).sPurgeAfterGen = "N"
                tmSGE(ilLoop).sPurgeTime = gFormatTime(edcPurgeTime.text)
            End If
            tmSGE(ilLoop).sGenTime = gFormatTime(edcSchdGenTime.text)
        ElseIf ilLoop = 1 Then
            tmSGE(ilLoop).sType = "A"
            tmSGE(ilLoop).sSubType = "P"
            For llRow = grdAuto.FixedRows To grdAuto.FixedRows + 6 Step 1
                Select Case llRow - grdAuto.FixedRows
                    Case 0
                        tmSGE(ilLoop).iGenMo = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 1
                        tmSGE(ilLoop).iGenTu = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 2
                        tmSGE(ilLoop).iGenWe = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 3
                        tmSGE(ilLoop).iGenTh = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 4
                        tmSGE(ilLoop).iGenFr = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 5
                        tmSGE(ilLoop).iGenSa = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 6
                        tmSGE(ilLoop).iGenSu = Val(grdAuto.TextMatrix(llRow, 1))
                End Select
            Next llRow
            If rbcPurge(1).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = "Y"
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            ElseIf rbcPurge(0).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = ""
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            Else
                tmSGE(ilLoop).sPurgeAfterGen = "N"
                tmSGE(ilLoop).sPurgeTime = gFormatTime(edcPurgeTime.text)
            End If
            tmSGE(ilLoop).sGenTime = gFormatTime(edcAutoGenTime.text)
        ElseIf ilLoop = 2 Then
            tmSGE(ilLoop).sType = "S"
            tmSGE(ilLoop).sSubType = "T"
            For llRow = grdSchd.FixedRows To grdSchd.FixedRows + 6 Step 1
                Select Case llRow - grdSchd.FixedRows
                    Case 0
                        tmSGE(ilLoop).iGenMo = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 1
                        tmSGE(ilLoop).iGenTu = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 2
                        tmSGE(ilLoop).iGenWe = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 3
                        tmSGE(ilLoop).iGenTh = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 4
                        tmSGE(ilLoop).iGenFr = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 5
                        tmSGE(ilLoop).iGenSa = Val(grdSchd.TextMatrix(llRow, 1))
                    Case 6
                        tmSGE(ilLoop).iGenSu = Val(grdSchd.TextMatrix(llRow, 1))
                End Select
            Next llRow
            If rbcPurgeTest(0).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = "Y"
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            ElseIf rbcPurgeTest(1).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = ""
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            Else
                tmSGE(ilLoop).sPurgeAfterGen = "N"
                tmSGE(ilLoop).sPurgeTime = gFormatTime(edcPurgeTimeTest.text)
            End If
            tmSGE(ilLoop).sGenTime = gFormatTime(edcSchdGenTimeTest.text)
        ElseIf ilLoop = 3 Then
            tmSGE(ilLoop).sType = "A"
            tmSGE(ilLoop).sSubType = "T"
            For llRow = grdAuto.FixedRows To grdAuto.FixedRows + 6 Step 1
                Select Case llRow - grdAuto.FixedRows
                    Case 0
                        tmSGE(ilLoop).iGenMo = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 1
                        tmSGE(ilLoop).iGenTu = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 2
                        tmSGE(ilLoop).iGenWe = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 3
                        tmSGE(ilLoop).iGenTh = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 4
                        tmSGE(ilLoop).iGenFr = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 5
                        tmSGE(ilLoop).iGenSa = Val(grdAuto.TextMatrix(llRow, 1))
                    Case 6
                        tmSGE(ilLoop).iGenSu = Val(grdAuto.TextMatrix(llRow, 1))
                End Select
            Next llRow
            If rbcPurgeTest(1).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = "Y"
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            ElseIf rbcPurgeTest(0).Value Then
                tmSGE(ilLoop).sPurgeAfterGen = ""
                tmSGE(ilLoop).sPurgeTime = gFormatTime("12AM")
            Else
                tmSGE(ilLoop).sPurgeAfterGen = "N"
                tmSGE(ilLoop).sPurgeTime = gFormatTime(edcPurgeTimeTest.text)
            End If
            tmSGE(ilLoop).sGenTime = gFormatTime(edcAutoGenTimeTest.text)
        End If
        tmSGE(ilLoop).sUnused = ""
    Next ilLoop
    
    
    tmSSE.iCode = 0
    tmSSE.iSoeCode = 0
    tmSSE.sEMailHost = edcHost.text
    tmSSE.iEMailPort = Val(edcPort.text)
    tmSSE.sEMailAcctName = edcAccount.text
    tmSSE.sEMailPassword = edcPassword.text
    If rbcTLS(0).Value Then
        tmSSE.sEMailTLS = "Y"
    Else
        tmSSE.sEMailTLS = "N"
    End If
    tmSSE.sUnused = ""

End Sub

Private Sub mMoveRecToCtrls()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim llRow As Long
    
    imSoeCode = tmSOE.iCode
    edcName.text = Trim$(tmSOE.sClientName)
    edcAddress(0).text = Trim$(tmSOE.sAddr1)
    edcAddress(1).text = Trim$(tmSOE.sAddr2)
    edcAddress(2).text = Trim$(tmSOE.sAddr3)
    edcPhone.text = Trim$(tmSOE.sPhone)
    edcFax.text = Trim$(tmSOE.sFax)
    edcRetainSchd.text = Trim$(Str$(tmSOE.iDaysRetainAsAir))
    edcRetainActivityLog.text = Trim$(Str$(tmSOE.iDaysRetainActive))
    edcChgInterval.text = gLongToStrTimeInTenth(tmSOE.lChgInterval)
    edcDateFormat.text = Trim$(tmSOE.sMergeDateFormat)
    edcTimeFormat.text = Trim$(tmSOE.sMergeTimeFormat)
    edcImportFileFormat.text = Trim$(tmSOE.sMergeFileFormat)
    edcImportExt.text = Trim$(tmSOE.sMergeFileExt)
    edcMergeChkStartTime.text = Format$(tmSOE.sMergeStartTime, sgShowTimeWSecForm)
    edcMergeChkEnd = Format$(tmSOE.sMergeEndTime, sgShowTimeWSecForm)
    edcMergeChkInterval.text = Trim$(Str$(tmSOE.iMergeChkInterval))
    If tmSOE.sMergeStopFlag = "Y" Then
        rbcMerge(0).Value = True
    Else
        rbcMerge(1).Value = True
    End If
    edcAlertInterval.text = Trim$(Str$(tmSOE.iAlertInterval))
    edcTimeTolerance.text = gLongToStrLengthInTenth(tmSOE.lTimeTolerance, False)
    edcLengthTolerance.text = gLongToStrLengthInTenth(tmSOE.lLengthTolerance, False)
    edcMinEventID.text = Trim$(Str$(tmSOE.lMinEventID))
    edcMaxEventID.text = Trim$(Str$(tmSOE.lMaxEventID))
    edcCurrEventID.text = Trim$(Str$(tmSOE.lCurrEventID))
    If tmSOE.sSchAutoGenSeq = "S" Then
        rbcSchOrAuto(0).Value = True
    ElseIf tmSOE.sSchAutoGenSeq = "A" Then
        rbcSchOrAuto(1).Value = True
    Else
        rbcSchOrAuto(2).Value = True
    End If
    If tmSOE.sSchAutoGenSeqTst = "S" Then
        rbcSchOrAutoTest(0).Value = True
    ElseIf tmSOE.sSchAutoGenSeqTst = "A" Then
        rbcSchOrAutoTest(1).Value = True
    Else
        rbcSchOrAutoTest(2).Value = True
    End If
    If tmSOE.sMergeStopFlagTst = "Y" Then
        rbcMergeTest(0).Value = True
    Else
        rbcMergeTest(1).Value = True
    End If
    edcSpotItemIDWindow.text = Trim$(Str$(tmSOE.iSpotItemIDWindow))
    If Trim$(tmSOE.sMatchATNotB) = "N" Then
        ckcMatchATNotB.Value = vbUnchecked
    Else
        ckcMatchATNotB.Value = vbChecked
    End If
    If Trim$(tmSOE.sMatchATBNotI) = "N" Then
        ckcMatchATBNotI.Value = vbUnchecked
    Else
        ckcMatchATBNotI.Value = vbChecked
    End If
    If Trim$(tmSOE.sMatchANotT) = "N" Then
        ckcMatchANotT.Value = vbUnchecked
    Else
        ckcMatchANotT.Value = vbChecked
    End If
    If Trim$(tmSOE.sMatchBNotT) = "N" Then
        ckcMatchBNotT.Value = vbUnchecked
    Else
        ckcMatchBNotT.Value = vbChecked
    End If
    imVersion = tmSOE.iVersion
    If UBound(tmITE) > LBound(tmITE) Then
        For ilLoop = 0 To UBound(tmITE) - 1 Step 1
            If tmITE(ilLoop).sType = "P" Then
                edcDASSPriName.text = Trim$(tmITE(ilLoop).sName)
                edcDASSPriDataBits.text = Trim$(Str$(tmITE(ilLoop).iDataBits))
                If tmITE(ilLoop).sParity = "E" Then
                    cbcDASSPriParity.ListIndex = 0
                ElseIf tmITE(ilLoop).sParity = "M" Then
                    cbcDASSPriParity.ListIndex = 1
                ElseIf tmITE(ilLoop).sParity = "O" Then
                    cbcDASSPriParity.ListIndex = 3
                ElseIf tmITE(ilLoop).sParity = "S" Then
                    cbcDASSPriParity.ListIndex = 4
                Else
                    cbcDASSPriParity.ListIndex = 2
                End If
                If tmITE(ilLoop).sStopBit = "2" Then
                    cbcDASSPriStopBit.ListIndex = 1
                Else
                    cbcDASSPriStopBit.ListIndex = 0
                End If
                edcDASSPriBaud.text = Trim$(Str$(tmITE(ilLoop).iBaud))
                edcDASSPriMachID.text = Trim$(tmITE(ilLoop).sMachineID)
                edcDASSPriStartChar.text = Trim$(tmITE(ilLoop).sStartCode)
                edcDASSPriReplyChar.text = Trim$(tmITE(ilLoop).sReplyCode)
                edcDASSPriMinID.text = Trim$(Str$(tmITE(ilLoop).iMinMgsID))
                edcDASSPriMaxID.text = Trim$(Str$(tmITE(ilLoop).iMaxMgsID))
                edcDASSPriCurID.text = Trim$(Str$(tmITE(ilLoop).iCurrMgsID))
                edcDASSPriMessType.text = Trim$(tmITE(ilLoop).sMgsType)
                edcDASSPriMessEnd.text = Trim$(tmITE(ilLoop).sMgsEndCode)
                edcDASSPriTitleID.text = Trim$(tmITE(ilLoop).sTitleID)
                edcDASSPriLengthID.text = Trim$(tmITE(ilLoop).sLengthID)
                edcDASSPriConnectSeq.text = Trim$(tmITE(ilLoop).sConnectSeq)
                edcDASSPriMgsErrType.text = Trim$(tmITE(ilLoop).sMgsErrType)
                If Trim$(tmITE(ilLoop).sCheckSum) = "N" Then
                    ckcDASSPriCheckSum.Value = vbUnchecked
                Else
                    ckcDASSPriCheckSum.Value = vbChecked
                End If
            Else
                edcDASSSecName.text = Trim$(tmITE(ilLoop).sName)
                edcDASSSecDataBits.text = Trim$(Str$(tmITE(ilLoop).iDataBits))
                If tmITE(ilLoop).sParity = "E" Then
                    cbcDASSSecParity.ListIndex = 0
                ElseIf tmITE(ilLoop).sParity = "M" Then
                    cbcDASSSecParity.ListIndex = 1
                ElseIf tmITE(ilLoop).sParity = "O" Then
                    cbcDASSSecParity.ListIndex = 3
                ElseIf tmITE(ilLoop).sParity = "S" Then
                    cbcDASSSecParity.ListIndex = 4
                Else
                    cbcDASSSecParity.ListIndex = 2
                End If
                If tmITE(ilLoop).sStopBit = "2" Then
                    cbcDASSSecStopBit.ListIndex = 1
                Else
                    cbcDASSSecStopBit.ListIndex = 0
                End If
                edcDASSSecBaud.text = Trim$(Str$(tmITE(ilLoop).iBaud))
                edcDASSSecMachID.text = Trim$(tmITE(ilLoop).sMachineID)
                edcDASSSecStartChar.text = Trim$(tmITE(ilLoop).sStartCode)
                edcDASSSecReplyChar.text = Trim$(tmITE(ilLoop).sReplyCode)
                edcDASSSecMinID.text = Trim$(Str$(tmITE(ilLoop).iMinMgsID))
                edcDASSSecMaxID.text = Trim$(Str$(tmITE(ilLoop).iMaxMgsID))
                edcDASSSecCurID.text = Trim$(Str$(tmITE(ilLoop).iCurrMgsID))
                edcDASSSecMessType.text = Trim$(tmITE(ilLoop).sMgsType)
                edcDASSSecMessEnd.text = Trim$(tmITE(ilLoop).sMgsEndCode)
                edcDASSSecTitleID.text = Trim$(tmITE(ilLoop).sTitleID)
                edcDASSSecLengthID.text = Trim$(tmITE(ilLoop).sLengthID)
                edcDASSSecConnectSeq.text = Trim$(tmITE(ilLoop).sConnectSeq)
                edcDASSSecMgsErrType.text = Trim$(tmITE(ilLoop).sMgsErrType)
                If Trim$(tmITE(ilLoop).sCheckSum) = "N" Then
                    ckcDASSSecCheckSum.Value = vbUnchecked
                Else
                    ckcDASSSecCheckSum.Value = vbChecked
                End If
            End If
        Next ilLoop
    Else
        edcDASSPriName.text = ""
        edcDASSPriDataBits.text = ""
        cbcDASSPriParity.ListIndex = 2
        cbcDASSPriStopBit.ListIndex = 0
        edcDASSPriBaud.text = "9600"
        edcDASSPriMachID.text = ""
        edcDASSPriStartChar.text = ""
        edcDASSPriReplyChar.text = ""
        edcDASSPriMinID.text = ""
        edcDASSPriMaxID.text = ""
        edcDASSPriCurID.text = ""
        edcDASSPriMessType.text = ""
        edcDASSPriMessEnd.text = ""
        edcDASSPriTitleID.text = ""
        edcDASSPriLengthID.text = ""
        edcDASSPriConnectSeq.text = ""
        edcDASSSecName.text = ""
        edcDASSSecDataBits.text = "8"
        cbcDASSPriParity.ListIndex = 2
        cbcDASSSecStopBit.ListIndex = 0
        edcDASSSecBaud.text = "9600"
        edcDASSSecMachID.text = ""
        edcDASSSecStartChar.text = ""
        edcDASSSecReplyChar.text = ""
        edcDASSSecMinID.text = ""
        edcDASSSecMaxID.text = ""
        edcDASSSecCurID.text = ""
        edcDASSSecMessType.text = ""
        edcDASSSecMessEnd.text = ""
        edcDASSSecTitleID.text = ""
        edcDASSSecLengthID.text = ""
        edcDASSSecConnectSeq.text = ""
    End If
    If UBound(tmSPE) > LBound(tmSPE) Then
        For ilLoop = 0 To UBound(tmSPE) - 1 Step 1
            If (tmSPE(ilLoop).sType = "SP") And ((tmSPE(ilLoop).sSubType = "P") Or (Trim$(tmSPE(ilLoop).sSubType) = "")) Then
                edcPriServerImportPath.text = Trim$(tmSPE(ilLoop).sPath)
            ElseIf tmSPE(ilLoop).sType = "SB" Then
                edcBkupServerImportPath.text = Trim$(tmSPE(ilLoop).sPath)
            ElseIf (tmSPE(ilLoop).sType = "CP") And ((tmSPE(ilLoop).sSubType = "P") Or (Trim$(tmSPE(ilLoop).sSubType) = "")) Then
                edcPriClientImportPath.text = Trim$(tmSPE(ilLoop).sPath)
            ElseIf tmSPE(ilLoop).sType = "CB" Then
                edcBkupClientImportPath.text = Trim$(tmSPE(ilLoop).sPath)
            ElseIf (tmSPE(ilLoop).sType = "SP") And (tmSPE(ilLoop).sSubType = "T") Then
                edcPriServerImportPathTest.text = Trim$(tmSPE(ilLoop).sPath)
            ElseIf (tmSPE(ilLoop).sType = "CP") And (tmSPE(ilLoop).sSubType = "T") Then
                edcPriClientImportPathTest.text = Trim$(tmSPE(ilLoop).sPath)
            End If
        Next ilLoop
    Else
        edcPriServerImportPath.text = ""
        edcBkupServerImportPath.text = ""
        edcPriClientImportPath.text = ""
        edcBkupClientImportPath.text = ""
        edcPriServerImportPathTest.text = ""
        edcPriClientImportPathTest.text = ""
    End If
    grdSchd.Redraw = False
    grdAuto.Redraw = False
    If UBound(tmSGE) > LBound(tmSGE) Then
        For ilLoop = 0 To UBound(tmSGE) - 1 Step 1
            If (tmSGE(ilLoop).sType = "S") And ((tmSGE(ilLoop).sSubType = "P") Or (Trim$(tmSGE(ilLoop).sSubType) = "")) Then
                For llRow = grdSchd.FixedRows To grdSchd.FixedRows + 6 Step 1
    '                If llSchdRow + 1 > grdSchd.Rows Then
    '                    grdSchd.AddItem ""
    '                End If
                    Select Case llRow - grdSchd.FixedRows
                        Case 0
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenMo)
                        Case 1
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenTu)
                        Case 2
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenWe)
                        Case 3
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenTh)
                        Case 4
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenFr)
                        Case 5
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenSa)
                        Case 6
                            grdSchd.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenSu)
                    End Select
                Next llRow
                'Don't add extra row
    '            If llRow >= grdSchd.Rows Then
    '                 grdSchd.AddItem ""
    '            End If
                If tmSGE(ilLoop).sPurgeAfterGen = "Y" Then
                    rbcPurge(0).Value = True
                    edcPurgeTime.text = ""
                ElseIf tmSGE(ilLoop).sPurgeAfterGen = "N" Then
                    rbcPurge(2).Value = True
                    edcPurgeTime.text = Trim$(tmSGE(ilLoop).sPurgeTime)
                End If
                If rbcSchOrAuto(0).Value Or rbcSchOrAuto(2).Value Then
                    edcSchdGenTime.text = Trim$(tmSGE(ilLoop).sGenTime)
                Else
                    edcSchdGenTime.text = ""
                End If
            ElseIf (tmSGE(ilLoop).sType = "A") And ((tmSGE(ilLoop).sSubType = "P") Or (Trim$(tmSGE(ilLoop).sSubType) = "")) Then
                For llRow = grdAuto.FixedRows To grdAuto.FixedRows + 6 Step 1
                    Select Case llRow - grdAuto.FixedRows
                        Case 0
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenMo)
                        Case 1
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenTu)
                        Case 2
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenWe)
                        Case 3
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenTh)
                        Case 4
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenFr)
                        Case 5
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenSa)
                        Case 6
                            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenSu)
                    End Select
                Next llRow
                If tmSGE(ilLoop).sPurgeAfterGen = "Y" Then
                    rbcPurge(1).Value = True
                    edcPurgeTime.text = ""
                ElseIf tmSGE(ilLoop).sPurgeAfterGen = "N" Then
                    rbcPurge(2).Value = True
                    edcPurgeTime.text = Trim$(tmSGE(ilLoop).sPurgeTime)
                End If
                If rbcSchOrAuto(1).Value Or rbcSchOrAuto(2).Value Then
                    edcAutoGenTime.text = Trim$(tmSGE(ilLoop).sGenTime)
                Else
                    edcAutoGenTime.text = ""
                End If
            ElseIf (tmSGE(ilLoop).sType = "S") And (tmSGE(ilLoop).sSubType = "T") Then
                'Don't add extra row
    '            If llRow >= grdSchd.Rows Then
    '                 grdSchd.AddItem ""
    '            End If
                If tmSGE(ilLoop).sPurgeAfterGen = "Y" Then
                    rbcPurgeTest(0).Value = True
                    edcPurgeTimeTest.text = ""
                ElseIf tmSGE(ilLoop).sPurgeAfterGen = "N" Then
                    rbcPurgeTest(2).Value = True
                    edcPurgeTimeTest.text = Trim$(tmSGE(ilLoop).sPurgeTime)
                End If
                If rbcSchOrAutoTest(0).Value Or rbcSchOrAutoTest(2).Value Then
                    edcSchdGenTimeTest.text = Trim$(tmSGE(ilLoop).sGenTime)
                Else
                    edcSchdGenTimeTest.text = ""
                End If
            ElseIf (tmSGE(ilLoop).sType = "A") And (tmSGE(ilLoop).sSubType = "T") Then
                If tmSGE(ilLoop).sPurgeAfterGen = "Y" Then
                    rbcPurgeTest(1).Value = True
                    edcPurgeTimeTest.text = ""
                ElseIf tmSGE(ilLoop).sPurgeAfterGen = "N" Then
                    rbcPurgeTest(2).Value = True
                    edcPurgeTimeTest.text = Trim$(tmSGE(ilLoop).sPurgeTime)
                End If
                If rbcSchOrAutoTest(1).Value Or rbcSchOrAutoTest(2).Value Then
                    edcAutoGenTimeTest.text = Trim$(tmSGE(ilLoop).sGenTime)
                Else
                    edcAutoGenTimeTest.text = ""
                End If
            End If
        Next ilLoop
    Else
        For llRow = grdSchd.FixedRows To grdSchd.FixedRows + 6 Step 1
            grdSchd.TextMatrix(llRow, 1) = ""
        Next llRow
        rbcPurge(0).Value = False
        rbcPurge(1).Value = False
        rbcPurge(2).Value = False
        edcPurgeTime.text = ""
        edcSchdGenTime.text = Format$("00:00:00", sgShowTimeWSecForm)
        For llRow = grdAuto.FixedRows To grdAuto.FixedRows + 6 Step 1
            grdAuto.TextMatrix(llRow, 1) = Trim$(tmSGE(ilLoop).iGenMo)
        Next llRow
        edcAutoGenTime.text = Format$("00:00:00", sgShowTimeWSecForm)
    
        rbcPurgeTest(0).Value = False
        rbcPurgeTest(1).Value = False
        rbcPurgeTest(2).Value = False
        edcPurgeTimeTest.text = ""
        edcSchdGenTimeTest.text = Format$("00:00:00", sgShowTimeWSecForm)
        edcAutoGenTimeTest.text = Format$("00:00:00", sgShowTimeWSecForm)
    End If
    grdSchd.Redraw = True
    grdAuto.Redraw = True
    
    edcHost.text = Trim$(tmSSE.sEMailHost)
    If tmSSE.iEMailPort <= 0 Then
        edcPort.text = ""
    Else
        edcPort.text = Trim$(Str$(tmSSE.iEMailPort))
    End If
    edcAccount.text = Trim$(tmSSE.sEMailAcctName)
    edcPassword.text = Trim$(tmSSE.sEMailPassword)
    If tmSSE.sEMailTLS = "Y" Then
        rbcTLS(0).Value = True
    Else
        rbcTLS(1).Value = True
    End If

End Sub

Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlUte As UTE
    
    Screen.MousePointer = vbHourglass
    If Not mCheckFields(True) Then
        Screen.MousePointer = vbDefault
        mSave = False
        Exit Function
    End If
    mMoveCtrlsToRec
    If imSoeCode <= 0 Then
        ilRet = gPutInsert_SOE_SiteOption(0, tmSOE, "Site Option-mSave")
    Else
        ilRet = gPutUpdate_SOE_SiteOption(0, tmSOE, "Site Option-mSave")
    End If
    'Insert Task setting
    For ilLoop = 0 To UBound(tmITE) - 1 Step 1
        If (imSoeCode <= 0) Or (tmITE(ilLoop).iSoeCode = 0) Then
            tmITE(ilLoop).iSoeCode = tmSOE.iCode
        End If
        ilRet = gPutInsert_ITE_ItemTest(tmITE(ilLoop), "Site Option-mSave: ITE")
    Next ilLoop
    For ilLoop = 0 To UBound(tmSGE) - 1 Step 1
        If (imSoeCode <= 0) Or (tmSGE(ilLoop).iSoeCode = 0) Then
            tmSGE(ilLoop).iSoeCode = tmSOE.iCode
        End If
        ilRet = gPutInsert_SGE_SiteGenSchd(tmSGE(ilLoop), "Site Option-mSave: SGE")
    Next ilLoop
    For ilLoop = 0 To UBound(tmSPE) - 1 Step 1
        If (imSoeCode <= 0) Or (tmSPE(ilLoop).iSoeCode = 0) Then
            tmSPE(ilLoop).iSoeCode = tmSOE.iCode
        End If
        ilRet = gPutInsert_SPE_SitePath(tmSPE(ilLoop), "Site Option-mSave: SPE")
    Next ilLoop
    tmSSE.iSoeCode = tmSOE.iCode
    ilRet = gPutInsert_SSE_Site_SMTP_Info(tmSSE, "Site Option-mSave: SSE")
    
    sgCurrSOEStamp = ""
    sgCurrSGEStamp = ""
    sgCurrSPEStamp = ""
    
    ilRet = gGetTypeOfRecs_SOE_SiteOption("C", sgCurrSOEStamp, "LogIn-mChkForSiteOption", tgCurrSOE())
    LSet tgSOE = tgCurrSOE(0)
    ilRet = gGetRecs_ITE_ItemTest(sgCurrITEStamp, tgSOE.iCode, "LogIn-mChkForSiteOpion: Get ITE", tgCurrITE())
    ilRet = gGetRecs_SGE_SiteGenSchd(sgCurrSGEStamp, tgSOE.iCode, "LogIn-mChkForSiteOpion: Get SGE", tgCurrSGE())
    ilRet = gGetRecs_SPE_SitePath(sgCurrSPEStamp, tgSOE.iCode, "LogIn-mChkForSiteOpion: Get SPE", tgCurrSPE())
    ilRet = gGetRec_SSE_Site_SMTP_Info(tgSOE.iCode, "Site Option-mInit: Get SSE", tgCurrSSE)
    imFieldChgd = False
    mSetCommands
    Screen.MousePointer = vbDefault
    mSave = True
End Function

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
End Sub



Private Sub mGridColumns(grdCtrl As MSHFlexGrid)
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    gGrid_AlignAllColsLeft grdCtrl
    'Hide column 26
    grdCtrl.ColWidth(CODEINDEX) = 0
    grdCtrl.ColWidth(LOGDAYINDEX) = grdCtrl.Width / 5
    grdCtrl.ColWidth(LEADDAYINDEX) = grdCtrl.Width - grdCtrl.ColWidth(LOGDAYINDEX) '- GRIDSCROLLWIDTH
    'Set Titles
    grdCtrl.TextMatrix(0, 0) = "Log"
    grdCtrl.TextMatrix(1, 0) = "Mo"
    grdCtrl.TextMatrix(2, 0) = "Tu"
    grdCtrl.TextMatrix(3, 0) = "We"
    grdCtrl.TextMatrix(4, 0) = "Th"
    grdCtrl.TextMatrix(5, 0) = "Fr"
    grdCtrl.TextMatrix(6, 0) = "Sa"
    grdCtrl.TextMatrix(7, 0) = "Su"
    grdCtrl.TextMatrix(0, 1) = "Lead Time (Days)"
    grdCtrl.Height = 8 * grdCtrl.RowHeight(0) + 30
    gGrid_IntegralHeight grdCtrl
    'gGrid_Clear grdEvent, True
    For ilRow = 1 To 7 Step 1
        'grdCtrl.TextMatrix(ilRow, 1) = ""
        grdCtrl.Row = ilRow
        grdCtrl.Col = 0
        grdCtrl.CellBackColor = LIGHTYELLOW
    Next ilRow
    grdCtrl.Row = grdCtrl.FixedRows
End Sub



Private Sub mClearControls()
    Dim ilCol As Integer
    imVersion = -1
    gClearControls EngrSiteOption
    For ilCol = 1 To 7 Step 1
        grdSchd.TextMatrix(ilCol, 1) = ""
    Next ilCol
    imFieldChgd = False
End Sub

Private Sub cbcDASSPriParity_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cbcDASSPriStopBit_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cbcDASSSecParity_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cbcDASSSecStopBit_Change()
    imFieldChgd = True
    mSetCommands
End Sub


Private Sub ckcDASSPriCheckSum_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub ckcDASSSecCheckSum_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub ckcMatchANotT_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub ckcMatchATBNotI_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub ckcMatchATNotB_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub ckcMatchBNotT_Click()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cmcCancel_Click()
    Unload EngrSiteOption
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        Unload EngrSiteOption
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Unload EngrSiteOption
    Exit Sub
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    
    If imFieldChgd = True Then
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
    End If
End Sub

Private Sub cmcVerify_Click()
    Set ogEmailer = New CEmail
    Dim blTls As Boolean
    
    edcResult.ForeColor = vbBlack
    edcResult.text = "attempting to send test message...please wait."
    Screen.MousePointer = vbHourglass
    If (LenB(edcHost.text) = 0) Or (LenB(edcPassword.text) = 0) Or (LenB(edcAccount.text) = 0) Or (LenB(edcPort.text) = 0) Or (LenB(edcTo.text) = 0) Then
        edcResult.text = "A field is blank.  Cannot verify."
        Screen.MousePointer = vbDefault
        Set ogEmailer = Nothing
        Exit Sub
    End If
    If rbcTLS(0) = vbChecked Then
        blTls = True
    Else
        blTls = False
    End If
    With ogEmailer
        .Message = "Email is working correctly."
        .Subject = "verify email"
        .ToAddress = Trim$(edcTo.text)
        .FromAddress = "testVerify@csi.net"
        .FromName = "csi site test"
        .SetHost Trim$(edcHost.text), Trim$(edcPort.text), Trim$(edcAccount.text), Trim$(edcPassword.text), blTls
    End With
    If ogEmailer.Send(edcResult) Then
        If InStr(1, edcResult.text, "sent", vbTextCompare) > 0 Then
            edcResult.text = "Verified"
        End If
    Else
        If InStr(1, edcResult.text, "11004", vbTextCompare) > 0 Then
            edcResult.text = "Host name not recognized."
        ElseIf InStr(1, edcResult.text, "535") > 0 Then
            edcResult.text = "Password not correct."
        ElseIf InStr(1, edcResult.text, "454") > 0 Then
            edcResult.text = "acount name not recognized."
        ElseIf InStr(1, edcResult.text, "Time", vbTextCompare) > 0 Then
            edcResult.text = "Session timed out.  Port setting may not be correct."
        ElseIf InStr(1, edcResult.text, "18") > 0 Then
            edcResult.text = "Transcript Layer Security must be set to false."
        End If
    End If
    Set ogEmailer = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub edcAccount_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcAccount_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAddress_Change(Index As Integer)
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcAddress_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAlertInterval_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcAlertInterval_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAutoGenTime_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcAutoGenTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAutoGenTime_LostFocus()
    Dim slStr As String
    
    slStr = edcAutoGenTime.text
    edcAutoGenTime.text = gFormatTime(slStr)
End Sub

Private Sub edcAutoGenTimeTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcAutoGenTimeTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAutoGenTimeTest_LostFocus()
    Dim slStr As String
    
    slStr = edcAutoGenTimeTest.text
    edcAutoGenTimeTest.text = gFormatTime(slStr)
End Sub

Private Sub edcAutoGrid_Change()
    Dim slStr As String
    
    Select Case grdAuto.Col
        Case LOGDAYINDEX
        Case LEADDAYINDEX
            If grdAuto.text <> edcAutoGrid.text Then
                imFieldChgd = True
            End If
            grdAuto.text = edcAutoGrid.text
            grdAuto.CellForeColor = vbBlack
    End Select
    mSetCommands
End Sub

Private Sub edcAutoGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcBkupClientImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcBkupClientImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcBkupServerImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcBkupServerImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcChgInterval_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcChgInterval_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcChgInterval_LostFocus()
    Dim slStr As String
    
    slStr = edcChgInterval.text
    edcChgInterval.text = gFormatTimeTenths(slStr)
End Sub

Private Sub edcCurrEventID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcCurrEventID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriBaud_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriBaud_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriConnectSeq_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriConnectSeq_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriCurID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriCurID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriDataBits_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriDataBits_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriLengthID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriLengthID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMachID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMachID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMaxID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMaxID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMessEnd_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMessEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMessType_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMessType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMinID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMinID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriMgsErrType_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriMgsErrType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriReplyChar_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriReplyChar_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriStartChar_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriStartChar_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSPriTitleID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSPriTitleID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecBaud_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecBaud_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecConnectSeq_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecConnectSeq_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecCurID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecCurID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecDataBits_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecDataBits_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecLengthID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecLengthID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMachID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMachID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMaxID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMaxID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMessEnd_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMessEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMessType_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMessType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMinID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMinID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecMgsErrType_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecMgsErrType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecReplyChar_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecReplyChar_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecStartChar_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecStartChar_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDASSSecTitleID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDASSSecTitleID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDateFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDateFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFax_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcHost_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcHost_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcImportExt_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcImportExt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcImportFileFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcImportFileFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcLengthTolerance_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcLengthTolerance_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMaxEventID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcMaxEventID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMergeChkEnd_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcMergeChkEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMergeChkEnd_LostFocus()
    Dim slStr As String
    
    slStr = edcMergeChkEnd.text
    edcMergeChkEnd.text = gFormatTime(slStr)
End Sub

Private Sub edcMergeChkInterval_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcMergeChkInterval_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMergeChkStartTime_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcMergeChkStartTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMergeChkStartTime_LostFocus()
    Dim slStr As String
    
    slStr = edcMergeChkStartTime.text
    edcMergeChkStartTime.text = gFormatTime(slStr)
End Sub



Private Sub edcMinEventID_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcMinEventID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPassword_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPhone_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPort_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPort_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPort_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcPriClientImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriClientImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriClientImportPathTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriClientImportPathTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriServerImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriServerImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriServerImportPathTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriServerImportPathTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPurgeTimeTest_Change()
    imFieldChgd = True
    If edcPurgeTime <> "" Then
        rbcPurge(2).Value = True
    End If
    mSetCommands
End Sub

Private Sub edcPurgeTimeTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPurgeTimeTest_LostFocus()
    Dim slStr As String
    
    slStr = edcPurgeTimeTest.text
    edcPurgeTimeTest.text = gFormatTime(slStr)
End Sub

Private Sub edcRetainActivityLog_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcRetainActivityLog_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcRetainSchd_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcRetainSchd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSchdGenTime_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSchdGenTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSchdGenTime_LostFocus()
    Dim slStr As String
    
    slStr = edcSchdGenTime.text
    edcSchdGenTime.text = gFormatTime(slStr)
End Sub

Private Sub edcSchdGenTimeTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSchdGenTimeTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSchdGenTimeTest_LostFocus()
    Dim slStr As String
    
    slStr = edcSchdGenTimeTest.text
    edcSchdGenTimeTest.text = gFormatTime(slStr)
End Sub

Private Sub edcSchdGrid_Change()
    Dim slStr As String
    
    Select Case grdSchd.Col
        Case LOGDAYINDEX
        Case LEADDAYINDEX
            If grdSchd.text <> edcSchdGrid.text Then
                imFieldChgd = True
            End If
            grdSchd.text = edcSchdGrid.text
            grdSchd.CellForeColor = vbBlack
    End Select
    mSetCommands
End Sub

Private Sub edcSchdGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPurgeTime_Change()
    imFieldChgd = True
    If edcPurgeTime <> "" Then
        rbcPurge(2).Value = True
    End If
    mSetCommands
End Sub

Private Sub edcPurgeTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPurgeTime_LostFocus()
    Dim slStr As String
    
    slStr = edcPurgeTime.text
    edcPurgeTime.text = gFormatTime(slStr)
End Sub

Private Sub edcSpotItemIDWindow_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSpotItemIDWindow_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTimeFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcTimeFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTimeTolerance_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcTimeTolerance_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    mGridColumns grdSchd
    mGridColumns grdAuto
    mSetTab
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrSiteOption
    gCenterFormModal EngrSiteOption
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    mInit
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Site Option-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Site Option-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub Form_Resize()
    mSetTab
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase tmITE
    Erase tmSGE
    Erase tmSPE
    Set EngrSiteOption = Nothing
End Sub



Private Sub edcName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub





Private Sub grdAuto_Click()
    If grdAuto.Col >= grdAuto.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAuto_EnterCell()
    mAutoSetShow
End Sub

Private Sub grdAuto_GotFocus()
    If grdAuto.Col >= grdAuto.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAuto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdAuto.RowHeight(0) Then
'        mSortCol grdAuto.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAuto, x, y)
    If Not ilFound Then
        grdAuto.Redraw = True
        cmcCancel.SetFocus
        Exit Sub
    End If
    If grdAuto.Col >= grdAuto.Cols - 1 Then
        grdAuto.Redraw = True
        Exit Sub
    End If
'    lmTopRow = grdAuto.TopRow
    
'    llRow = grdAuto.Row
'    If grdAuto.TextMatrix(llRow, LEADDAYINDEX) = "" Then
'        grdAuto.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdAuto.TextMatrix(llRow, LEADDAYINDEX) = ""
'        grdAuto.Row = llRow + 1
'        grdAuto.Col = LEADDAYINDEX
'        grdAuto.Redraw = True
'    End If
    grdAuto.Redraw = True
    mAutoEnableBox
End Sub

Private Sub grdSchd_Click()
    If grdSchd.Col >= grdSchd.Cols - 1 Then
        Exit Sub
    End If

End Sub

Private Sub grdSchd_EnterCell()
    mSchdSetShow
End Sub

Private Sub grdSchd_GotFocus()
    If grdSchd.Col >= grdSchd.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdSchd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
'    lmTopRow = grdSchd.TopRow
'    grdSchd.Redraw = False
End Sub

Private Sub grdSchd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdSchd.RowHeight(0) Then
'        mSortCol grdSchd.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdSchd, x, y)
    If Not ilFound Then
        grdSchd.Redraw = True
        cmcCancel.SetFocus
        Exit Sub
    End If
    If grdSchd.Col >= grdSchd.Cols - 1 Then
        grdSchd.Redraw = True
        Exit Sub
    End If
'    lmTopRow = grdSchd.TopRow
    
'    llRow = grdSchd.Row
'    If grdSchd.TextMatrix(llRow, LEADDAYINDEX) = "" Then
'        grdSchd.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdSchd.TextMatrix(llRow, LEADDAYINDEX) = ""
'        grdSchd.Row = llRow + 1
'        grdSchd.Col = LEADDAYINDEX
'        grdSchd.Redraw = True
'    End If
    grdSchd.Redraw = True
    mSchdEnableBox
End Sub

Private Sub pbcAutoSTab_GotFocus()
    If GetFocus() <> pbcAutoSTab.hwnd Then
        Exit Sub
    End If
    If imAutoFromArrow Then
        imAutoFromArrow = False
        mAutoEnableBox
        Exit Sub
    End If
    If edcAutoGrid.Visible Then
        mAutoSetShow
        If grdAuto.Col = LEADDAYINDEX Then
            If grdAuto.Row > grdAuto.FixedRows Then
'                lmTopRow = -1
                grdAuto.Row = grdAuto.Row - 1
                If Not grdAuto.RowIsVisible(grdAuto.Row) Then
                    grdAuto.TopRow = grdAuto.TopRow - 1
                End If
                grdAuto.Col = LEADDAYINDEX
                mAutoEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdAuto.Col = grdAuto.Col - 1
            mAutoEnableBox
        End If
    Else
'        lmTopRow = -1
        grdAuto.TopRow = grdAuto.FixedRows
        grdAuto.Col = LEADDAYINDEX
        grdAuto.Row = grdAuto.FixedRows
        mAutoEnableBox
    End If

End Sub

Private Sub pbcAutoTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcAutoTab.hwnd Then
        Exit Sub
    End If
    If edcAutoGrid.Visible Then
        mAutoSetShow
        If grdAuto.Col = LEADDAYINDEX Then
            llRow = grdAuto.Rows
            Do
                llRow = llRow - 1
            Loop While grdAuto.TextMatrix(llRow, LEADDAYINDEX) = ""
            llRow = llRow + 1
            If (grdAuto.Row + 1 < llRow) Then
'                lmTopRow = -1
                grdAuto.Row = grdAuto.Row + 1
                If Not grdAuto.RowIsVisible(grdAuto.Row) Then
                    grdAuto.TopRow = grdAuto.TopRow + 1
                End If
                grdAuto.Col = LEADDAYINDEX
                'grdAuto.TextMatrix(grdAuto.Row, CODEINDEX) = 0
                If Trim$(grdAuto.TextMatrix(grdAuto.Row, LEADDAYINDEX)) <> "" Then
                    mAutoEnableBox
                Else
                    imAutoFromArrow = True
                    pbcAutoArrow.Move grdAuto.Left - pbcAutoArrow.Width - 30, grdAuto.Top + grdAuto.RowPos(grdAuto.Row) + (grdAuto.RowHeight(grdAuto.Row) - pbcAutoArrow.Height) / 2
                    pbcAutoArrow.Visible = True
                    pbcAutoArrow.SetFocus
                End If
            Else
                If Trim$(grdAuto.TextMatrix(lmAutoEnableRow, LEADDAYINDEX)) <> "" Then
'                    lmTopRow = -1
                    If grdAuto.Row + 1 >= grdAuto.Rows Then
                        If edcAutoGenTime.Enabled Then
                            edcAutoGenTime.SetFocus
                            Exit Sub
                        End If
                        If edcAutoGenTime.Enabled Then
                            edcAutoGenTime.SetFocus
                            Exit Sub
                        End If
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                    grdAuto.Row = grdAuto.Row + 1
                    If Not grdAuto.RowIsVisible(grdAuto.Row) Then
                        grdAuto.TopRow = grdAuto.TopRow + 1
                    End If
                    grdAuto.Col = LEADDAYINDEX
                    'grdAuto.TextMatrix(grdAuto.Row, CODEINDEX) = 0
                    'mAutoEnableBox
                    imAutoFromArrow = True
                    pbcAutoArrow.Move grdAuto.Left - pbcAutoArrow.Width - 30, grdAuto.Top + grdAuto.RowPos(grdAuto.Row) + (grdAuto.RowHeight(grdAuto.Row) - pbcAutoArrow.Height) / 2
                    pbcAutoArrow.Visible = True
                    pbcAutoArrow.SetFocus
                Else
                    If edcAutoGenTime.Enabled Then
                        edcAutoGenTime.SetFocus
                        Exit Sub
                    End If
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                
            End If
        Else
            grdAuto.Col = grdAuto.Col + 1
            mAutoEnableBox
        End If
    Else
'        lmTopRow = -1
        grdAuto.TopRow = grdAuto.FixedRows
        grdAuto.Col = LEADDAYINDEX
        grdAuto.Row = grdAuto.FixedRows
        mAutoEnableBox
    End If
End Sub

Private Sub pbcSchdSTab_GotFocus()
    If GetFocus() <> pbcSchdSTab.hwnd Then
        Exit Sub
    End If
    If imSchdFromArrow Then
        imSchdFromArrow = False
        mSchdEnableBox
        Exit Sub
    End If
    If edcSchdGrid.Visible Then
        mSchdSetShow
        If grdSchd.Col = LEADDAYINDEX Then
            If grdSchd.Row > grdSchd.FixedRows Then
'                lmTopRow = -1
                grdSchd.Row = grdSchd.Row - 1
                If Not grdSchd.RowIsVisible(grdSchd.Row) Then
                    grdSchd.TopRow = grdSchd.TopRow - 1
                End If
                grdSchd.Col = LEADDAYINDEX
                mSchdEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdSchd.Col = grdSchd.Col - 1
            mSchdEnableBox
        End If
    Else
'        lmTopRow = -1
        grdSchd.TopRow = grdSchd.FixedRows
        grdSchd.Col = LEADDAYINDEX
        grdSchd.Row = grdSchd.FixedRows
        mSchdEnableBox
    End If

End Sub

Private Sub pbcSchdTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcSchdTab.hwnd Then
        Exit Sub
    End If
    If edcSchdGrid.Visible Then
        mSchdSetShow
        If grdSchd.Col = LEADDAYINDEX Then
            llRow = grdSchd.Rows
            Do
                llRow = llRow - 1
            Loop While grdSchd.TextMatrix(llRow, LEADDAYINDEX) = ""
            llRow = llRow + 1
            If (grdSchd.Row + 1 < llRow) Then
'                lmTopRow = -1
                grdSchd.Row = grdSchd.Row + 1
                If Not grdSchd.RowIsVisible(grdSchd.Row) Then
                    grdSchd.TopRow = grdSchd.TopRow + 1
                End If
                grdSchd.Col = LEADDAYINDEX
                'grdSchd.TextMatrix(grdSchd.Row, CODEINDEX) = 0
                If Trim$(grdSchd.TextMatrix(grdSchd.Row, LEADDAYINDEX)) <> "" Then
                    mSchdEnableBox
                Else
                    imSchdFromArrow = True
                    pbcSchdArrow.Move grdSchd.Left - pbcSchdArrow.Width - 30, grdSchd.Top + grdSchd.RowPos(grdSchd.Row) + (grdSchd.RowHeight(grdSchd.Row) - pbcSchdArrow.Height) / 2
                    pbcSchdArrow.Visible = True
                    pbcSchdArrow.SetFocus
                End If
            Else
                If Trim$(grdSchd.TextMatrix(lmSchdEnableRow, LEADDAYINDEX)) <> "" Then
'                    lmTopRow = -1
                    If grdSchd.Row + 1 >= grdSchd.Rows Then
                        If edcSchdGenTime.Enabled Then
                            edcSchdGenTime.SetFocus
                            Exit Sub
                        End If
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                    grdSchd.Row = grdSchd.Row + 1
                    If Not grdSchd.RowIsVisible(grdSchd.Row) Then
                        grdSchd.TopRow = grdSchd.TopRow + 1
                    End If
                    grdSchd.Col = LEADDAYINDEX
                    'grdSchd.TextMatrix(grdSchd.Row, CODEINDEX) = 0
                    'mSchdEnableBox
                    imSchdFromArrow = True
                    pbcSchdArrow.Move grdSchd.Left - pbcSchdArrow.Width - 30, grdSchd.Top + grdSchd.RowPos(grdSchd.Row) + (grdSchd.RowHeight(grdSchd.Row) - pbcSchdArrow.Height) / 2
                    pbcSchdArrow.Visible = True
                    pbcSchdArrow.SetFocus
                Else
                    If edcSchdGenTime.Enabled Then
                        edcSchdGenTime.SetFocus
                        Exit Sub
                    End If
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                
            End If
        Else
            grdSchd.Col = grdSchd.Col + 1
            mSchdEnableBox
        End If
    Else
'        lmTopRow = -1
        grdSchd.TopRow = grdSchd.FixedRows
        grdSchd.Col = LEADDAYINDEX
        grdSchd.Row = grdSchd.FixedRows
        mSchdEnableBox
    End If
End Sub

Private Sub rbcMerge_Click(Index As Integer)
    If rbcMerge(Index).Value Then
        imFieldChgd = True
        mSetCommands
    End If
End Sub

Private Sub rbcPurge_Click(Index As Integer)
    If (rbcPurge(Index).Value) And (Index <> 2) Then
        edcPurgeTime.text = ""
    End If
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub rbcPurgeTest_Click(Index As Integer)
    If (rbcPurgeTest(Index).Value) And (Index <> 2) Then
        edcPurgeTimeTest.text = ""
    End If
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub rbcSchOrAuto_Click(Index As Integer)
    If rbcSchOrAuto(Index).Value Then
        Select Case Index
            Case 0
                edcSchdGenTime.Enabled = True
                edcAutoGenTime.Enabled = False
                If rbcPurge(0).Value = True Then
                    rbcPurge(0).Value = False
                End If
                rbcPurge(0).Enabled = False
                rbcPurge(1).Enabled = True
            Case 1
                edcSchdGenTime.Enabled = False
                edcAutoGenTime.Enabled = True
                If rbcPurge(1).Value = True Then
                    rbcPurge(1).Value = False
                End If
                rbcPurge(1).Enabled = False
                rbcPurge(0).Enabled = True
            Case 2
                edcSchdGenTime.Enabled = True
                edcAutoGenTime.Enabled = True
                rbcPurge(0).Enabled = True
                rbcPurge(1).Enabled = True
        End Select
        imFieldChgd = True
    End If
    mSetCommands
End Sub

Private Sub rbcSchOrAutoTest_Click(Index As Integer)
    If rbcSchOrAutoTest(Index).Value Then
        Select Case Index
            Case 0
                edcSchdGenTimeTest.Enabled = True
                edcAutoGenTimeTest.Enabled = False
                If rbcPurgeTest(0).Value = True Then
                    rbcPurgeTest(0).Value = False
                End If
                rbcPurgeTest(0).Enabled = False
                rbcPurgeTest(1).Enabled = True
            Case 1
                edcSchdGenTimeTest.Enabled = False
                edcAutoGenTimeTest.Enabled = True
                If rbcPurgeTest(1).Value = True Then
                    rbcPurgeTest(1).Value = False
                End If
                rbcPurgeTest(1).Enabled = False
                rbcPurgeTest(0).Enabled = True
            Case 2
                edcSchdGenTimeTest.Enabled = True
                edcAutoGenTimeTest.Enabled = True
                rbcPurgeTest(0).Enabled = True
                rbcPurgeTest(1).Enabled = True
        End Select
        imFieldChgd = True
    End If
    mSetCommands
End Sub

Private Sub tabAuto_Click()
    If imTabIndex = tabAuto.SelectedItem.Index Then
        Exit Sub
    End If
    frcTab(tabAuto.SelectedItem.Index - 1).Visible = True
    frcTab(imTabIndex - 1).Visible = False
    imTabIndex = tabAuto.SelectedItem.Index
End Sub

Private Sub mSetTab()
    'tabAuto.Left = frcSelect.Left
    'tabAuto.Height = cmcCancel.Top - (frcSelect.Top + frcSelect.Height + 300)  'TabAuto.ClientTop - TabAuto.Top + (10 * frcTab(0).Height) / 9
    frcTab(0).Move tabAuto.ClientLeft, tabAuto.ClientTop ', tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(1).Move tabAuto.ClientLeft, tabAuto.ClientTop ', tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(2).Move tabAuto.ClientLeft, tabAuto.ClientTop ', tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(3).Move tabAuto.ClientLeft, tabAuto.ClientTop ', tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(4).Move tabAuto.ClientLeft, tabAuto.ClientTop ', tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    frcTab(3).BorderStyle = 0
    frcTab(4).BorderStyle = 0
End Sub

Private Sub mSchdEnableBox()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(MATERIALTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdSchd.Row >= grdSchd.FixedRows) And (grdSchd.Row < grdSchd.Rows) And (grdSchd.Col >= 0) And (grdSchd.Col < grdSchd.Cols - 1) Then
        lmSchdEnableRow = grdSchd.Row
        lmSchdEnableCol = grdSchd.Col
        imShowGridBox = True
        pbcSchdArrow.Move grdSchd.Left - pbcSchdArrow.Width - 30, grdSchd.Top + grdSchd.RowPos(grdSchd.Row) + (grdSchd.RowHeight(grdSchd.Row) - pbcSchdArrow.Height) / 2
        pbcSchdArrow.Visible = True
        Select Case grdSchd.Col
            Case LOGDAYINDEX  'Call Letters
            Case LEADDAYINDEX  'Date
                edcSchdGrid.Move grdSchd.Left + grdSchd.ColPos(grdSchd.Col) + 30, grdSchd.Top + grdSchd.RowPos(grdSchd.Row) + 15, grdSchd.ColWidth(grdSchd.Col) - 45, grdSchd.RowHeight(grdSchd.Row) - 15
                edcSchdGrid.MaxLength = 2
                edcSchdGrid.text = grdSchd.text
                edcSchdGrid.Visible = True
                edcSchdGrid.SetFocus
        End Select
    End If
End Sub

Private Sub mSchdSetShow()
    If (lmSchdEnableRow >= grdSchd.FixedRows) And (lmSchdEnableRow < grdSchd.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmSchdEnableCol
            Case LOGDAYINDEX
            Case LEADDAYINDEX
        End Select
    End If
    imShowGridBox = False
    pbcSchdArrow.Visible = False
    edcSchdGrid.Visible = False
End Sub

Private Sub mAutoEnableBox()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(MATERIALTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdAuto.Row >= grdAuto.FixedRows) And (grdAuto.Row < grdAuto.Rows) And (grdAuto.Col >= 0) And (grdAuto.Col < grdAuto.Cols - 1) Then
        lmAutoEnableRow = grdAuto.Row
        lmAutoEnableCol = grdAuto.Col
        imShowGridBox = True
        pbcAutoArrow.Move grdAuto.Left - pbcAutoArrow.Width - 30, grdAuto.Top + grdAuto.RowPos(grdAuto.Row) + (grdAuto.RowHeight(grdAuto.Row) - pbcAutoArrow.Height) / 2
        pbcAutoArrow.Visible = True
        Select Case grdAuto.Col
            Case LOGDAYINDEX  'Call Letters
            Case LEADDAYINDEX  'Date
                edcAutoGrid.Move grdAuto.Left + grdAuto.ColPos(grdAuto.Col) + 30, grdAuto.Top + grdAuto.RowPos(grdAuto.Row) + 15, grdAuto.ColWidth(grdAuto.Col) - 45, grdAuto.RowHeight(grdAuto.Row) - 15
                edcAutoGrid.MaxLength = 2
                edcAutoGrid.text = grdAuto.text
                edcAutoGrid.Visible = True
                edcAutoGrid.SetFocus
        End Select
    End If
End Sub
Private Sub mAutoSetShow()
    If (lmAutoEnableRow >= grdAuto.FixedRows) And (lmAutoEnableRow < grdAuto.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmAutoEnableCol
            Case LOGDAYINDEX
            Case LEADDAYINDEX
        End Select
    End If
    imShowGridBox = False
    pbcAutoArrow.Visible = False
    edcAutoGrid.Visible = False
End Sub

