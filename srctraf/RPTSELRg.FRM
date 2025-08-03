VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelRg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remote Invoice Worksheet"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   9510
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
   ScaleHeight     =   5550
   ScaleWidth      =   9510
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8115
      TabIndex        =   34
      Top             =   90
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   17
      Top             =   615
      Width           =   2055
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6675
      Top             =   -180
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
      Left            =   7215
      TabIndex        =   32
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
      Left            =   7575
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox edcCopies 
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
         Height          =   300
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   21
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   22
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox cbcFileType 
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
         Height          =   300
         Left            =   780
         TabIndex        =   24
         Top             =   270
         Width           =   2925
      End
      Begin VB.TextBox edcFileName 
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
         Height          =   300
         Left            =   780
         TabIndex        =   26
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Remote Invoice Worksheet"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3795
      Left            =   15
      TabIndex        =   28
      Top             =   1680
      Width           =   9330
      Begin VB.PictureBox pbcSelC 
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
         Height          =   3360
         Left            =   105
         ScaleHeight     =   3360
         ScaleWidth      =   4920
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4920
         Begin V81TrafficReports.CSI_Calendar CSI_CalEnd 
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Text            =   "8/23/2019"
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
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_calStart 
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Text            =   "8/23/2019"
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
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.CheckBox ckcExclSpotsLackReg 
            Caption         =   "Exclude Spots Lacking Regional Copy"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   435
            Width           =   3495
         End
         Begin VB.CheckBox ckcInclMissed 
            Caption         =   "Include Missed Spots"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   735
            Width           =   2190
         End
         Begin VB.ListBox lbcMultiCntr 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RPTSELRg.frx":0000
            Left            =   3165
            List            =   "RPTSELRg.frx":0007
            Sorted          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2490
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.ListBox lbcCntrCode 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RPTSELRg.frx":0019
            Left            =   3300
            List            =   "RPTSELRg.frx":0020
            Sorted          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   2850
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox edcContract 
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
            Height          =   300
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1050
            Width           =   1215
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   75
            ScaleHeight     =   510
            ScaleWidth      =   4275
            TabIndex        =   35
            Top             =   1455
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Vehicle, Station, Date, Time"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   37
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   3105
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Station, Vehicle, Date, Time"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   720
               TabIndex        =   36
               Top             =   255
               Visible         =   0   'False
               Width           =   3135
            End
         End
         Begin MSComDlg.CommonDialog cdcSetup 
            Left            =   1485
            Top             =   2460
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".Txt"
            Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
            FilterIndex     =   1
            FontSize        =   0
            MaxFileSize     =   256
         End
         Begin VB.Label lacEndDate 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2520
            TabIndex        =   5
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   1095
            Width           =   960
         End
         Begin VB.Label lacStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Dates- Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.PictureBox pbcOption 
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
         Height          =   3570
         Left            =   4920
         ScaleHeight     =   3570
         ScaleWidth      =   4245
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4245
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   2
            ItemData        =   "RPTSELRg.frx":0031
            Left            =   120
            List            =   "RPTSELRg.frx":0033
            MultiSelect     =   2  'Extended
            TabIndex        =   15
            Top             =   2100
            Width           =   4185
         End
         Begin VB.CheckBox ckcAllVehicles 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   135
            TabIndex        =   14
            Top             =   1785
            Width           =   1425
         End
         Begin VB.CheckBox ckcAllContracts 
            Caption         =   "All Contracts"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2295
            TabIndex        =   12
            Top             =   60
            Width           =   1740
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            ItemData        =   "RPTSELRg.frx":0035
            Left            =   120
            List            =   "RPTSELRg.frx":0037
            MultiSelect     =   2  'Extended
            TabIndex        =   11
            Top             =   420
            Width           =   1995
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   0
            ItemData        =   "RPTSELRg.frx":0039
            Left            =   2265
            List            =   "RPTSELRg.frx":003B
            MultiSelect     =   2  'Extended
            TabIndex        =   13
            Top             =   420
            Width           =   1995
         End
         Begin VB.CheckBox ckcAllAdvt 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   60
            Width           =   1740
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   555
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Width           =   360
   End
End
Attribute VB_Name = "RptSelRg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRg.Frm - Producer & Provider Reports
'            1-15-04 DH
'
'
' Release: 5.1
'

Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAllAdvt As Integer 'True=Set list box; False= don't change list box
Dim imAllAdvtClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllCnt As Integer 'True=Set list box; False= don't change list box
Dim imAllCntClicked As Integer  'True=All box clicked (don't call ckcAllCnt within lbcSelection)
Dim imSetAllVeh As Integer 'True=Set list box; False= don't change list box
Dim imAllVehClicked As Integer  'True=All box clicked (don't call ckcAllVehicles within lbcSelection)

Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name

'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
Dim imTerminate As Integer
Dim smPlcSelC4 As String
'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvVirtVehPop             *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVirtVehPop(ilIndex As Integer)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelRg, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelRg
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



Private Sub mCntrPop(slCntrStatus As String, ilHOState As Integer)
'
'   mCntrPop
'   Where:
'       slcntrStatus(I)- O; H; W; C; I; D or blank for all
'       ilHOState(I) - 1 only get cnt (w/o revision) H & O only
'                      2 combo - get latest orders includ revisions (H O G or N) if G or N, show instead of the H or O
'                      3 everything - revision & orders (HOGNWCI) if GNWCI, show over the H or O
'
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim ilCurrent As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim llCntrNo As Long
    Dim ilShow As Integer
    Dim slCntrType As String
    Dim ilAdfCode As Integer
    Dim llLen As Long
    Dim ilErr As Integer
    Dim slShow As String
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slRevNo As String
    llLen = 0
    ilErr = False
    lbcSelection(0).Clear
    lbcCntrCode.Clear
    For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
        If lbcSelection(1).Selected(ilLoop) Then
            sgMultiCntrCodeTag = ""             'init the date stamp so the box will be populated
            ReDim tgMultiCntrCodeCT(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            slNameCode = tgAdvertiser(ilLoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'ilCurrent = 1   '0=Current; 1=All
            'ilFilter = Val(slCode)   'by contract #; -101=by advertiser
            'ilVehCode = -1  'All vehicles
            'ilRet = gPopCntrBox(RptSelCt, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcMultiCntr, lbcMultiCntrCode, True, False, False, False)
            'slCntrType = ""                                 'all Types
            slCntrType = "C"
            If tgUrf(0).sResvType <> "H" Then
                slCntrType = slCntrType & "V"
            End If
            If tgUrf(0).sRemType <> "H" Then
                slCntrType = slCntrType & "T"
            End If
            If tgUrf(0).sDRType <> "H" Then
                slCntrType = slCntrType & "R"
            End If
            If tgUrf(0).sPIType <> "H" Then
                slCntrType = slCntrType & "Q"
            End If
            If tgUrf(0).sPSAType <> "H" Then
                slCntrType = slCntrType & "S"
            End If
            If tgUrf(0).sPromoType <> "H" Then
                slCntrType = slCntrType & "M"
            End If
            If slCntrType = "CVTRQSM" Then
                slCntrType = ""
            End If
            'ilShow = 1
            ilShow = 7                  '10-28-03 show product name instead of advt
            ilCurrent = 1
            ilAdfCode = Val(slCode)
            'load up list box with contracts with matching adv
            'ilRet = gPopCntrForAASBox(RptSelCt, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            ilRet = gPopCntrForAASBox(RptSelRg, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeCT(), sgMultiCntrCodeTagCT)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelRg
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeCT) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeCT(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    ilErr = True
                    Exit For
                End If
                lbcCntrCode.AddItem slName  'lbcMultiCntrCode.List(ilIndex)
            Next ilIndex
            If ilErr Then
                Exit For
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        slNameCode = lbcCntrCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slRevNo)
        ilRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slCode, 2, "-", slRevNo)
        ilExtRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If (ilRevNo > 0) Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        ilRet = gParseItem(slName, 6, "|", slCode)
        slShow = slShow & " " & slCode
        lbcSelection(0).AddItem Trim$(slShow)  'Add ID to list box
    Next ilLoop
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub





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
Private Sub mAdvtPop(lbcSelection As Control)
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCt, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptSelRg, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelRg
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcFileType_Click()
    imComboBoxIndex = cbcFileType.ListIndex
    imFTSelectedIndex = cbcFileType.ListIndex
    mSetCommands
End Sub
Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub
Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcAllAdvt_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllAdvt.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added

Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
Dim ilListIndex As Integer
Dim ilIndex

    ilListIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllAdvt Then
        imAllAdvtClicked = True
        'select all advt
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllAdvtClicked = False
        ckcAllContracts.Visible = False
        ckcAllContracts.Value = vbChecked       'insure all contracts obtained
        lbcSelection(0).Visible = False
        lbcSelection(0).Clear
        sgMultiCntrCodeTagCT = ""       'make sure the contracts are repopulated
        If ilValue = False Then
            ckcAllContracts.Visible = True
            lbcSelection(0).Visible = True
            lbcSelection(0).Clear

        End If
        
    End If
    mSetCommands
End Sub
Private Sub ckcAllAdvt_GotFocus()
    gCtrlGotFocus ckcAllAdvt
End Sub

Private Sub ckcAllContracts_Click()
Dim Value As Integer
    Value = False
    If ckcAllContracts.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added

Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
Dim ilListIndex As Integer
Dim ilIndex

    ilListIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllCnt Then
        imAllCntClicked = True
        'select all advt
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllCntClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllVehicles_Click()
    Dim Value As Integer
    Value = False
    If ckcAllVehicles.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added

Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
Dim ilListIndex As Integer
Dim ilIndex

    ilListIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllVeh Then
        imAllVehClicked = True
        'select all advt
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllVehClicked = False
    End If
    mSetCommands
End Sub

Private Sub cmcBrowse_Click()
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    mSetCommands
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
End Sub
Private Sub cmcBrowse_GotFocus()
    gCtrlGotFocus cmcBrowse
End Sub
Private Sub cmcCancel_Click()
    If igGenRpt Then
        Exit Sub
    End If
    'mTerminate True
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim ilListIndex As Integer
    Dim slRptName As String

    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    'igWhen = frcWhen.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    'igReportType = frcRptType.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    'frcWhen.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    'frcRptType.Enabled = False


    ilListIndex = lbcRptType.ListIndex


    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If ilListIndex = RG_REGION_ASSIGN Then
            slRptName = "RegCopy.rpt"
        Else
            slRptName = "RegCopyTrace.rpt"
        End If
    If Not gOpenPrtJob(Trim$(slRptName)) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gcmcgenRg()
        
        If ilRet <> 1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            'mTerminate
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        
        End If

        Screen.MousePointer = vbHourglass
        'Generate the prepass for the producers or providers
        gGenRegionalCopy

        Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-21-02
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
End Sub
Private Sub cmcGen_GotFocus()
    gCtrlGotFocus cmcGen
End Sub
Private Sub cmcGen_KeyDown(KeyCode As Integer, Shift As Integer)
    imGenShiftKey = Shift
End Sub
Private Sub cmcList_Click()
    If igGenRpt Then
        Exit Sub
    End If
    mTerminate True
End Sub
Private Sub cmcSetup_Click()
    'cdcSetup.flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub CSI_CalEnd_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalEnd_GotFocus()
    gCtrlGotFocus CSI_CalEnd
End Sub

Private Sub CSI_calStart_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_calStart_GotFocus()
    gCtrlGotFocus CSI_calStart
End Sub

Private Sub edcCopies_Change()
    mSetCommands
End Sub
Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
End Sub
Private Sub edcCopies_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

'Private Sub edcEndDate_Change()
'    mSetCommands
'End Sub
'
'Private Sub edcEndDate_GotFocus()
'    gCtrlGotFocus edcEndDate
'End Sub

Private Sub edcFileName_Change()
    mSetCommands
End Sub
Private Sub edcFileName_GotFocus()
    gCtrlGotFocus edcFileName
End Sub
Private Sub edcFileName_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer

    ilPos = InStr(edcFileName.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcFileName.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    If (KeyAscii <= 32) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

'Private Sub edcStartDate_Change()
'    mSetCommands
'End Sub
'
'Private Sub edcStartDate_GotFocus()
'    gCtrlGotFocus edcStartDate
'
'End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Load()
    imTerminate = False
    igGenRpt = False
    mParseCmmdLine
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelRg.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    
    Set RptSelRg = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
Dim ilListIndex As Integer
Dim slCntrStatus As String
Dim ilHOState As Integer
    ilListIndex = lbcRptType.ListIndex

    If Index = 0 Then
        If Not imAllCntClicked Then
            imSetAllCnt = False
            ckcAllContracts.Value = vbUnchecked
            imSetAllCnt = True
        End If
    ElseIf Index = 1 Then
        If Not imAllAdvtClicked Then
            If Index = 1 Then           'advertiser selected
                lbcSelection(0).Visible = True      'show the contract selection box
                ckcAllContracts.Value = vbUnchecked 'insre all contracts selected off
                ckcAllContracts.Visible = True
                slCntrStatus = ""
                
                slCntrStatus = "HO"         'include holds and orders
                ilHOState = 1                   'dealing with spots, only get the scheduled H & O
                mCntrPop slCntrStatus, ilHOState    'fill the contract list box
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
            End If
            imSetAllAdvt = False
            ckcAllAdvt.Value = vbUnchecked  '9-12-02 False
            imSetAllAdvt = True
        End If
    Else
        If Not imAllVehClicked Then
            imSetAllVeh = False
            ckcAllVehicles.Value = vbUnchecked
            imSetAllVeh = True
        End If
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Place focus before populating all lists  *                                                   *
'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim ilLoop As Integer
Dim slStr As String

    Screen.MousePointer = vbHourglass
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If

    RptSelRg.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllAdvtClicked = False
    imSetAllAdvt = True
    
    imAllCntClicked = False
    imSetAllCnt = True
 
    imAllVehClicked = False
    imSetAllVeh = True


    'lbcSelection(0).Move 15, ckcAll.Height + 90, 4380
    'lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(1).Top
    'pbcSelC.Move 90, 255, 4515, 3360
    mAdvtPop lbcSelection(1)
    If imTerminate Then
        Exit Sub
    End If
    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    
    mSellConvVirtVehPop 2
    If imTerminate Then
        Exit Sub
    End If

    gCenterStdAlone RptSelRg
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilRet As Integer
    'cbcWhenDay.AddItem "One Time"
    'cbcWhenDay.AddItem "Every M-F"
    'cbcWhenDay.AddItem "Every M-Sa"
    'cbcWhenDay.AddItem "Every M-Su"
    'cbcWhenDay.AddItem "Every Monday"
    'cbcWhenDay.AddItem "Every Tuesday"
    'cbcWhenDay.AddItem "Every Wednesday"
    'cbcWhenDay.AddItem "Every Thursday"
    'cbcWhenDay.AddItem "Every Friday"
    'cbcWhenDay.AddItem "Every Saturday"
    'cbcWhenDay.AddItem "Every Sunday"
    'cbcWhenDay.AddItem "Cal Month End+1"
    'cbcWhenDay.AddItem "Cal Month End+2"
    'cbcWhenDay.AddItem "Cal Month End+3"
    'cbcWhenDay.AddItem "Cal Month End+4"
    'cbcWhenDay.AddItem "Cal Month End+5"
    'cbcWhenDay.AddItem "Std Month End+1"
    'cbcWhenDay.AddItem "Std Month End+2"
    'cbcWhenDay.AddItem "Std Month End+3"
    'cbcWhenDay.AddItem "Std Month End+4"
    'cbcWhenDay.AddItem "Std Month End+5"
    'cbcWhenDay.ListIndex = 0
    'cbcWhenTime.AddItem "Right Now"
    'cbcWhenTime.AddItem "at 10PM"
    'cbcWhenTime.AddItem "at 12AM"
    'cbcWhenTime.AddItem "at 2AM"
    'cbcWhenTime.AddItem "at 4AM"
    'cbcWhenTime.AddItem "at 6AM"
    'cbcWhenTime.ListIndex = 0
    'Setup report output types
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0

    gPopExportTypes cbcFileType     '2-21-02
    pbcSelC.Visible = False
    lbcRptType.Clear

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True

    lbcRptType.AddItem "Regional Copy Assignment", RG_REGION_ASSIGN
    lbcRptType.AddItem "Regional Copy Tracing", RG_REGION_ASSIGNTRACE

    pbcOption.Visible = True
    pbcSelC.Visible = True
    frcOption.Enabled = True

    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If

    mSetCommands
    Screen.MousePointer = vbDefault
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
    Dim slRptListCmmd As String

    slCommand = sgCommandStr
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
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
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelRg, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Remote Invoice Worksheet"
    '    igRptCallType = -1  'unused in standalone exe, CONTRACTSJOB 'SLSPCOMMSJOB   'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = -1  'unused in standalone exe   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
        End If
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. LeVine      *
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
    Dim ilEnable As Integer
    Dim ilListIndex As Integer
    Dim ilIndex As Integer
    Dim ilHowMany As Integer

    ilListIndex = lbcRptType.ListIndex

    ilEnable = False

'    If RptSelRg!edcStartDate.Text <> "" And RptSelRg!edcEndDate <> "" Then
    If RptSelRg!CSI_calStart.Text <> "" And RptSelRg!CSI_CalEnd.Text <> "" Then      '8-23-19 use csi calendar control vs edit box
        'lbcSelection(1) = advts, lbcSelection(0) = contracts, lbcSelection(2) = vehicles
        'if All advt selected, no contracts are populated
        If lbcSelection(1).SelCount > 0 And (lbcSelection(0).SelCount > 0 Or ckcAllContracts.Value = vbChecked Or ckcAllAdvt.Value = vbChecked) And lbcSelection(2).SelCount > 0 Then      'see if any items selected
            ilEnable = True
        End If
    End If
    If ilEnable Then
        If rbcOutput(0).Value Then  'Display
            ilEnable = True
        ElseIf rbcOutput(1).Value Then  'Print
            If edcCopies.Text <> "" Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        Else    'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If
    End If
    cmcGen.Enabled = ilEnable
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
Private Sub mTerminate(ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'


    If ilFromCancel Then
        igRptReturn = True
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelRg
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    '    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    '    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    '    Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print smPlcSelC4
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of Coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcOutput_GotFocus(Index As Integer)
    If imFirstTime Then
        'mInitDDE
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub






Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
