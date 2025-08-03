VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmSetMG 
   Caption         =   "Set MG"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "AffSetMG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9900
   Begin VB.OptionButton rbcSetMG 
      Caption         =   "Date Range"
      Height          =   195
      Index           =   3
      Left            =   1470
      TabIndex        =   45
      Top             =   330
      Width           =   1695
   End
   Begin VB.Frame frcSetMG 
      Caption         =   "Date Range"
      Height          =   2235
      Index           =   3
      Left            =   1020
      TabIndex        =   40
      Top             =   1545
      Visible         =   0   'False
      Width           =   5985
      Begin V81Affiliate.CSI_Calendar edcStartDate 
         Height          =   285
         Left            =   1245
         TabIndex        =   41
         Top             =   345
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "1/1/2015"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar edcEndDate 
         Height          =   285
         Left            =   4215
         TabIndex        =   42
         Top             =   345
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "1/2/2015"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label lacStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lacEndDate 
         Caption         =   "End Date"
         Height          =   255
         Left            =   3075
         TabIndex        =   43
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.OptionButton rbcSetMG 
      Caption         =   "Station/Vehicles"
      Height          =   210
      Index           =   2
      Left            =   4770
      TabIndex        =   3
      Top             =   75
      Width           =   1695
   End
   Begin VB.Frame frcSetMG 
      Caption         =   "Station/Vehicles"
      Height          =   3360
      Index           =   2
      Left            =   660
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   6030
      Begin V81Affiliate.CSI_Calendar edcSVDate 
         Height          =   285
         Left            =   1290
         TabIndex        =   22
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "11/8/2010"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.TextBox edcSVStation 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Station"
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox edcSVVehicle 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Vehicles"
         Top             =   660
         Width           =   3810
      End
      Begin VB.CheckBox chkSVAllStation 
         Caption         =   "All"
         Height          =   195
         Left            =   135
         TabIndex        =   29
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ListBox lbcSVStation 
         Height          =   2010
         ItemData        =   "AffSetMG.frx":08CA
         Left            =   120
         List            =   "AffSetMG.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   915
         Width           =   1695
      End
      Begin VB.TextBox edcSVWeeks 
         Height          =   285
         Left            =   4740
         TabIndex        =   24
         Text            =   "1"
         Top             =   225
         Width           =   405
      End
      Begin VB.CheckBox chkSVAllVehicles 
         Caption         =   "All"
         Height          =   195
         Left            =   2010
         TabIndex        =   30
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ListBox lbcSVVehicles 
         Height          =   2010
         ItemData        =   "AffSetMG.frx":08CE
         Left            =   1980
         List            =   "AffSetMG.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   915
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lacSVStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lacSVWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3225
         TabIndex        =   23
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.Frame frcSetMG 
      Caption         =   "Agreement Code"
      Height          =   2520
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   765
      Width           =   5985
      Begin V81Affiliate.CSI_Calendar edcACDate 
         Height          =   285
         Left            =   1245
         TabIndex        =   8
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "11/8/2010"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.TextBox edcAC 
         Height          =   285
         Left            =   1635
         TabIndex        =   6
         Top             =   225
         Width           =   990
      End
      Begin VB.TextBox edcACWeeks 
         Height          =   285
         Left            =   4695
         TabIndex        =   10
         Text            =   "1"
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lacAC 
         Caption         =   "Agreement Code"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lacACWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3180
         TabIndex        =   9
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label lacACStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   1215
      End
   End
   Begin VB.OptionButton rbcSetMG 
      Caption         =   "Vehicle/Stations"
      Height          =   210
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   75
      Width           =   1695
   End
   Begin VB.OptionButton rbcSetMG 
      Caption         =   "Agreement Code"
      Height          =   210
      Index           =   0
      Left            =   1470
      TabIndex        =   1
      Top             =   75
      Width           =   1605
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   3840
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2985
      ItemData        =   "AffSetMG.frx":08D2
      Left            =   6585
      List            =   "AffSetMG.frx":08D4
      TabIndex        =   36
      Top             =   660
      Width           =   2820
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   -75
      Top             =   2460
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4935
      FormDesignWidth =   9900
   End
   Begin VB.CommandButton cmcSetMG 
      Caption         =   "&Set MG"
      Height          =   375
      Left            =   5910
      TabIndex        =   33
      Top             =   4290
      Width           =   1665
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   34
      Top             =   4290
      Width           =   1665
   End
   Begin VB.Frame frcSetMG 
      Caption         =   "Vehicle/Stations"
      Height          =   3360
      Index           =   1
      Left            =   285
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   6030
      Begin V81Affiliate.CSI_Calendar edcVSDate 
         Height          =   285
         Left            =   1290
         TabIndex        =   13
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "11/8/2010"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.ListBox lbcVSVehicles 
         Height          =   2010
         ItemData        =   "AffSetMG.frx":08D6
         Left            =   120
         List            =   "AffSetMG.frx":08D8
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   915
         Width           =   3855
      End
      Begin VB.CheckBox chkVSAllVehicles 
         Caption         =   "All"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox edcVSWeeks 
         Height          =   285
         Left            =   4740
         TabIndex        =   15
         Text            =   "1"
         Top             =   225
         Width           =   405
      End
      Begin VB.ListBox lbcVSStation 
         Height          =   2010
         ItemData        =   "AffSetMG.frx":08DA
         Left            =   4230
         List            =   "AffSetMG.frx":08DC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   915
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkVSAllStation 
         Caption         =   "All"
         Height          =   195
         Left            =   4245
         TabIndex        =   32
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox edcTitle1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Vehicle"
         Top             =   675
         Width           =   3810
      End
      Begin VB.TextBox edcTitle3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4305
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Stations"
         Top             =   675
         Width           =   1635
      End
      Begin VB.Label lacVSWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3225
         TabIndex        =   14
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lacVSStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Label lacStartTime 
      Height          =   195
      Left            =   5955
      TabIndex        =   39
      Top             =   4680
      Width           =   3345
   End
   Begin VB.Label lacCounts 
      Height          =   195
      Left            =   5940
      TabIndex        =   38
      Top             =   3975
      Width           =   3345
   End
   Begin VB.Label lacSetMG 
      Caption         =   "Set MG's by"
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   135
      TabIndex        =   37
      Top             =   4215
      Width           =   5580
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6915
      TabIndex        =   35
      Top             =   390
      Width           =   1965
   End
End
Attribute VB_Name = "frmSetMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private smDate As String     'Export Date
Private imWeeks As Integer
Private smEndDate As String
Private imVefCode As Integer
Private smVefName As String
Private imShttCode As Integer
Private lmAttCode As Long
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imAllVehicleClick As Integer
Private imSettingMG As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmFrom As Integer
Private cprst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private AgreementInfo_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private dat_rst As ADODB.Recordset
Private lmTotalProcessedCount As Long
Private lmTotalToProcess As Long
Private smSvLogActivityInto As String
Private tmDat As DAT




'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
    ilRet = 0
    slNowDate = Format$(gNow(), sgShowDateForm)
    slToFile = sgMsgDirectory & "SetMGAffiliateSpots.Txt"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Set MG Affiliate Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mVSFillVehicle()
    Dim iLoop As Integer
    lbcVSVehicles.Clear
    lbcMsg.Clear
    chkVSAllVehicles.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVSVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVSVehicles.ItemData(lbcVSVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
End Sub

Private Sub mSVFillStation()
    Dim iLoop As Integer
    lbcSVStation.Clear
    lbcMsg.Clear
    chkVSAllVehicles.Value = 0
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        lbcSVStation.AddItem Trim$(tgStationInfo(iLoop).sCallLetters)
        lbcSVStation.ItemData(lbcSVStation.NewIndex) = tgStationInfo(iLoop).iCode
    Next iLoop
End Sub

Private Sub chkSVAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkSVAllStation.Value = vbChecked Then
        iValue = True
        If lbcVSStation.ListCount > 1 Then
            edcSVVehicle.Visible = False
            chkSVAllVehicles.Visible = False
            lbcSVVehicles.Visible = False
            lbcSVVehicles.Clear
        Else
            edcSVVehicle.Visible = True
            chkSVAllVehicles.Visible = True
            lbcSVVehicles.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcSVStation.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcSVStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSVStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
End Sub

Private Sub chkSVAllVehicles_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllVehicleClick Then
        Exit Sub
    End If
    If chkSVAllVehicles.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcSVVehicles.ListCount > 0 Then
        imAllVehicleClick = True
        lRg = CLng(lbcSVVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSVVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllVehicleClick = False
    End If
End Sub

Private Sub chkVSAllVehicles_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkVSAllVehicles.Value = vbChecked Then
        iValue = True
        If lbcVSVehicles.ListCount > 1 Then
            edcTitle3.Visible = False
            chkVSAllStation.Visible = False
            lbcVSStation.Visible = False
            lbcVSStation.Clear
        Else
            edcTitle3.Visible = True
            chkVSAllStation.Visible = True
            lbcVSStation.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcVSVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVSVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVSVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub chkVSAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllStationClick Then
        Exit Sub
    End If
    If chkVSAllStation.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVSStation.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcVSStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVSStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If

End Sub


Private Sub cmcSetMG_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim slMoDate As String
    Dim llMoDate As Long
    Dim sNowDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim ilShtt As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim slStr As String
    Dim llAttCode As Long
    Dim slVehicles As String
    Dim slstations As String
    Dim llPrevAtt As Long
    Dim llPrevDate As Long
    Dim slSQLQuery As String

    On Error GoTo ErrHand
    
    If imSettingMG Then
        Exit Sub
    End If
    imSettingMG = True
    lacCounts.Caption = ""
    lacResult.Caption = ""
    lacStartTime.Caption = ""
    lbcMsg.Clear
    mCloseAgreementInfo
    Set AgreementInfo_rst = mInitAgreementInfo()
    llPrevAtt = -1
    llUpper = 0
    slVehicles = ""
    slstations = ""
    If rbcSetMG(0).Value Then
        If edcAC.Text = "" Then
            gMsgBox "Agreement Code must be specified.", vbOKOnly
            edcACDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        If edcACDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcACDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        If gIsDate(edcACDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcACDate.SetFocus
            imSettingMG = False
            Exit Sub
        Else
            smDate = Format(edcACDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcACDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcACWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcACWeeks.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        llEDate = DateValue(gAdjYear(Format$(DateAdd("ww", imWeeks - 1, slMoDate), "mm/dd/yy")))
        lacResult.Caption = "Gathering Agreement Information"
        For llDate = llSDate To llEDate Step 7
            llAttCode = Val(edcAC.Text)
            llMoDate = gDateValue(gObtainPrevMonday(Format(llDate, "m/d/yy")))
            mAddAgreementInfo llAttCode, llMoDate
        Next llDate
        lacResult.Caption = ""
    ElseIf rbcSetMG(1).Value Then
        If lbcVSVehicles.ListIndex < 0 Then
            imSettingMG = False
            Exit Sub
        End If
        If edcVSDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcVSDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        If gIsDate(edcVSDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcVSDate.SetFocus
            imSettingMG = False
            Exit Sub
        Else
            smDate = Format(edcVSDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcVSDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcVSWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcVSWeeks.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        llEDate = DateValue(gAdjYear(Format$(DateAdd("ww", imWeeks - 1, slMoDate), "mm/dd/yy")))
        lacResult.Caption = "Gathering Agreement Information"
        DoEvents
        For ilVef = 0 To lbcVSVehicles.ListCount - 1
            If lbcVSVehicles.Selected(ilVef) Then
                If slVehicles = "" Then
                    slVehicles = lbcVSVehicles.List(ilVef)
                Else
                    slVehicles = slVehicles & ", " & lbcVSVehicles.List(ilVef)
                End If
                imVefCode = lbcVSVehicles.ItemData(ilVef)
                For ilShtt = 0 To lbcVSStation.ListCount - 1
                    If lbcVSStation.Selected(ilShtt) Then
                        If slstations = "" Then
                            slstations = lbcVSStation.List(ilShtt)
                        Else
                            slstations = slstations & ", " & lbcVSStation.List(ilShtt)
                        End If
                        imShttCode = lbcVSStation.ItemData(ilShtt)
                        For llDate = llSDate To llEDate Step 7
                            slDate = Format(llDate, "m/d/yy")
                            SQLQuery = "SELECT attCode"
                            SQLQuery = SQLQuery + " FROM att"
                            SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
                            SQLQuery = SQLQuery + " AND attShfCode = " & imShttCode
                            SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')"
                            SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')" & ")"
                            Set att_rst = gSQLSelectCall(SQLQuery)
                            If Not att_rst.EOF Then
                                llAttCode = att_rst!attCode
                                llMoDate = gDateValue(gObtainPrevMonday(slDate))
                                mAddAgreementInfo llAttCode, llMoDate
                            End If
                        Next llDate
                    End If
                Next ilShtt
            End If
        Next ilVef
        lacResult.Caption = ""
    ElseIf rbcSetMG(2).Value Then
        If lbcSVStation.ListIndex < 0 Then
            imSettingMG = False
            Exit Sub
        End If
        If edcSVDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcSVDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        If gIsDate(edcSVDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcSVDate.SetFocus
            imSettingMG = False
            Exit Sub
        Else
            smDate = Format(edcSVDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcSVDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcSVWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcSVWeeks.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        llEDate = DateValue(gAdjYear(Format$(DateAdd("ww", imWeeks - 1, slMoDate), "mm/dd/yy")))
        lacResult.Caption = "Gathering Agreement Information"
        DoEvents
        For ilShtt = 0 To lbcSVStation.ListCount - 1
            If lbcSVStation.Selected(ilShtt) Then
                If slstations = "" Then
                    slstations = lbcSVStation.List(ilShtt)
                Else
                    slstations = slstations & "," & lbcSVStation.List(ilShtt)
                End If
                imShttCode = lbcSVStation.ItemData(ilShtt)
                For ilVef = 0 To lbcSVVehicles.ListCount - 1
                    If lbcSVVehicles.Selected(ilVef) Then
                        If slVehicles = "" Then
                            slVehicles = lbcSVVehicles.List(ilVef)
                        Else
                            slVehicles = slVehicles & ", " & lbcSVVehicles.List(ilVef)
                        End If
                        imVefCode = lbcSVVehicles.ItemData(ilVef)
                        For llDate = llSDate To llEDate Step 7
                            slDate = Format(llDate, "m/d/yy")
                            SQLQuery = "SELECT attCode"
                            SQLQuery = SQLQuery + " FROM att"
                            SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
                            SQLQuery = SQLQuery + " AND attShfCode = " & imShttCode
                            SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')"
                            SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')" & ")"
                            Set att_rst = gSQLSelectCall(SQLQuery)
                            If Not att_rst.EOF Then
                                llAttCode = att_rst!attCode
                                llMoDate = gDateValue(gObtainPrevMonday(slDate))
                                mAddAgreementInfo llAttCode, llMoDate
                            End If
                        Next llDate
                    End If
                Next ilVef
            End If
        Next ilShtt
        lacResult.Caption = ""
    ElseIf rbcSetMG(3).Value Then
        If edcStartDate.Text <> "" Then
            If gIsDate(edcStartDate.Text) = False Then
                Beep
                gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
                edcStartDate.SetFocus
                imSettingMG = False
                Exit Sub
            Else
                smDate = gObtainPrevMonday(Format(edcStartDate.Text, sgShowDateForm))
            End If
        Else
            smDate = ""
        End If
        If edcEndDate.Text <> "" Then
            If gIsDate(edcEndDate.Text) = False Then
                Beep
                gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
                edcEndDate.SetFocus
                imSettingMG = False
                Exit Sub
            Else
                smEndDate = gObtainNextSunday(Format(edcEndDate.Text, sgShowDateForm))
            End If
        Else
            smEndDate = ""
        End If
        ilRet = 0
        On Error GoTo cmcSetMGErr:
        SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast where "
        SQLQuery = SQLQuery & "dateadd(Day,CASE datepart(weekday,astFeedDate) WHEN 1 THEN -6 ELSE datediff(Day,datepart(weekday,astFeedDate),2) END,astFeedDate) <> dateadd(Day,CASE datepart(weekday,astAirDate) WHEN 1 THEN -6 ELSE datediff(Day,datepart(weekday,astAirDate),2) END,astAirDate) And "
        If (smDate = "") And (smEndDate = "") Then
            'SQLQuery = SQLQuery & "astCPStatus = 1 and Mod(astStatus, 100) <> 8 and Mod(astStatus, 100) <> 4 and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
            SQLQuery = SQLQuery & "astCPStatus = 1 and Mod(astStatus, 100) Not In (4, 8, 14) and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
        ElseIf (smDate <> "") And (smEndDate = "") Then
            'SQLQuery = SQLQuery & "astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) <> 8 and Mod(astStatus, 100) <> 4 and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
            SQLQuery = SQLQuery & "astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) Not In (4, 8, 14) and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
        Else
            'SQLQuery = SQLQuery & "astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astFeedDate <= '" & Format(smEndDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) <> 8 and Mod(astStatus, 100) <> 4 and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
            SQLQuery = SQLQuery & "astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astFeedDate <= '" & Format(smEndDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) Not In (4, 8, 14) and Mod(astStatus, 100) <= 10 and astLkAstCode = 0 order by astatfcode, astfeeddate"
        End If
        lacResult.Caption = "Gathering Agreement Information"
        DoEvents
        Set cprst = gSQLSelectCall(SQLQuery)
        If iRet <> 0 Then
            lacResult.Caption = ""
            gMsgBox "SQL Call Structure in Error: " & SQLQuery, vbOKOnly
            edcACDate.SetFocus
            imSettingMG = False
            Exit Sub
        End If
        If cprst.EOF Then
            lacResult.Caption = ""
            gMsgBox "No Records return by the SQL Call", vbOKOnly
            imSettingMG = False
            Exit Sub
        End If
        On Error GoTo ErrHand
        Screen.MousePointer = vbHourglass
        On Error GoTo cmcSetMGErr:
        Do While Not cprst.EOF
            llAttCode = cprst!astAtfCode
            llMoDate = gDateValue(gObtainPrevMonday(Format(cprst!astFeedDate, "m/d/yy")))
            If (llPrevAtt <> llAttCode) Or (llPrevDate <> llMoDate) Then
                llPrevAtt = llAttCode
                llPrevDate = llMoDate
                mAddAgreementInfo llAttCode, llMoDate
            End If
            cprst.MoveNext
        Loop
        On Error GoTo ErrHand
        lacResult.Caption = ""
    Else
        Beep
        gMsgBox "'Set MG by' must be specified", vbCritical
        imSettingMG = False
        Exit Sub
    End If
    lacStartTime.Caption = Now
    AgreementInfo_rst.Filter = adFilterNone
    lmTotalToProcess = AgreementInfo_rst.RecordCount
    If lmTotalToProcess <= 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "No Agreements found to be processed.", vbOKOnly
        imSettingMG = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If Not mOpenMsgFile(sMsgFileName) Then
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        imSettingMG = False
        Exit Sub
    End If
    
    If rbcSetMG(0).Value Or rbcSetMG(1).Value Or rbcSetMG(2).Value Then
        llSDate = DateValue(gAdjYear(slMoDate))
        llEDate = DateValue(gAdjYear(Format$(DateAdd("d", 7 * (imWeeks - 1) + 6, slMoDate), "mm/dd/yy")))
    End If
    If rbcSetMG(0).Value Then
        Print #hmMsg, "Set MG by Agreement Code"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Agreement Code: " & Trim$(edcAC.Text)
    ElseIf rbcSetMG(1).Value Then
        Print #hmMsg, "Set MG by Vehicle/Station"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Vehicle: " & slVehicles
        Print #hmMsg, "  Stations: " & slstations
    ElseIf rbcSetMG(2).Value Then
        Print #hmMsg, "Set MG by Station/Vehicle"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Station: " & slstations
        Print #hmMsg, "  Vehicles: " & slVehicles
    ElseIf rbcSetMG(3).Value Then
        Print #hmMsg, "Set MG by Date Range"
        If (smDate = "") And (smEndDate = "") Then
            Print #hmMsg, "  Date Range: All Dates"
        ElseIf (smDate <> "") And (smEndDate = "") Then
            Print #hmMsg, "  Date Range: " & Format(smDate, "m/d/yy") & "- TFN"
        Else
            Print #hmMsg, "  Date Range: " & Format(smDate, "m/d/yy") & "-" & Format(smEndDate, "m/d/yy")
        End If
    End If
    On Error GoTo 0
    
    Screen.MousePointer = vbHourglass
    iRet = mSetMG()
    If (iRet = False) Then
        'Stop the Pervasive API engine
        Print #hmMsg, "** Terminated **"
        Close #hmMsg
        imSettingMG = False
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        'Stop the Pervasive API engine
        Print #hmMsg, "** User Terminated **"
        Close #hmMsg
        imSettingMG = False
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    imSettingMG = False
    Print #hmMsg, "** Completed Set MG Affiliate Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacStartTime.Caption = lacStartTime.Caption & " to " & Now
    lacResult.Caption = "Results: " & sMsgFileName
    cmcSetMG.Enabled = False
    cmcCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmcSetMGErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-cmcSetMG"
    Exit Sub
End Sub

Private Sub cmcCancel_Click()
    If imSettingMG Then
        imTerminate = True
        Exit Sub
    End If
    edcVSDate.Text = ""
    Unload frmSetMG
End Sub


Private Sub edcAC_Change()
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcACDate_CalendarChanged()
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcACWeeks_Change()
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub



Private Sub edcEndDate_Change()
    lbcMsg.Clear
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcStartDate_Change()
    lbcMsg.Clear
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcVSDate_Change()
    lbcMsg.Clear
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    Dim llEqtCode As Long
    
    If imFirstTime Then
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    smSvLogActivityInto = sgLogActivityInto
    sgLogActivityInto = ""
    imAllClick = False
    imAllStationClick = False
    imAllVehicleClick = False
    imTerminate = False
    imSettingMG = False
    imFirstTime = True
    
    frcSetMG(0).Move lacSetMG.Left - 15, rbcSetMG(3).Top + rbcSetMG(3).Height + 90
    frcSetMG(1).Move frcSetMG(0).Left, frcSetMG(0).Top
    frcSetMG(2).Move frcSetMG(0).Left, frcSetMG(0).Top
    frcSetMG(3).Move frcSetMG(0).Left, frcSetMG(0).Top
    edcACDate.Text = ""
    edcSVDate.Text = ""
    edcVSDate.Text = ""
    edcStartDate.Text = ""
    edcEndDate.Text = ""
    lbcVSStation.Clear
    mVSFillVehicle
    lbcSVVehicles.Clear
    mSVFillStation

    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If imSettingMG Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    sgLogActivityInto = smSvLogActivityInto
    mCloseAgreementInfo
    cprst.Close
    att_rst.Close
    ast_rst.Close
    lst_rst.Close
    dat_rst.Close
    Set frmSetMG = Nothing
End Sub


Private Sub lbcSVStation_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcSVVehicles.Clear
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
    If chkSVAllVehicles.Value = vbChecked Then
        chkSVAllVehicles.Value = vbUnchecked
    End If
    If imAllClick Then
        Exit Sub
    End If
    If chkSVAllStation.Value = vbChecked Then
        imAllClick = True
        chkSVAllStation.Value = vbUnchecked
        imAllClick = False
    End If
    For iLoop = 0 To lbcSVStation.ListCount - 1 Step 1
        If lbcSVStation.Selected(iLoop) Then
            imShttCode = lbcSVStation.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    If iCount = 1 Then
        edcSVVehicle.Visible = True
        chkSVAllVehicles.Visible = True
        lbcSVVehicles.Visible = True
        mSVFillVehicle
    Else
        edcSVVehicle.Visible = False
        chkSVAllVehicles.Visible = False
        lbcSVVehicles.Visible = False
    End If

End Sub

Private Sub lbcSVVehicles_Click()
    If imAllVehicleClick Then
        Exit Sub
    End If
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
    If chkSVAllVehicles.Value = vbChecked Then
        imAllVehicleClick = True
        chkSVAllVehicles.Value = vbUnchecked
        imAllVehicleClick = False
    End If

End Sub

Private Sub lbcVSStation_Click()
    If imAllStationClick Then
        Exit Sub
    End If
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
    If chkVSAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkVSAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
End Sub

Private Sub lbcVSVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcVSStation.Clear
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
    If chkVSAllStation.Value = vbChecked Then
        chkVSAllStation.Value = vbUnchecked
    End If
    If imAllClick Then
        Exit Sub
    End If
    If chkVSAllVehicles.Value = vbChecked Then
        imAllClick = True
        chkVSAllVehicles.Value = vbUnchecked
        imAllClick = False
    End If
    For iLoop = 0 To lbcVSVehicles.ListCount - 1 Step 1
        If lbcVSVehicles.Selected(iLoop) Then
            imVefCode = lbcVSVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    If iCount = 1 Then
        edcTitle3.Visible = True
        chkVSAllStation.Visible = True
        lbcVSStation.Visible = True
        mVSFillStation
    Else
        edcTitle3.Visible = False
        chkVSAllStation.Visible = False
        lbcVSStation.Visible = False
    End If
End Sub

Private Function mSetMG() As Integer
    Dim ilRet As Integer
    Dim llAttCode As Long
    Dim llMoDate As Long
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slPdDate As String
    Dim slPdTime As String
    Dim slAirDate As String
    Dim slFdDate As String
    Dim slFdTime As String
    Dim slSelect As String
    Dim slSQLQuery As String
    Dim ilDay As Integer
    Dim ilPledged As Integer
    Dim ilPdDay As Integer
    Dim ilFdDay As Integer
    Dim ilAdjDay As Integer
    Dim blCreateMG As Boolean
    Dim llMGLstCode As Long
    Dim llMGAstCode As Long
    Dim slVehicleName As String
    Dim slCallLetters As String
    Dim ilAddedMGCount As Integer
    Dim llVef As Long
    Dim llShf As Long
        
    On Error GoTo ErrHand
    lmTotalProcessedCount = 0
    slSelect = "Select * from ast where "
    slSelect = slSelect & "dateadd(Day,CASE datepart(weekday,astFeedDate) WHEN 1 THEN -6 ELSE datediff(Day,datepart(weekday,astFeedDate),2) END,astFeedDate) <> dateadd(Day,CASE datepart(weekday,astAirDate) WHEN 1 THEN -6 ELSE datediff(Day,datepart(weekday,astAirDate),2) END,astAirDate) And "
    On Error Resume Next
    AgreementInfo_rst.Filter = adFilterNone
    If Not (AgreementInfo_rst.EOF And AgreementInfo_rst.BOF) Then
        AgreementInfo_rst.MoveFirst
    End If
    'one record for each agreement/week
    Do While Not AgreementInfo_rst.EOF
        If imTerminate Then
            mSetMG = True
            Exit Function
        End If
        ilAddedMGCount = 0
        llAttCode = AgreementInfo_rst!attCode
        llMoDate = AgreementInfo_rst!MoDate
        slMoDate = Format(llMoDate, "yyyy-mm-dd")
        slSuDate = Format(llMoDate + 6, "yyyy-mm-dd")
        SQLQuery = slSelect & "astAtfCode = " & llAttCode & " And "
        'SQLQuery = SQLQuery & "astFeedDate >= '" & slMoDate & "' and astFeedDate <= '" & slSuDate & "' and astCPStatus = 1 and Mod(astStatus, 100) <> 8 and Mod(astStatus, 100) <> 4 and Mod(astStatus, 100) <= 10 and astLkAstCode = 0"
        SQLQuery = SQLQuery & "astFeedDate >= '" & slMoDate & "' and astFeedDate <= '" & slSuDate & "' and astCPStatus = 1 and Mod(astStatus, 100) Not In (4, 8, 14) and Mod(astStatus, 100) <= 10 and astLkAstCode = 0"
        DoEvents
        Set ast_rst = gSQLSelectCall(SQLQuery)
        If Not ast_rst.EOF Then
            llVef = gBinarySearchVef(CLng(ast_rst!astVefCode))
            If llVef <> -1 Then
                slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
            End If
            llShf = gBinarySearchStationInfoByCode(ast_rst!astShfCode)
            If llShf <> -1 Then
                slCallLetters = Trim$(tgStationInfoByCode(llShf).sCallLetters)
            End If
            lacResult.Caption = "Processing: " & slVehicleName & " " & slCallLetters
            Print #hmMsg, "    Processing: Vehicle- " & slVehicleName & " Station- " & slCallLetters
        End If
        Do While Not ast_rst.EOF
            blCreateMG = False
            slFdDate = Format(ast_rst!astFeedDate, sgShowDateForm)
            slFdTime = Format(ast_rst!astFeedTime, sgShowTimeWSecForm)
            slAirDate = Format(ast_rst!astAirDate, sgShowDateForm)
            blCreateMG = gDeterminePledgeDateTime(dat_rst, ast_rst!astDatCode, slFdDate, slFdTime, slAirDate, slPdDate, slPdTime)
            If blCreateMG Then
                'Create the MG spots
                llMGLstCode = mAddMGLst()
                If llMGLstCode > 0 Then
                    llMGAstCode = mAddAstMG(llMGLstCode)
                    If llMGAstCode > 0 Then
                        mChgAstToMissed llMGAstCode, slPdDate, slPdTime
                        ilAddedMGCount = ilAddedMGCount + 1
                    Else
                        Print #hmMsg, "      " & "Unable to create MG Affiliate Spot (AST) for astCode " & ast_rst!astCode & " Feed Date " & slFdDate & " Feed Time " & slFdTime & " Air Date " & slAirDate & " MG not Created"
                    End If
                Else
                    Print #hmMsg, "      " & "Unable to create MG Affiliate Log Spot (LST) for  astCode " & ast_rst!astCode & " Feed Date " & slFdDate & " Feed Time " & slFdTime & " Air Date " & slAirDate & " MG not Created"
                End If
            End If
            ast_rst.MoveNext
        Loop
        If ilAddedMGCount > 0 Then
            Print #hmMsg, "      " & Format(slMoDate, "m/d/yy") & "-" & Format(slSuDate, "m/d/yy") & " " & ilAddedMGCount & " MG's created"
        End If
        lmTotalProcessedCount = lmTotalProcessedCount + 1
        lacCounts.Caption = "Processed: " & lmTotalProcessedCount & " of " & lmTotalToProcess
        AgreementInfo_rst.MoveNext
    Loop
    mSetMG = True
    Exit Function
mSetMGErr:
    ilRet = Err
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mSetMG"
    mSetMG = False
    Exit Function
    
End Function

Private Sub mVSFillStation()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    chkVSAllStation.Value = vbUnchecked
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery + " FROM shtt, att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
    'SQLQuery = SQLQuery + " AND attExportType = 2 "
    SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcVSStation.AddItem Trim$(rst!shttCallLetters)
        lbcVSStation.ItemData(lbcVSStation.NewIndex) = rst!shttCode
        rst.MoveNext
    Wend
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mVSFillStation"

End Sub

Private Sub mSVFillVehicle()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    chkSVAllVehicles.Value = vbUnchecked
    SQLQuery = "SELECT DISTINCT vefName, vefCode"
    SQLQuery = SQLQuery + " FROM vef_Vehicles, att"
    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode
    SQLQuery = SQLQuery + " AND vefCode = attVefCode)"
    SQLQuery = SQLQuery + " ORDER BY vefName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcSVVehicles.AddItem Trim$(rst!vefName)
        lbcSVVehicles.ItemData(lbcSVVehicles.NewIndex) = rst!vefCode
        rst.MoveNext
    Wend
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mSVFillVehicle"

End Sub
Private Sub rbcSetMG_Click(Index As Integer)
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
    frcSetMG(0).Visible = False
    frcSetMG(1).Visible = False
    frcSetMG(2).Visible = False
    frcSetMG(3).Visible = False
    frcSetMG(Index).Visible = True
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmSetMG
End Sub

Private Sub edcVSWeeks_Change()
    If cmcSetMG.Enabled = False Then
        mClearControls
    End If
End Sub

Private Function mInitAgreementInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "attCode", adInteger
        .Append "MoDate", adInteger
        .Append "SpotAired", adBoolean
    End With
    rst.Open
    rst!attCode.Properties("optimize") = True
    rst.Sort = "attCode,MoDate"
    Set mInitAgreementInfo = rst
End Function
Private Sub mCloseAgreementInfo()
    On Error Resume Next
    If Not AgreementInfo_rst Is Nothing Then
        If (AgreementInfo_rst.State And adStateOpen) <> 0 Then
            AgreementInfo_rst.Close
        End If
        Set AgreementInfo_rst = Nothing
    End If
End Sub

Private Function mAddAgreementInfo(llAttCode As Long, llMoDate As Long) As Integer
    mAddAgreementInfo = False
    AgreementInfo_rst.Filter = "attCode = " & llAttCode & " And MoDate = " & llMoDate
    If AgreementInfo_rst.EOF Then
        AgreementInfo_rst.AddNew Array("attCode", "MoDate"), Array(llAttCode, llMoDate)
        mAddAgreementInfo = True
    End If
End Function



Private Sub mClearControls()
    lacStartTime.Caption = ""
    lacResult.Caption = ""
    lacCounts.Caption = ""
    cmcSetMG.Enabled = True
    cmcCancel.Caption = "&Cancel"
    lbcMsg.Clear
End Sub



Private Function mAddMGLst() As Long
    Dim llLst As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT * From lst Where lstCode = " & ast_rst!astLsfCode
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If lst_rst.EOF Then
        mAddMGLst = 0
        Exit Function
    End If
    
    SQLQuery = "Insert Into lst ( "
    SQLQuery = SQLQuery & "lstCode, "
    SQLQuery = SQLQuery & "lstType, "
    SQLQuery = SQLQuery & "lstSdfCode, "
    SQLQuery = SQLQuery & "lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, "
    SQLQuery = SQLQuery & "lstAgfCode, "
    SQLQuery = SQLQuery & "lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, "
    SQLQuery = SQLQuery & "lstLnVefCode, "
    SQLQuery = SQLQuery & "lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, "
    SQLQuery = SQLQuery & "lstMon, "
    SQLQuery = SQLQuery & "lstTue, "
    SQLQuery = SQLQuery & "lstWed, "
    SQLQuery = SQLQuery & "lstThu, "
    SQLQuery = SQLQuery & "lstFri, "
    SQLQuery = SQLQuery & "lstSat, "
    SQLQuery = SQLQuery & "lstSun, "
    SQLQuery = SQLQuery & "lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, "
    SQLQuery = SQLQuery & "lstPrice, "
    SQLQuery = SQLQuery & "lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, "
    SQLQuery = SQLQuery & "lstLogDate, "
    SQLQuery = SQLQuery & "lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, "
    SQLQuery = SQLQuery & "lstAud, "
    SQLQuery = SQLQuery & "lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, "
    SQLQuery = SQLQuery & "lstBreakNo, "
    SQLQuery = SQLQuery & "lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, "
    SQLQuery = SQLQuery & "lstZone, "
    SQLQuery = SQLQuery & "lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, "
    SQLQuery = SQLQuery & "lstCrfCsfCode, "
    SQLQuery = SQLQuery & "lstStatus, "
    SQLQuery = SQLQuery & "lstLen, "
    SQLQuery = SQLQuery & "lstUnits, "
    SQLQuery = SQLQuery & "lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, "
    SQLQuery = SQLQuery & "lstEvtIDCefCode, "
    SQLQuery = SQLQuery & "lstSplitNetwork, "
    SQLQuery = SQLQuery & "lstRafCode, "
    SQLQuery = SQLQuery & "lstFsfCode, "
    SQLQuery = SQLQuery & "lstGsfCode, "
    SQLQuery = SQLQuery & "lstImportedSpot, "
    SQLQuery = SQLQuery & "lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, "
    SQLQuery = SQLQuery & "lstLnEndTime, "
    SQLQuery = SQLQuery & "lstUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & 2 & ", "    'lstType
    SQLQuery = SQLQuery & lst_rst!lstSdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCntrNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstAdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstProd) & "', "
    SQLQuery = SQLQuery & lst_rst!lstLineNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstLnVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirDate, sgSQLDateForm) & "', "      'lstStartDate
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirDate, sgSQLDateForm) & "', "      'lstEndDate
    SQLQuery = SQLQuery & 0 & ", " 'lstMon
    SQLQuery = SQLQuery & 0 & ", " 'lstTue
    SQLQuery = SQLQuery & 0 & ", " 'lstWed
    SQLQuery = SQLQuery & 0 & ", " 'lstThu
    SQLQuery = SQLQuery & 0 & ", " 'lstFri
    SQLQuery = SQLQuery & 0 & ", " 'lstSat
    SQLQuery = SQLQuery & 0 & ", " 'lstSun
    SQLQuery = SQLQuery & 0 & ", " 'lstSpotsWk
    SQLQuery = SQLQuery & lst_rst!lstPriceType & ", "   'lstPriceType
    SQLQuery = SQLQuery & 0 & ", "   'lstPrice
    SQLQuery = SQLQuery & 5 & ", "    'lstSpotType
    SQLQuery = SQLQuery & lst_rst!lstLogVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("0") & "', "  'lstDemo
    SQLQuery = SQLQuery & 0 & ", " 'lstAud
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstISCI) & "', "
    SQLQuery = SQLQuery & 0 & ", "    'lstWkNo
    SQLQuery = SQLQuery & 0 & ", " 'lstBreakNo
    SQLQuery = SQLQuery & 0 & ", "  'lstPositionNo
    SQLQuery = SQLQuery & 0 & ", "   'lstSeqNo
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstCart) & "', "
    SQLQuery = SQLQuery & lst_rst!lstCpfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCrfCsfCode & ", "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", "  'lstStatus
    SQLQuery = SQLQuery & lst_rst!lstLen & ", "
    SQLQuery = SQLQuery & 0 & ", "   'lstUnit
    SQLQuery = SQLQuery & lst_rst!lstCifCode & ", "
    SQLQuery = SQLQuery & 0 & ", " 'lstAnfCode
    SQLQuery = SQLQuery & 0 & ", "    'lstEvtIDCefCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstSplitNetwork
    SQLQuery = SQLQuery & 0 & ", "     'lstRafCode
    SQLQuery = SQLQuery & 0 & ", " 'lstFsfCode
    SQLQuery = SQLQuery & 0 & ", " 'lstGsfCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstImportedSpot
    SQLQuery = SQLQuery & 0 & ", "    'lstBkoutLstCode
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "  'lstLnStartTime
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "    'lstLnEndTime
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llLst = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    If llLst > 0 Then
        mAddMGLst = llLst
    Else
        mAddMGLst = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
End Function

Private Function mAddAstMG(llLstCode As Long) As Long
    Dim llAst As Long

    On Error GoTo ErrHand
    SQLQuery = "Insert Into ast ( "
    SQLQuery = SQLQuery & "astCode, "
    SQLQuery = SQLQuery & "astAtfCode, "
    SQLQuery = SQLQuery & "astShfCode, "
    SQLQuery = SQLQuery & "astVefCode, "
    SQLQuery = SQLQuery & "astSdfCode, "
    SQLQuery = SQLQuery & "astLsfCode, "
    SQLQuery = SQLQuery & "astAirDate, "
    SQLQuery = SQLQuery & "astAirTime, "
    SQLQuery = SQLQuery & "astStatus, "
    SQLQuery = SQLQuery & "astCPStatus, "
    SQLQuery = SQLQuery & "astFeedDate, "
    SQLQuery = SQLQuery & "astFeedTime, "
    SQLQuery = SQLQuery & "astAdfCode, "
    SQLQuery = SQLQuery & "astDatCode, "
    SQLQuery = SQLQuery & "astCpfCode, "
    SQLQuery = SQLQuery & "astRsfCode, "
    SQLQuery = SQLQuery & "astStationCompliant, "
    SQLQuery = SQLQuery & "astAgencyCompliant, "
    SQLQuery = SQLQuery & "astAffidavitSource, "
    SQLQuery = SQLQuery & "astCntrNo, "
    SQLQuery = SQLQuery & "astLen, "
    SQLQuery = SQLQuery & "astLkAstCode, "
    SQLQuery = SQLQuery & "astMissedMnfCode, "
    SQLQuery = SQLQuery & "astUstCode "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & ast_rst!astAtfCode & ", "
    SQLQuery = SQLQuery & ast_rst!astShfCode & ", "
    SQLQuery = SQLQuery & ast_rst!astVefCode & ", "
    SQLQuery = SQLQuery & 0 & ", "         'astsdfCode
    SQLQuery = SQLQuery & llLstCode & ", "  'astlsfCode
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", " 'astStatus
    SQLQuery = SQLQuery & ast_rst!astCPStatus & ", "
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirDate, sgSQLDateForm) & "', "  'astFeedDate
    SQLQuery = SQLQuery & "'" & Format$(ast_rst!astAirTime, sgSQLTimeForm) & "', "  'astFeedTime
    SQLQuery = SQLQuery & ast_rst!astAdfCode & ", "
    SQLQuery = SQLQuery & ast_rst!astDatCode & ", "
    SQLQuery = SQLQuery & ast_rst!astCpfCode & ", "
    SQLQuery = SQLQuery & ast_rst!astRsfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'astStationCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "   'astAgencyCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote(gRemoveIllegalChars(ast_rst!astAffidavitSource)) & "', "
    SQLQuery = SQLQuery & ast_rst!astCntrNo & ", "
    SQLQuery = SQLQuery & ast_rst!astLen & ", "
    SQLQuery = SQLQuery & ast_rst!astCode & ", "    'astLkAstCode
    SQLQuery = SQLQuery & 0 & ", "   'astMissedMnfCode
    SQLQuery = SQLQuery & igUstCode 'astUstCode
    SQLQuery = SQLQuery & ") "
    llAst = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
    If llAst > 0 Then
        mAddAstMG = llAst
    Else
        mAddAstMG = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
End Function


Private Sub mChgAstToMissed(llMGAstCode As Long, slPdDate As String, slPdTime As String)
    On Error GoTo ErrHand
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery & "astAirDate = '" & Format$(slPdDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "astAirTime = '" & Format$(slPdTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery + "astLkAstCode = " & llMGAstCode & ", "
    SQLQuery = SQLQuery + "astAgencyCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStationCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStatus = " & 4
    SQLQuery = SQLQuery + " WHERE (astCode = " & ast_rst!astCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffErrorLog.txt", "SetMG-mChgAstToMissed"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
'    Return
End Sub

