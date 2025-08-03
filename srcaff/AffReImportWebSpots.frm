VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmReImportWebSpots 
   Caption         =   "Re-Import Affiliate Spots"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "AffReImportWebSpots.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9900
   Begin VB.Frame frcReImport 
      Caption         =   "SQL Call"
      Height          =   1995
      Index           =   0
      Left            =   1290
      TabIndex        =   6
      Top             =   1935
      Visible         =   0   'False
      Width           =   5925
      Begin VB.TextBox edcSQLCall 
         Height          =   930
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   270
         Width           =   5550
      End
      Begin VB.Label lacWFileInfo 
         Caption         =   "(SQL Call Select must have either astAtfCode and astFeedDate or cpttAtfCode and cpttStartDate or * with From of ast or cptt)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   135
         TabIndex        =   47
         Top             =   1290
         Width           =   5580
      End
   End
   Begin VB.Frame frcReImport 
      Caption         =   "TTP 7906"
      Height          =   2235
      Index           =   4
      Left            =   1020
      TabIndex        =   37
      Top             =   1395
      Visible         =   0   'False
      Width           =   5985
      Begin V81Affiliate.CSI_Calendar edcTTPStartDate 
         Height          =   285
         Left            =   1245
         TabIndex        =   39
         Top             =   345
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "12/23/2022"
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
      Begin V81Affiliate.CSI_Calendar edcTTPEndDate 
         Height          =   285
         Left            =   4200
         TabIndex        =   41
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "12/23/2022"
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
      Begin VB.Label lacTTPEndDate 
         Caption         =   "End Date"
         Height          =   255
         Left            =   3075
         TabIndex        =   40
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label lacTTPStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.OptionButton rbcReImport 
      Caption         =   "TTP 7906"
      Height          =   210
      Index           =   4
      Left            =   3120
      TabIndex        =   5
      Top             =   300
      Width           =   1140
   End
   Begin VB.OptionButton rbcReImport 
      Caption         =   "Station/Vehicles"
      Height          =   210
      Index           =   3
      Left            =   4770
      TabIndex        =   3
      Top             =   75
      Width           =   1695
   End
   Begin VB.Frame frcReImport 
      Caption         =   "Station/Vehicles"
      Height          =   3360
      Index           =   3
      Left            =   660
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   6030
      Begin V81Affiliate.CSI_Calendar edcSVDate 
         Height          =   285
         Left            =   1290
         TabIndex        =   26
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "12/23/2022"
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
         TabIndex        =   29
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Vehicles"
         Top             =   660
         Width           =   3810
      End
      Begin VB.CheckBox chkSVAllStation 
         Caption         =   "All"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ListBox lbcSVStation 
         Height          =   2010
         ItemData        =   "AffReImportWebSpots.frx":08CA
         Left            =   120
         List            =   "AffReImportWebSpots.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   915
         Width           =   1695
      End
      Begin VB.TextBox edcSVWeeks 
         Height          =   285
         Left            =   4740
         TabIndex        =   28
         Text            =   "1"
         Top             =   225
         Width           =   405
      End
      Begin VB.CheckBox chkSVAllVehicles 
         Caption         =   "All"
         Height          =   195
         Left            =   2010
         TabIndex        =   34
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ListBox lbcSVVehicles 
         Height          =   2010
         ItemData        =   "AffReImportWebSpots.frx":08CE
         Left            =   1980
         List            =   "AffReImportWebSpots.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   915
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lacSVStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lacSVWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3225
         TabIndex        =   27
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.Frame frcReImport 
      Caption         =   "Agreement Code"
      Height          =   2520
      Index           =   1
      Left            =   420
      TabIndex        =   8
      Top             =   765
      Visible         =   0   'False
      Width           =   5985
      Begin V81Affiliate.CSI_Calendar edcACDate 
         Height          =   285
         Left            =   1245
         TabIndex        =   12
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "12/23/2022"
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
         TabIndex        =   10
         Top             =   225
         Width           =   990
      End
      Begin VB.TextBox edcACWeeks 
         Height          =   285
         Left            =   4695
         TabIndex        =   14
         Text            =   "1"
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lacAC 
         Caption         =   "Agreement Code"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lacACWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3180
         TabIndex        =   13
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label lacACStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   1215
      End
   End
   Begin VB.OptionButton rbcReImport 
      Caption         =   "Vehicle/Stations"
      Height          =   210
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   75
      Width           =   1695
   End
   Begin VB.OptionButton rbcReImport 
      Caption         =   "Agreement Code"
      Height          =   210
      Index           =   1
      Left            =   1470
      TabIndex        =   1
      Top             =   75
      Width           =   1605
   End
   Begin VB.OptionButton rbcReImport 
      Caption         =   "SQL Call"
      Height          =   210
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   300
      Width           =   1035
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   3840
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2985
      ItemData        =   "AffReImportWebSpots.frx":08D2
      Left            =   6585
      List            =   "AffReImportWebSpots.frx":08D4
      TabIndex        =   45
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
   Begin VB.CommandButton cmcReImport 
      Caption         =   "&Re-Import"
      Height          =   375
      Left            =   5910
      TabIndex        =   42
      Top             =   4290
      Width           =   1665
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   43
      Top             =   4290
      Width           =   1665
   End
   Begin VB.Frame frcReImport 
      Caption         =   "Vehicle/Stations"
      Height          =   3360
      Index           =   2
      Left            =   285
      TabIndex        =   15
      Top             =   540
      Visible         =   0   'False
      Width           =   6030
      Begin V81Affiliate.CSI_Calendar edcVSDate 
         Height          =   285
         Left            =   1290
         TabIndex        =   17
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Text            =   "12/23/2022"
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
         ItemData        =   "AffReImportWebSpots.frx":08D6
         Left            =   120
         List            =   "AffReImportWebSpots.frx":08D8
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   915
         Width           =   3855
      End
      Begin VB.CheckBox chkVSAllVehicles 
         Caption         =   "All"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   3060
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox edcVSWeeks 
         Height          =   285
         Left            =   4740
         TabIndex        =   19
         Text            =   "1"
         Top             =   225
         Width           =   405
      End
      Begin VB.ListBox lbcVSStation 
         Height          =   2010
         ItemData        =   "AffReImportWebSpots.frx":08DA
         Left            =   4230
         List            =   "AffReImportWebSpots.frx":08DC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   915
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkVSAllStation 
         Caption         =   "All"
         Height          =   195
         Left            =   4245
         TabIndex        =   36
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
         TabIndex        =   20
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
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Stations"
         Top             =   675
         Width           =   1635
      End
      Begin VB.Label lacVSWeeks 
         Caption         =   "Number of Weeks"
         Height          =   255
         Left            =   3225
         TabIndex        =   18
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lacVSStartDate 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Label lbcWebType 
      Caption         =   "Production Website"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   50
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lacStartTime 
      Height          =   195
      Left            =   5955
      TabIndex        =   49
      Top             =   4680
      Width           =   3345
   End
   Begin VB.Label lacCounts 
      Height          =   195
      Left            =   5940
      TabIndex        =   48
      Top             =   3975
      Width           =   3345
   End
   Begin VB.Label lacReImport 
      Caption         =   "Re-Import by"
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   135
      TabIndex        =   46
      Top             =   4215
      Width           =   5580
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6915
      TabIndex        =   44
      Top             =   390
      Width           =   1965
   End
End
Attribute VB_Name = "frmReImportWebSpots"
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
Private imReImporting As Integer
Private lmDaysRetainSpots As Long
Private lmCheckHistoryFirstDate As Long
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmFrom As Integer
Private cprst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private AgreementInfo_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private lmTotalProcessedCount As Long
Private lmTotalToProcess As Long
Private smSvLogActivityInto As String


Dim myImport As CMarketron



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
    slToFile = sgMsgDirectory & "ReImportAffiliateSpots.Txt"
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
    Print #hmMsg, "** Re-Import Affiliate Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
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


Private Sub cmcReImport_Click()
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
    
    If imReImporting Then
        Exit Sub
    End If
    imReImporting = True
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
    If rbcReImport(0).Value Then
        If Trim$(edcSQLCall.Text) = "" Then
            gMsgBox "SQL Call must be specified.", vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        ilRet = 0
        On Error GoTo cmcReImportErr:
        slStr = Trim$(edcSQLCall.Text)
        SQLQuery = slStr
        If (InStr(1, slStr, " ast", vbTextCompare) > 0) Or (InStr(1, slStr, "ast,", vbTextCompare) > 0) Then
        ElseIf (InStr(1, slStr, " cptt", vbTextCompare) > 0) Or (InStr(1, slStr, "cptt,", vbTextCompare) > 0) Then
        Else
            gMsgBox "SQL Call missing ast or cptt reference", vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        lacResult.Caption = "Gathering Agreement Information"
        DoEvents
        Set cprst = gSQLSelectCall(SQLQuery)
        If iRet <> 0 Then
            lacResult.Caption = ""
            gMsgBox "SQL Call Structure in Error: " & SQLQuery, vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If cprst.EOF Then
            lacResult.Caption = ""
            gMsgBox "No Records return by the SQL Call", vbOKOnly
            imReImporting = False
            Exit Sub
        End If
        On Error GoTo ErrHand
        If (InStr(1, slStr, " ast", vbTextCompare) > 0) Or (InStr(1, slStr, "ast,", vbTextCompare) > 0) Then
            Screen.MousePointer = vbHourglass
            On Error GoTo cmcReImportErr:
            Do While Not cprst.EOF
                iRet = 0
                llAttCode = cprst!astAtfCode
                If iRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    gMsgBox "SQL Call Structure in  missing astAtfCode reference: " & SQLQuery, vbOKOnly
                    edcACDate.SetFocus
                    imReImporting = False
                    Exit Sub
                End If
                llMoDate = gDateValue(gObtainPrevMonday(Format(cprst!astFeedDate, "m/d/yy")))
                If iRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    gMsgBox "SQL Call Structure  missing astFeedDate reference: " & SQLQuery, vbOKOnly
                    edcACDate.SetFocus
                    imReImporting = False
                    Exit Sub
                End If
                If (llPrevAtt <> llAttCode) Or (llPrevDate <> llMoDate) Then
                    llPrevAtt = llAttCode
                    llPrevDate = llMoDate
                    mAddAgreementInfo llAttCode, llMoDate
                End If
                cprst.MoveNext
            Loop
            On Error GoTo ErrHand
        ElseIf (InStr(1, slStr, " cptt", vbTextCompare) > 0) Or (InStr(1, slStr, "cptt,", vbTextCompare) > 0) Then
            Screen.MousePointer = vbHourglass
            On Error GoTo cmcReImportErr:
            Do While Not cprst.EOF
                iRet = 0
                llAttCode = cprst!cpttatfCode
                If iRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    gMsgBox "SQL Call Structure missing cpttAtfCode reference: " & SQLQuery, vbOKOnly
                    edcACDate.SetFocus
                    imReImporting = False
                    Exit Sub
                End If
                llMoDate = gDateValue(gObtainPrevMonday(Format(cprst!CpttStartDate, "m/d/yy")))
                If iRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    gMsgBox "SQL Call Structure missing cpttStartDate reference: " & SQLQuery, vbOKOnly
                    edcACDate.SetFocus
                    imReImporting = False
                    Exit Sub
                End If
                If (llPrevAtt <> llAttCode) Or (llPrevDate <> llMoDate) Then
                    llPrevAtt = llAttCode
                    llPrevDate = llMoDate
                    mAddAgreementInfo llAttCode, llMoDate
                End If
                cprst.MoveNext
            Loop
            On Error GoTo ErrHand
        End If
        lacResult.Caption = ""
    ElseIf rbcReImport(1).Value Then
        If edcAC.Text = "" Then
            gMsgBox "Agreement Code must be specified.", vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If edcACDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If gIsDate(edcACDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        Else
            smDate = Format(edcACDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcACWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcACWeeks.SetFocus
            imReImporting = False
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
    ElseIf rbcReImport(2).Value Then
        If lbcVSVehicles.ListIndex < 0 Then
            imReImporting = False
            Exit Sub
        End If
        If edcVSDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcVSDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If gIsDate(edcVSDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcVSDate.SetFocus
            imReImporting = False
            Exit Sub
        Else
            smDate = Format(edcVSDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcVSDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcVSWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcVSWeeks.SetFocus
            imReImporting = False
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
    ElseIf rbcReImport(3).Value Then
        If lbcSVStation.ListIndex < 0 Then
            imReImporting = False
            Exit Sub
        End If
        If edcSVDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcSVDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If gIsDate(edcSVDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcSVDate.SetFocus
            imReImporting = False
            Exit Sub
        Else
            smDate = Format(edcSVDate.Text, sgShowDateForm)
        End If
        sNowDate = Format$(gNow(), "m/d/yy")
        If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
            Beep
            gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
            edcSVDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        slMoDate = gObtainPrevMonday(smDate)
        llSDate = DateValue(gAdjYear(slMoDate))
        imWeeks = Val(edcSVWeeks.Text)
        If imWeeks <= 0 Then
            gMsgBox "Number of Weeks must be specified.", vbOKOnly
            edcSVWeeks.SetFocus
            imReImporting = False
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
    ElseIf rbcReImport(4).Value Then
        If edcTTPStartDate.Text <> "" Then
            If gIsDate(edcTTPStartDate.Text) = False Then
                Beep
                gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
                edcTTPStartDate.SetFocus
                imReImporting = False
                Exit Sub
            Else
                smDate = gObtainPrevMonday(Format(edcTTPStartDate.Text, sgShowDateForm))
            End If
        Else
            smDate = ""
        End If
        If edcTTPEndDate.Text <> "" Then
            If gIsDate(edcTTPEndDate.Text) = False Then
                Beep
                gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
                edcTTPEndDate.SetFocus
                imReImporting = False
                Exit Sub
            Else
                smEndDate = gObtainNextSunday(Format(edcTTPEndDate.Text, sgShowDateForm))
            End If
        Else
            smEndDate = ""
        End If
        ilRet = 0
        On Error GoTo cmcReImportErr:
        If (smDate = "") And (smEndDate = "") Then
            'SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astCPStatus = 1 and astStatus <> 8 and astStatus <> 4 and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
            SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astCPStatus = 1 and Mod(astStatus, 100) Not In(4, 8, 14) and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
        ElseIf (smDate <> "") And (smEndDate = "") Then
            'SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astCPStatus = 1 and astStatus <> 8 and astStatus <> 4 and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
            SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) Not In(4, 8, 14)  and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
        Else
            'SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astFeedDate <= '" & Format(smEndDate, sgSQLDateForm) & "' and astCPStatus = 1 and astStatus <> 8 and astStatus <> 4 and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
            SQLQuery = "Select Distinct astAtfCode, astFeedDate from ast left outer join dat on astDatCode = datcode where astFeedDate >= '" & Format(smDate, sgSQLDateForm) & "' and astFeedDate <= '" & Format(smEndDate, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) Not In(4, 8, 14) and datEmbeddedOrROS = 'R' and astAirTime = datPdStTime order by astatfcode, astfeeddate"
        End If
        lacResult.Caption = "Gathering Agreement Information"
        DoEvents
        Set cprst = gSQLSelectCall(SQLQuery)
        If iRet <> 0 Then
            lacResult.Caption = ""
            gMsgBox "SQL Call Structure in Error: " & SQLQuery, vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        If cprst.EOF Then
            lacResult.Caption = ""
            gMsgBox "No Records return by the SQL Call", vbOKOnly
            imReImporting = False
            Exit Sub
        End If
        On Error GoTo ErrHand
        Screen.MousePointer = vbHourglass
        On Error GoTo cmcReImportErr:
        Do While Not cprst.EOF
            llAttCode = cprst!astAtfCode
            llMoDate = gDateValue(gObtainPrevMonday(Format(cprst!astFeedDate, "m/d/yy")))
            If (llPrevAtt <> llAttCode) Or (llPrevDate <> llMoDate) Then
                llPrevAtt = llAttCode
                llPrevDate = llMoDate
                'slSQLQuery = "Select Count(*) as Total from ast left outer join dat on astDatCode = datcode where astAtfCode= " & llAttCode & " and astFeedDate >= '" & Format(llMoDate, sgSQLDateForm) & "' And astFeedDate <= '" & Format(llMoDate + 6, sgSQLDateForm) & "' and astCPStatus = 1 and astStatus <> 8 and astStatus <> 4 and datEmbeddedOrROS = 'R' and astAirTime <> datPdStTime"
                slSQLQuery = "Select Count(*) as Total from ast left outer join dat on astDatCode = datcode where astAtfCode= " & llAttCode & " and astFeedDate >= '" & Format(llMoDate, sgSQLDateForm) & "' And astFeedDate <= '" & Format(llMoDate + 6, sgSQLDateForm) & "' and astCPStatus = 1 and Mod(astStatus, 100) Not IN (4, 8, 14) and datEmbeddedOrROS = 'R' and astAirTime <> datPdStTime"
                Set ast_rst = gSQLSelectCall(slSQLQuery)
                If ast_rst!Total = 0 Then
                    mAddAgreementInfo llAttCode, llMoDate
                End If
            End If
            cprst.MoveNext
        Loop
        On Error GoTo ErrHand
        lacResult.Caption = ""
    Else
        Beep
        gMsgBox "'Re-Import by' must be specified", vbCritical
        imReImporting = False
        Exit Sub
    End If
    lacStartTime.Caption = Now
    AgreementInfo_rst.Filter = adFilterNone
    lmTotalToProcess = AgreementInfo_rst.RecordCount
    If lmTotalToProcess <= 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "No Agreements found to be processed.", vbOKOnly
        imReImporting = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lmCheckHistoryFirstDate = DateValue(gObtainPrevSunday(Format$(DateAdd("d", -(lmDaysRetainSpots - 14), Format(Now, "mm/dd/yy")), "mm/dd/yy")))
    
    If Not mOpenMsgFile(sMsgFileName) Then
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        imReImporting = False
        Exit Sub
    End If
    
    ilRet = gPopAttInfo()
    ilRet = gPopAll

    If rbcReImport(1).Value Or rbcReImport(2).Value Or rbcReImport(3).Value Then
        llSDate = DateValue(gAdjYear(slMoDate))
        llEDate = DateValue(gAdjYear(Format$(DateAdd("d", 7 * (imWeeks - 1) + 6, slMoDate), "mm/dd/yy")))
    End If
    If rbcReImport(0).Value Then
        Print #hmMsg, "Re-Import by SQL Call"
        Print #hmMsg, "  SQL Call: " & Trim$(edcSQLCall.Text)
    ElseIf rbcReImport(1).Value Then
        Print #hmMsg, "Re-Import by Agreement Code"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Agreement Code: " & Trim$(edcAC.Text)
    ElseIf rbcReImport(2).Value Then
        Print #hmMsg, "Re-Import by Vehicle/Station"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Vehicle: " & slVehicles
        Print #hmMsg, "  Stations: " & slstations
    ElseIf rbcReImport(3).Value Then
        Print #hmMsg, "Re-Import by Station/Vehicle"
        Print #hmMsg, "  Date Range: " & Format(llSDate, "m/d/yy") & "-" & Format(llEDate, "m/d/yy")
        Print #hmMsg, "  Station: " & slstations
        Print #hmMsg, "  Vehicles: " & slVehicles
    ElseIf rbcReImport(4).Value Then
        Print #hmMsg, "Re-Import by TTP 7906"
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
    If (rbcReImport(4).Value) And (smEndDate = "") Then
        iRet = 0
        On Error GoTo cmcReImportErr:
        lacResult.Caption = "Clearing file abf_Ast_Build_Queue"
        SQLQuery = "Delete From abf_Ast_Build_Queue"
        Set cprst = gSQLSelectCall(SQLQuery)
        If iRet <> 0 Then
            lacResult.Caption = ""
            gMsgBox "SQL Call Structure in Error: " & SQLQuery, vbOKOnly
            edcACDate.SetFocus
            imReImporting = False
            Exit Sub
        End If
        On Error GoTo ErrHand
        lacResult.Caption = ""
    End If
    iRet = mReImportSpots()
    If (iRet = False) Then
        'Stop the Pervasive API engine
        Print #hmMsg, "** Terminated **"
        Close #hmMsg
        imReImporting = False
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        'Stop the Pervasive API engine
        Print #hmMsg, "** User Terminated **"
        Close #hmMsg
        imReImporting = False
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    imReImporting = False
    Print #hmMsg, "** Completed Re-Import Affiliate Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacStartTime.Caption = lacStartTime.Caption & " to " & Now
    lacResult.Caption = "Results: " & sMsgFileName
    cmcReImport.Enabled = False
    cmcCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmcReImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    imReImporting = False
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-cmcReImport"
    Exit Sub
End Sub

Private Sub cmcCancel_Click()
    If imReImporting Then
        imTerminate = True
        Exit Sub
    End If
    edcVSDate.Text = ""
    Unload frmReImportWebSpots
End Sub


Private Sub edcAC_Change()
    If cmcReImport.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcACDate_CalendarChanged()
    If cmcReImport.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcACWeeks_Change()
    If cmcReImport.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcSQLCall_Change()
    If cmcReImport.Enabled = False Then
        mClearControls
    End If
End Sub

Private Sub edcVSDate_Change()
    lbcMsg.Clear
    If cmcReImport.Enabled = False Then
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
    imReImporting = False
    imFirstTime = True
        '10000
    lbcWebType.FontSize = 6
    If igDemoMode Then
        lbcWebType.Caption = "Demo Mode"
    ElseIf gIsTestWebServer() Then
        lbcWebType.Caption = "Test Website"
    End If
    frcReImport(0).Move lacReImport.Left - 15, rbcReImport(0).Top + rbcReImport(0).Height + 90
    frcReImport(1).Move frcReImport(0).Left, frcReImport(0).Top
    frcReImport(2).Move frcReImport(0).Left, frcReImport(0).Top
    frcReImport(3).Move frcReImport(0).Left, frcReImport(0).Top
    frcReImport(4).Move frcReImport(0).Left, frcReImport(0).Top
    lbcVSStation.Clear
    mVSFillVehicle
    lbcSVVehicles.Clear
    mSVFillStation
    lmDaysRetainSpots = 180
    SQLQuery = "Select siteDaysRetainSpots From site where sitecode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        lmDaysRetainSpots = Val(rst!siteDaysRetainSpots)
    End If

    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If imReImporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    sgLogActivityInto = smSvLogActivityInto
    mCloseAgreementInfo
    cprst.Close
    att_rst.Close
    ast_rst.Close
    Set frmReImportWebSpots = Nothing
End Sub


Private Sub lbcSVStation_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcSVVehicles.Clear
    If cmcReImport.Enabled = False Then
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
    If cmcReImport.Enabled = False Then
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
    If cmcReImport.Enabled = False Then
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
    If cmcReImport.Enabled = False Then
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

Private Function mReImportSpots() As Integer
    Dim ilRet As Integer
    Dim llAttCode As Long
    Dim llMoDate As Long
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slFileName As String
    Dim slIniValue As String
    Dim slLocation As String
    Dim llLen As Long
    Dim fs As New FileSystemObject
    Dim slSvCommand As String
    Dim slSelect As String
    Dim slWhere As String
    Dim llUnpostedSpotCount As Long
    Dim llPostedSpotCount As Long
    Dim llExportCount As Long
    Dim blAttFound As Boolean
    Dim blContinue As Boolean
    Dim ilVefCode As Integer
    Dim blGetFromWeb As Boolean
    Dim slSQLQuery As String
        
    On Error GoTo ErrHand
    lmTotalProcessedCount = 0
    slSvCommand = sgCommand
    slIniValue = "WebImports"
    Call gLoadOption(sgWebServerSection, slIniValue, slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    
    slSelect = "Select attCode, Advt, Prod, PledgeStartDate, PledgeEndDate, PledgeStartTime, PledgeEndTime, SpotLen, Cart, ISCI, CreativeTitle, astCode, ActualDateTime, ActualDateTime, statusCode, FeedDate, FeedTime"
    'slSelect = "Select attCode,Advt,Prod,convert(char(10), PledgeStartDate, 20) as PledgeStartDate1,convert(char(10), PledgeEndDate, 20) as PledgeEndDate,FORMAT(PledgeStartTime,'hh:mm:ss tt') as PledgeStartTime,FORMAT(PledgeEndTime,'hh:mm:ss tt') as PledgeEndTime,SpotLen,Cart,ISCI,CreativeTitle,astCode,convert(char(10), ActualDateTime, 20) as ActualAirDate1,FORMAT(ActualDateTime,'hh:mm:ss tt') as ActualAirTime1,ISNULL(statusCode, -1) as statusCode,convert(char(10), FeedDate, 105) as FeedDate,FORMAT(FeedTime,'hh:mm:ss tt') as FeedTime,ISNULL(RecType, 0) as RecType,ISNULL(MRReason, 0) as MRReason,ISNULL(OrgAstCode, 0 ) as OrgAstCode,ISNULL(NewAstCode, 0) as NewAstCode,ISNULL(srcAttCode, 0) as srcAttCode,ISNULL(gsfCode,0) as gsfCode,Source"
    
    On Error Resume Next
    AgreementInfo_rst.Filter = adFilterNone
    If Not (AgreementInfo_rst.EOF And AgreementInfo_rst.BOF) Then
        AgreementInfo_rst.MoveFirst
    End If
    'one record for each agreement/week
    Do While Not AgreementInfo_rst.EOF
        If imTerminate Then
            mReImportSpots = True
            Exit Function
        End If
        llAttCode = AgreementInfo_rst!attCode
        blAttFound = True
        SQLQuery = "Select vefName, vefCode, shttCallLetters, vatWvtVendorId, wvtImportMethod from att left outer join vef_vehicles on attVefCode = vefCode Left Outer Join shtt On attShfCode = shttCode Left outer join VAT_Vendor_Agreement on attcode = vatAttCode Left Outer Join WVT_Vendor_Table On vatWvtVendorId = wvtVendorID Where attCode = " & llAttCode

        Set att_rst = gSQLSelectCall(SQLQuery)
        If Not att_rst.EOF Then
            lacResult.Caption = "Processing: " & Trim$(att_rst!vefName) & " " & Trim$(att_rst!shttCallLetters)
            Print #hmMsg, "    Processing: Vehicle- " & Trim$(att_rst!vefName) & " Station- " & Trim$(att_rst!shttCallLetters)
        Else
            blAttFound = False
            lbcMsg.AddItem "Agreement Not Found: " & llAttCode
            Print #hmMsg, "    Agreement Not Found: " & llAttCode
        End If
        If blAttFound Then
            'Obtain file from Web or Marketron?
            blGetFromWeb = True
            If (att_rst!vatwvtvendorid = Vendors.NetworkConnect) And (att_rst!wvtImportMethod = 0) Then
                blGetFromWeb = False
                Print #hmMsg, "    Re-Import via: Network Connect"
            Else
                Print #hmMsg, "    Re-Import via: Counterpoint Affidavit System"
            End If
            If blGetFromWeb Then
                'Web
                blContinue = True
                llMoDate = AgreementInfo_rst!MoDate
                slMoDate = Format(llMoDate, "yyyy-mm-dd")
                slSuDate = Format(llMoDate + 6, "yyyy-mm-dd")
                slWhere = " Where attCode = " & llAttCode & " And RecType <> 'D'" & " And FeedDate >= '" & slMoDate & "' And " & " FeedDate <= '" & slSuDate & "'"
                'slWhere = " Where attCode = " & llAttCode & "  And PledgeStartDate >= '" & slMoDate & "' And " & " PledgeStartDate <= '" & slSuDate & "'"
                
                If (llMoDate > lmCheckHistoryFirstDate) Then
                    'Obtain Web spots from Spots
                    SQLQuery = "Select Count(*) from Spots " & slWhere & " And postedFlag = 0"
                    llUnpostedSpotCount = gExecWebSQLWithRowsEffected(SQLQuery)
                    If llUnpostedSpotCount = 0 Then
                        SQLQuery = slSelect & " from Spots"
                    Else
                        'Add message: Spots not posted
                        Print #hmMsg, "      Spots Not Completely Posted " & Format(slMoDate, "m/d/yy") & "-" & Format(slSuDate, "m/d/yy")
                        blContinue = False
                    End If
                Else
                    'Check if Spots in Spot_History, if not then check in Spots
                    SQLQuery = "Select Count(*) from Spot_History " & slWhere & " And postedFlag = 1"
                    llPostedSpotCount = gExecWebSQLWithRowsEffected(SQLQuery)
                    If llPostedSpotCount = 0 Then
                        SQLQuery = slSelect & " from Spots"
                    Else
                        SQLQuery = slSelect & " from Spot_History"
                    End If
                End If
                If (blContinue) And InStr(1, SQLQuery, "Spots", vbTextCompare) > 0 Then
                    slSQLQuery = "Select Count(*) From Spots " & slWhere & " And ExportedFlag = 1"
                    llExportCount = gExecWebSQLWithRowsEffected(slSQLQuery)
                    If llExportCount <= 0 Then
                        Print #hmMsg, "      Spots Not Previously Exported to Affiliate " & Format(slMoDate, "m/d/yy") & "-" & Format(slSuDate, "m/d/yy")
                        blContinue = False
                    End If
                End If
                If blContinue Then
                    SQLQuery = SQLQuery & slWhere & " And postedFlag = 1"
                    slFileName = "WebSpots_" & "ReImport.txt"
                    ilRet = gRemoteExecSql(SQLQuery, slFileName, slIniValue, True, True, 30)
                    If fs.FILEEXISTS(slLocation & slFileName) Then
                        llLen = FileLen(slLocation & slFileName)
                        If llLen > 250 Then 'Head is 230 characters
                            mClearMGs llAttCode, llMoDate
                            mUpdateCpttAsNotPosted llAttCode, slMoDate
                            mUpdateAstAsNotPosted llAttCode, slMoDate
                            sgCommand = "/ReImport"
                            frmWebImportAiredSpot.Show vbModal
                            Screen.MousePointer = vbHourglass
                            If InStr(1, sgReImportStatus, "Successful", vbTextCompare) > 0 Then
                                mUpdateCpttToNotCreateAsts llAttCode, slMoDate
                            End If
                            Print #hmMsg, "       " & sgReImportStatus & " " & Format(slMoDate, "m/d/yy") & "-" & Format(slSuDate, "m/d/yy")
                        Else
                            'No spots returned message
                            Print #hmMsg, "       No Spots Found or Not Posted " & Format(slMoDate, "m/d/yy") & "-" & Format(slSuDate, "m/d/yy")
                        End If
                    End If
                End If
            Else
                'Marketron
                llMoDate = AgreementInfo_rst!MoDate
                slMoDate = Format(llMoDate, "m/d/yy")
                slSuDate = Format(llMoDate + 6, "m/d/yy")
                ilRet = mCreateMarketronImportFile(att_rst!vefCode, Trim$(att_rst!vefName), Trim$(att_rst!shttCallLetters), llMoDate)
                If ilRet Then
                    mClearMGs llAttCode, llMoDate
                    mUpdateCpttAsNotPosted llAttCode, slMoDate
                    mUpdateAstAsNotPosted llAttCode, slMoDate
                    sgCommand = "/ReImport"
                    frmImportMarketron.Show vbModal
                    Screen.MousePointer = vbHourglass
                    If InStr(1, sgReImportStatus, "Successful", vbTextCompare) > 0 Then
                        mUpdateCpttToNotCreateAsts llAttCode, slMoDate
                    End If
                    Print #hmMsg, "       " & sgReImportStatus & " " & slMoDate & "-" & slSuDate
                Else
                    Print #hmMsg, "       No Spots Found or Not Posted " & slMoDate & "-" & slSuDate
                End If
            End If
        End If
        lmTotalProcessedCount = lmTotalProcessedCount + 1
        lacCounts.Caption = "Processed: " & lmTotalProcessedCount & " of " & lmTotalToProcess
        AgreementInfo_rst.MoveNext
    Loop
    sgCommand = slSvCommand
    mReImportSpots = True
    Exit Function
mReImportSpotsErr:
    ilRet = Err
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mReImportSpots"
    mReImportSpots = False
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
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mFileStations"

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
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mSVFillVehicle"

End Sub
Private Sub rbcReImport_Click(Index As Integer)
    If cmcReImport.Enabled = False Then
        mClearControls
    End If
    frcReImport(0).Visible = False
    frcReImport(1).Visible = False
    frcReImport(2).Visible = False
    frcReImport(3).Visible = False
    frcReImport(4).Visible = False
    frcReImport(Index).Visible = True
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmReImportWebSpots
End Sub

Private Sub edcVSWeeks_Change()
    If cmcReImport.Enabled = False Then
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

Private Function mCreateMarketronImportFile(ilVefCode As Integer, slInVehicleName As String, slInCallLetters As String, llDate As Long) As Integer
    'myImport must be set at module level
    Dim slIniPath As String
    Dim slImportPath As String
    Dim ilOnMarketron As Integer
    Dim slOrderId As String
    Dim ilPass As Integer
    Dim slCallLetters As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim slVehicleName As String
    
    mCreateMarketronImportFile = True
    slDate = Format(llDate, "yyyymmdd")
    slVehicleName = gXMLNameFilter(slInVehicleName)
    slVehicleName = mSafeFileName(slVehicleName)
    slCallLetters = slInCallLetters
    ilPos = InStr(1, slCallLetters, "-", vbTextCompare)
    If ilPos > 0 Then
        slCallLetters = Left$(slInCallLetters, ilPos - 1) & Mid$(slInCallLetters, ilPos + 1)
    End If
    For ilPass = 0 To 1 Step 1
        'slOrderID = gXMLNamefilter(Vehicle Name-Vefcode\CallLettersBand\yyyymmdd)
        'For Pass zero, use the dash followed by the vefCode
        'For pass one, leave off the dash and the vefCode
        If ilPass = 0 Then
            slOrderId = gXMLNameFilter(slVehicleName & "-" & ilVefCode & "\" & slCallLetters & "\" & slDate)
        Else
            slOrderId = gXMLNameFilter(slVehicleName & "\" & slCallLetters & "\" & slDate)
        End If
        'slOrderId = "Townsquare Media Network\KSIIFM\20130610"
        'this name is funky, but that's what it is on the test server!
        'slOrderId = "UFUO-8186"
        slImportPath = sgImportDirectory
        slIniPath = gXmlIniPath(True)
        If LenB(slIniPath) = 0 Then
            'mSetResults "Xml.ini doesn't exist.  This form cannot be activated.", MESSAGERED
            Exit Function
        End If
        If Not mSetImportClass(slIniPath) Then
            'problem!
            Exit Function
        End If
        myImport.ImportPath = slImportPath
        myImport.fileName = "ReImportMarketron.Txt"
        'see orderid above  receivedOnly = true will get previously sent only  false will get previously unsent and previously sent.
        ilOnMarketron = myImport.GetOrdersByOrderId(slOrderId, True)
        If ilOnMarketron > 0 Then
            Exit For
        ElseIf Len(myImport.ErrorMessage) > 0 Then
            mCreateMarketronImportFile = False
            Exit For
        Else
            If ilPass = 1 Then
                mCreateMarketronImportFile = False
            End If
        End If
    Next ilPass
    
End Function
Private Function mSetImportClass(slIniPath As String) As Boolean
    'return true if values exist in ini file, not if created myExport
    Dim slRet As String
    Dim blRet As Boolean
    Dim slServicePage As String
    Dim slHost As String
    Dim slPassword As String
    Dim slUserName As String
'    Dim myXml As MSXML2.DOMDocument
'    Dim myElem As MSXML2.IXMLDOMElement
    Dim blReturnAll As Boolean
    '7539
    Dim slProxyUrl As String
    Dim slProxyPort As String
    Dim slProxyTestUrl As String
    Dim blUseSecure As Boolean
    Dim blUseProxySecure As Boolean
    
    blRet = False
    slUserName = ""
    slPassword = ""
    slProxyUrl = ""
    slProxyPort = ""
    slProxyTestUrl = ""
    blUseSecure = False
    blUseProxySecure = False
On Error GoTo ERRORBOX
    'treat as if not needed
    gLoadFromIni "MARKETRON", "ProxyServer", slIniPath, slRet
    If slRet <> "Not Found" Then
        slProxyUrl = slRet
        'must have port defined also
        gLoadFromIni "MARKETRON", "ProxyPort", slIniPath, slRet
        If slRet <> "Not Found" Then
            slProxyPort = slRet
            gLoadFromIni "MARKETRON", "ProxyTestURL", slIniPath, slRet
            If slRet <> "Not Found" Then
                slProxyTestUrl = slRet
            End If
            gLoadFromIni "MARKETRON", "UseSecureProxy", slIniPath, slRet
            If UCase(slRet) = "TRUE" Then
                blUseProxySecure = True
            End If
        Else
            slProxyUrl = ""
        End If
    End If
    gLoadFromIni "MARKETRON", "UseSecure", slIniPath, slRet
    If UCase(slRet) = "TRUE" Then
        blUseSecure = True
    End If
    'will assume 'available', which is the normal download.
    gLoadFromIni "MARKETRON", "OrderStatus", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = "Available"
    End If
    If slRet = "Received" Then
        blReturnAll = True
    Else
        blReturnAll = False
    End If
    'here on out is needed to continue
    gLoadFromIni "MARKETRON", "Host", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    slHost = slRet
    gLoadFromIni "MARKETRON", "WebServiceRcvURL", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    slServicePage = slRet
    gLoadFromIni "MARKETRON", "Authentication", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    blRet = True
    '7878
'    Set myXml = New MSXML2.DOMDocument
'    If Not myXml.loadXML(slRet) Then
'        mSetImportClass = blRet
'        Exit Function
'    End If
'    Set myElem = myXml.selectSingleNode("//Username")
'    If Not myElem Is Nothing Then
'        slUserName = myElem.Text
'    End If
'    Set myElem = myXml.selectSingleNode("//Password")
'    If Not myElem Is Nothing Then
'        slPassword = myElem.Text
'    End If
    slUserName = gParseXml(slRet, "Username", 0)
    slPassword = gParseXml(slRet, "Password", 0)
    If Len(slPassword) > 0 And Len(slUserName) > 0 Then
        Set myImport = New CMarketron
        With myImport
            If StrComp(slHost, "Test", vbTextCompare) = 0 Then
                .isTest = True
            End If
            .SoapUrl = slHost
            .ImportPage = slServicePage
            'couldn't set address
            If Len(.ErrorMessage) > 0 Then
                Set myImport = Nothing
                blRet = False
                GoTo Cleanup
            End If
            .Password = slPassword
            .UserName = slUserName
            .ReturnAll = blReturnAll
            .UseSecure = blUseSecure
            If Len(slProxyUrl) > 0 Then
                If Not .Proxy(slProxyUrl, slProxyPort, blUseProxySecure, slProxyTestUrl) Then
                    blRet = False
                    Set myImport = Nothing
                    GoTo Cleanup
                End If
            End If
         '   .LogPath = .CreateLogName(sgMsgDirectory & FILEDEBUG)
        End With
    End If
Cleanup:
'    Set myElem = Nothing
'    Set myXml = Nothing
    mSetImportClass = blRet
    Exit Function
ERRORBOX:
    blRet = False
'    myErrors.WriteError "mSetImportClass-" & Err.Description
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mSetImportClass"
    GoTo Cleanup
End Function
Private Function mSafeFileName(slOldName As String) As String
    Dim slTempName As String
    If igExportSource = 2 Then DoEvents
    slTempName = Replace(slOldName, "?", "-")
    slTempName = Replace(slTempName, "/", "-")
    slTempName = Replace(slTempName, "\", "-")
    slTempName = Replace(slTempName, "%", "-")
    slTempName = Replace(slTempName, "*", "-")
    slTempName = Replace(slTempName, ":", "-")
    slTempName = Replace(slTempName, "|", "-")
    slTempName = Replace(slTempName, """", "-")
    slTempName = Replace(slTempName, ".", "-")
    slTempName = Replace(slTempName, "<", "-")
    slTempName = Replace(slTempName, ">", "-")
    If igExportSource = 2 Then DoEvents
    mSafeFileName = slTempName
End Function

Private Sub mUpdateCpttToNotCreateAsts(llAttCode As Long, slStartDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE cptt set cpttASTStatus = 'C' "
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slStartDate, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mUpdateCpttToNotCreateAsts"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mUpdateCpttToNotCreateAsts"
End Sub
Private Sub mUpdateCpttAsNotPosted(llAttCode As Long, slStartDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE cptt set "
    SQLQuery = SQLQuery & " cpttStatus = 0,"
    SQLQuery = SQLQuery & " cpttPostingStatus = 0,"
    SQLQuery = SQLQuery & " cpttNoSpotsGen = 0,"
    SQLQuery = SQLQuery & " cpttNoSpotsAired = 0,"
    SQLQuery = SQLQuery & " cpttUsfCode = " & igUstCode
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slStartDate, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mUpdateCpttAsNotPosted"
        Exit Sub
    End If
    gFileChgdUpdate "cptt.mkd", True
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mUpdateCpttAsNotPosted"
End Sub
Private Sub mUpdateAstAsNotPosted(llAttCode As Long, slStartDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE ast set astCPStatus = 0,"
    SQLQuery = SQLQuery & " astStationCompliant = '" & "" & "',"
    SQLQuery = SQLQuery & " astAgencyCompliant = '" & "" & "',"
    SQLQuery = SQLQuery & " astAffidavitSource = '" & "" & "',"
    SQLQuery = SQLQuery & " astUstCode = " & igUstCode & ","
    SQLQuery = SQLQuery & " astStatus = " & "Case When astDatCode <= 0 Then 0 Else (Select datFdStatus From dat Where datCode = astDatCode) End"
    SQLQuery = SQLQuery & " WHERE astAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(DateAdd("d", 6, slStartDate), sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-tmUpdateAstAsNotPosted"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mUpdateAstAsNotPosted"
End Sub
Private Sub mClearControls()
    lacStartTime.Caption = ""
    lacResult.Caption = ""
    lacCounts.Caption = ""
    cmcReImport.Enabled = True
    cmcCancel.Caption = "&Cancel"
    lbcMsg.Clear
End Sub

Private Sub mClearMGs(llAttCode As Long, llMoDate As Long)
    Dim slSQLQuery As String
    slSQLQuery = "Select astCode, astLkAstCode from ast where astAtfCode= " & llAttCode & " and astFeedDate >= '" & Format(llMoDate, sgSQLDateForm) & "' And astFeedDate <= '" & Format(llMoDate + 6, sgSQLDateForm) & "' and astCPStatus = 1 and astLkAstCode > 0 and Mod(astStatus, 100) <= 10"
    Set ast_rst = gSQLSelectCall(slSQLQuery)
    Do While Not ast_rst.EOF
        slSQLQuery = "DELETE FROM Ast WHERE (astCode = " & ast_rst!astLkAstCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mClearMGs"
            Exit Sub
        End If
        slSQLQuery = "UPDATE ast SET "
        slSQLQuery = slSQLQuery & "astLkAstCode = 0"
        slSQLQuery = slSQLQuery + " WHERE (astCode = " & ast_rst!astCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mClearMGs"
            Exit Sub
        End If
        ast_rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mClearMGs"
End Sub
