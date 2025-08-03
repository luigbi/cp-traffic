VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDelqRpt 
   Caption         =   "Weeks  Report"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   Icon            =   "AffDelqRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   8025
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6855
      FormDesignWidth =   8025
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   7545
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6960
         Picture         =   "AffDelqRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Select Stations from File.."
         Top             =   420
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ComboBox cbcSort 
         Height          =   315
         ItemData        =   "AffDelqRpt.frx":0E34
         Left            =   840
         List            =   "AffDelqRpt.frx":0E36
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   2340
      End
      Begin VB.Frame frcPostType 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         TabIndex        =   41
         Top             =   3000
         Width           =   2985
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Both"
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   44
            Top             =   0
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Unposted"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   43
            Top             =   0
            Width           =   1035
         End
         Begin VB.OptionButton rbcPostType 
            Caption         =   "Partial"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.CheckBox ckcAllDelivery 
         Caption         =   "All Delivery Methods"
         Height          =   255
         Left            =   3870
         TabIndex        =   53
         Top             =   2640
         Width           =   2025
      End
      Begin VB.ListBox lbcDelivery 
         Height          =   2010
         ItemData        =   "AffDelqRpt.frx":0E38
         Left            =   3840
         List            =   "AffDelqRpt.frx":0E3F
         MultiSelect     =   2  'Extended
         TabIndex        =   54
         Top             =   2880
         Width           =   3615
      End
      Begin V81Affiliate.CSI_Calendar_UP CalOldestNCRDate 
         Height          =   2025
         Left            =   2160
         TabIndex        =   46
         Top             =   1800
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   3572
         Text            =   "12/22/2022"
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   -1  'True
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
         CSI_CurDayForeColor=   0
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalFromDate 
         Height          =   285
         Left            =   1140
         TabIndex        =   22
         Top             =   735
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Text            =   "12/22/2022"
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
         CSI_CurDayForeColor=   0
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   0
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         TabIndex        =   26
         Top             =   1800
         Width           =   1545
         Begin VB.OptionButton optPage 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton optPage 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   810
            TabIndex        =   28
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.CheckBox ckcUpdateNCR 
         Caption         =   "Update Critically Overdue Flag"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2940
      End
      Begin VB.Frame frcPassword 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   600
         TabIndex        =   38
         Top             =   2760
         Width           =   2145
         Begin VB.OptionButton rbcPassword 
            Caption         =   "Fax #"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton rbcPassword 
            Caption         =   "Password"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   40
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.CheckBox ckcInclComments 
         Caption         =   "Comments"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1815
         ItemData        =   "AffDelqRpt.frx":0E46
         Left            =   3840
         List            =   "AffDelqRpt.frx":0E48
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   810
         Width           =   3495
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select by"
         Height          =   600
         Left            =   3840
         TabIndex        =   48
         Top             =   180
         Width           =   2115
         Begin VB.OptionButton optVehAff 
            Caption         =   "Stations"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   50
            Top             =   240
            Width           =   870
         End
         Begin VB.OptionButton optVehAff 
            Caption         =   "Vehicles"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   6120
         TabIndex        =   51
         Top             =   420
         Width           =   1215
      End
      Begin V81Affiliate.CSI_Calendar CalToDate 
         Height          =   285
         Left            =   2535
         TabIndex        =   23
         Top             =   735
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Text            =   "12/22/2022"
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
      Begin VB.Frame Frame5 
         Caption         =   "Weeks"
         Height          =   960
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3540
         Begin VB.OptionButton OptWks 
            Caption         =   "From"
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   10
            Top             =   225
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton OptWks 
            Caption         =   "All"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   225
            Width           =   795
         End
         Begin VB.Label LabTo 
            Caption         =   "End:"
            Height          =   240
            Left            =   2025
            TabIndex        =   9
            Top             =   585
            Width           =   450
         End
         Begin VB.Label labFrom 
            Caption         =   "Date- Start:"
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1050
         End
      End
      Begin V81Affiliate.CSI_Calendar CalEffActionDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Top             =   2520
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Text            =   "12/22/2022"
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   -1  'True
         CSI_InputBoxBoxAlignment=   0
         CSI_CalBackColor=   16777130
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
         CSI_CurDayForeColor=   0
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   0
      End
      Begin VB.Frame frcExpired 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2010
         TabIndex        =   30
         Top             =   2040
         Width           =   1545
         Begin VB.OptionButton rbcExpired 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   32
            Top             =   0
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton rbcExpired 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   31
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         TabIndex        =   33
         Top             =   2280
         Width           =   1545
         Begin VB.OptionButton rbcSuppressNotice 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   810
            TabIndex        =   36
            Top             =   0
            Width           =   585
         End
         Begin VB.OptionButton rbcSuppressNotice 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.Label lacSort 
         Caption         =   "Sort by"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label lacPostType 
         Caption         =   "Post Type"
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label lacDateDisclaimer 
         Caption         =   "(no date indicates all dates in past)"
         Height          =   210
         Left            =   120
         TabIndex        =   47
         Top             =   3840
         Width           =   2670
      End
      Begin VB.Label lacOldestNCRDate 
         Caption         =   "Oldest Delinquent Date to Include"
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lacSuppressNotice 
         Caption         =   "Honor 'Suppress Notices'"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1890
      End
      Begin VB.Label lacExpired 
         Caption         =   "Include Expired Agreements"
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   2040
         Width           =   2220
      End
      Begin VB.Label lacNewPage 
         Caption         =   "New Page Each Group"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1710
      End
      Begin VB.Label lacPassword 
         Caption         =   "Show"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label lacAction 
         Caption         =   "Action"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label lacEffecCommentDate 
         Caption         =   "Effective Date"
         Height          =   240
         Left            =   1680
         TabIndex        =   13
         Top             =   2520
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4740
      TabIndex        =   17
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4500
      TabIndex        =   16
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4305
      TabIndex        =   15
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffDelqRpt.frx":0E4A
         Left            =   1050
         List            =   "AffDelqRpt.frx":0E4C
         TabIndex        =   4
         Top             =   780
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Mail List"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1185
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   810
         Width           =   885
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   525
         Width           =   1110
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmDelqRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDelqRpt - Delinquent Report
'*
'*  Created July,1998 by Dick LeVine
'*  Modified May, 2000 by D. Smith
'*
'*  Note: Now allow the user to pick any date range rather than a 13 week maximum.
'*        All changes occured in the cmdReport_Click() routine and one type def below
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'Private Const cMaxDates = 11 ' max number of dates to print on a line in the report
'Private Const cMaxDates = 8 ' max number of dates to print on a line in the report
'Private Const cMaxDates = 7 ' max number of dates to print on a line in the report
'Private Const cMaxDates = 6 ' 12-3-10 max number of dates to print on a line in the report
Private Const cMaxDates = 5 ' 12-3-10 max number of dates to print on a line in the report

Private Type DELINQUENTINFO
    iStnCode As Integer
    iVehCode As Integer
    lPrintStatus As Long
    'sStartDate(1 To cMaxDates) As String * 10
    sStartDate(0 To cMaxDates) As String * 10
    'iPostingStatus(1 To cMaxDates) As Integer
    iPostingStatus(0 To cMaxDates) As Integer
End Type

'User delivery selection
Private Type DELIVERSELECTED
    iManual As Integer
    iNetWeb As Integer
    iCumulus As Integer
    iUni As Integer
    iMarketron As Integer
    iCBS As Integer
    iClearCh As Integer
End Type

Dim tmDlq() As DELINQUENTINFO
Dim tmDeliverSelected As DELIVERSELECTED
Dim tmCurrDeliveryWeek As DELIVERSELECTED      'Week for Agreement delivery definitions
Dim tmPrevDeliveryWeek As DELIVERSELECTED      'Week for Agreement delivery definitions

Private imChkListBoxIgnore As Integer
Private imChkDeliveryBoxIgnore As Integer
Private hmMail As Integer
Private smToFile As String
Dim imNoDaysDelq As Integer         'from site - # weeks before CP overdue
Dim imNoWksNCR As Integer          'from site - # weeks behind considered NCR
Dim imRptIndex As Integer
Dim imConsecutiveWksNCR As Integer
Dim lmDefaultStartDate As Long
Dim lmDefaultEndDate As Long
Dim smDefaultStartDate As String
Dim smDefaultEndDate As String
Dim smOldestNCRDate As String
Dim lmOldestNCRDate As Long '10-14-09 Oldest delinquent date to show
Dim imSortSelected As Integer           '5-18-19
'5-18-19 change radio button sort options to list box
Const SORT_VEHICLE = 0
Const SORT_STATION = 1
Const SORT_DMA_MKTNAME = 2
Const SORT_DMA_MKTRANK = 3
Const SORT_MKTREP = 4
Const SORT_PRODUCER = 5
Const SORT_SVCREP = 6
Const SORT_AUDP12 = 7
Const SORT_OWNER = 8



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
Private Function OpenMsgFile() As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim sLetter As String

    'On Error GoTo OpenMsgFileErr:
    sLetter = "A"
    Do
        ilRet = 0
        smToFile = sgExportDirectory & "D" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & sLetter & ".csv"
        'slDateTime = FileDateTime(smToFile)
        ilRet = gFileExist(smToFile)
        If ilRet = 0 Then
            sLetter = Chr$(Asc(sLetter) + 1)
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    'ilRet = 0
    'On Error GoTo OpenMsgFileErr:
    'hmMail = FreeFile
    'Open smToFile For Output As hmMail
    ilRet = gFileOpen(smToFile, "Output", hmMail)
    If ilRet <> 0 Then
        Close hmMail
        hmMail = -1
        gMsgBox "Open File " & smToFile & " error#" & Str$(Err.Number), vbOKOnly
        OpenMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    OpenMsgFile = True
    Exit Function
'OpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function


Private Sub CalOldestNCRDate_GotFocus()
    CalOldestNCRDate.ZOrder (vbBringToFront)
End Sub
'5-20-19
Private Sub cbcSort_Click()
'    If optSortby(0).Value = True Or optSortby(4).Value = True Or optSortby(5).Value = True Or optSortby(6).Value = True Then
    If cbcSort.ItemData(cbcSort.ListIndex) = SORT_VEHICLE Or cbcSort.ItemData(cbcSort.ListIndex) = SORT_DMA_MKTRANK Or cbcSort.ItemData(cbcSort.ListIndex) = SORT_PRODUCER Or cbcSort.ItemData(cbcSort.ListIndex) = SORT_SVCREP Then
        optPage(0).Value = True
    Else
        'default station, market and market name skip to new page off
        optPage(1).Value = True
    End If

End Sub

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllDelivery_Click()
 Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkDeliveryBoxIgnore Then
        Exit Sub
    End If
    If ckcAllDelivery.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcDelivery.ListCount > 0 Then
        imChkDeliveryBoxIgnore = True
        lRg = CLng(lbcDelivery.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcDelivery.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkDeliveryBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcInclComments_Click()
    If ckcInclComments.Value = 1 Then
        lacEffecCommentDate.Visible = True
        CalEffActionDate.Visible = True
        CalEffActionDate.SetEnabled (True)
    Else
        'lacEffecCommentDate.Visible = False
        'CalEffActionDate.Visible = False
        CalEffActionDate.SetEnabled (False)
    End If
    End Sub
Private Sub cmdDone_Click()
    Unload frmDelqRpt
End Sub

'       5-17-04 DH Option to honor "suppress notices".
'       7-16-04 Incorrect Agreement Contact & phone # showing due to the wrong
'               atfcode in grf record

Private Sub cmdReport_Click()
    Dim iTtlUnqDates, iIsFirst, iNumDates As Integer
    'change iArrayIndx from integer to long
    Dim lArrayIdx As Long
    Dim i, iRet, iDateIdx, iTtlDatesInRange As Integer
    Dim sVehicles, sStations, sMail As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sOutput As String
    Dim DelinRst As ADODB.Recordset
    Dim UpdateNCRRst As ADODB.Recordset
    Dim FutureNCRRst As ADODB.Recordset
    Dim sGenDate As String
    Dim sGenTime As String
    Dim iPrevStnCode As Integer
    Dim iPrevVehCode As Integer
    Dim sPrevNCRFlag As String      'NCR flag from agreement
    Dim sPrevFormerNCRFlag As String    'former NCR offender
    Dim iPrevPostingStatus As Integer
    Dim lPrevPrintCode As Long
    
    Dim sPrevStartDate As String * 10
    Dim dFWeek As Date
    Dim dOldNCRDate As Date
    Dim lSDate, lEDate As Long
    Dim sCurDate As String
    Dim ilExportType As Integer
    Dim slExportName As String
    Dim ilRptDest As Integer
    Dim slRptName As String         'report name for Overdue CP
    Dim slNCRRptName As String      'report name for Non-compliance
    Dim slMissWksRptName As String      'report name for Missing Wks
    Dim sNCRStartDate As String     'entered start date to be considered non-compliant
    Dim lNCRStartDate As Long       'entered start date to be considered non-compliant for date testing
    Dim ilAtLeastOneNCR As Integer     'at least one week is delinquent between the overdue and non-compliant period
    'Dim NewForm As New frmViewReport
    Dim sActionDate As String
    Dim dActionDate As Date
    Dim slSuppressNotice As String
    Dim iGotAnyData As Integer
    Dim slNCRFlag As String             'flag from attNCR if NCR flag is set
    Dim lLoop As Long
    ReDim llncragreements(0 To 0) As Long
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sAttEndDate As String
    Dim sBaseEndDate As String
    Dim llConsecutiveWkCount As Long
    Dim ilIncludeThisDate As Integer        'include the delinquent date on the report
                                            'the NCR report will exclude any delinquent dates prior to a user entered date
    Dim ilDeliveryTypeFound As Integer
    Dim blOneFoundToOutput As Boolean
    
    On Error GoTo ErrHand
       
    sCurDate = "1/1/01"
    Screen.MousePointer = vbHourglass
    'CRpt1.Connect = "DSN = " & sgDatabaseName
    
    If (imNoDaysDelq = 0 Or imNoWksNCR = 0 Or imConsecutiveWksNCR < 0 Or (imNoWksNCR = imNoDaysDelq)) And (imRptIndex <> AFFSMISSINGWKS_RPT) Then
        'error
        MsgBox "Verify Site: Invalid # Weeks Behind Considered Delinquent and/or # Weeks Behind Considered Non-compliant.  Non-compliant must be larger "
        Unload frmDelqRpt
        Set frmDelqRpt = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If ckcInclComments.Value = 1 Then               'include comments, need comment effective date
        If gIsDate(CalEffActionDate.Text) = False Then
            '5-1-17 force earliest date
            sActionDate = "1/1/1970"
            dActionDate = CDate(sActionDate)
        Else
            sActionDate = Format(CalEffActionDate.Text, "m/d/yyyy")
            dActionDate = CDate(sActionDate)

        End If
    End If
    
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmDelqRpt, "DelqRpt.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex    '3-15-04
        ilRptDest = 2
        If ilExportType = 3 Or ilExportType = 4 Then        '3-27-07 any excel export needs to force to No page skip to get headers in exported file
            optPage(1).Value = True                        'force to no page skipping
        End If
    ElseIf optRptDest(3).Value = True Then
        iRet = OpenMsgFile()
        If iRet = False Then
            Unload frmDelqRpt
            Exit Sub
        End If
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    
    'Retrieve information from the list box
    If OptWks(0).Value Then         'all dates, default to earliest/latest possible
        sStartDate = "1/1/1970"
        sEndDate = "12/31/2069"
    Else                            'from a specific date, if user blanked out dates,
                                    'take the entire range
        If imRptIndex = NCR_RPT Then    'if Non-compliant report, need to get all the dates where affidavits not returned so that
                                        'all dates will be shown if non-compliant.
            sStartDate = "1/1/1970"
            sNCRStartDate = CalFromDate.Text
            If sNCRStartDate = "" Then
                sNCRStartDate = "1/1/1970"
            End If
            sNCRStartDate = gAdjYear(Format$(DateValue(sNCRStartDate), "m/d/yyyy"))
            lNCRStartDate = DateValue(gAdjYear(sNCRStartDate))
            If CalOldestNCRDate.Text = "" Then
                smOldestNCRDate = "1/1/1970"
            Else
                smOldestNCRDate = CalOldestNCRDate.Text
            End If
            smOldestNCRDate = gAdjYear(Format$(DateValue(smOldestNCRDate), "m/d/yyyy"))
            lmOldestNCRDate = DateValue(gAdjYear(smOldestNCRDate))

        Else            'overdue CP or Affs Missing Weeks report
            sStartDate = CalFromDate.Text
            If sStartDate = "" Then
                sStartDate = "1/1/1970"
            End If
        End If
        
        sEndDate = CalToDate.Text
        If sEndDate = "" Then
            sEndDate = "12/31/2069"
        End If
        sStartDate = gAdjYear(Format$(gDateValue(sStartDate), "m/d/yyyy"))
        sEndDate = Format$(gDateValue(gAdjYear(sEndDate)), "m/d/yyyy")
    End If
    lSDate = DateValue(gAdjYear(sStartDate))
    lEDate = DateValue(gAdjYear(sEndDate))
    
    If imRptIndex = NCR_RPT Then
        If gDateValue(sNCRStartDate) > lEDate Then
            gMsgBox "Start Date must be prior to end date", vbOKOnly
            CalFromDate.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If sNCRStartDate = "1/1/1970" Then
             'at least start date or end date must be entered
             gMsgBox "Must enter start  date", vbOKOnly
             CalFromDate.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
         If sEndDate = "12/31/2069" Then
             'at least start date or end date must be entered
             gMsgBox "Must enter end date", vbOKOnly
             CalToDate.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
'        If imNoWksNCR = 0 Then
'            MsgBox "Warning: # of weeks behind considered non-compliant has not been entered in Site Options.  NCR flag will not be updated"
'            ckcUpdateNCR.Value = vbUnchecked
'            ckcUpdateNCR.Enabled = False
'        End If
       ' imConsecutiveWksNCR = (lEDate - lNCRStartDate) / 7         '4-24-17
        If lmDefaultStartDate <> lNCRStartDate Or lmDefaultEndDate <> lEDate Then
            ckcUpdateNCR.Value = vbUnchecked            'disallow the flag to be updated if changing the dates
            ckcUpdateNCR.Enabled = False
            MsgBox "Update NCR Flag has been disabled due to change in dates"
        Else
            ckcUpdateNCR.Enabled = True
            lacOldestNCRDate.Enabled = True     'oldest ncr date
            CalOldestNCRDate.SetEnabled (True)    'oldest ncr date
        End If
        If lmOldestNCRDate > lNCRStartDate Then     'disallow the oldest delinquent date to show greater than the earliest date considered delinquent
            MsgBox "Oldest date to include must be prior to start date"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
   If lbcVehAff.SelCount = 0 Then   'And imRptIndex <> AFFSMISSINGWKS_RPT Then
        Beep
        If optVehAff(0).Value Then          'vehicle
            gMsgBox "At least 1 vehicle must be selected"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            gMsgBox "At least 1 station must be selected"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
  
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'sDateRange = "(cpttStartDate >= '" + Format$(sStartDate, sgSQLDateForm) & "')"
    '7-8-09 take the entire date range so that report can get the total # of discrepant affidvits.
    'On the Overdue CP report, only the within range dates are printed, but total discrepant weeks are shown.
    'On the Non-compliance report, all weeks are shown.
    sDateRange = "(cpttStartDate >= '" + "1970-01-01" & "')"
    sDateRange = sDateRange & " And (cpttStartDate <= '" + Format$(sEndDate, sgSQLDateForm) & "')"
    sDateRange = sDateRange & " And (cpttStartDate >= attOnAir)"
    sDateRange = sDateRange & " And (cpttStartDate <= attOffAir)"
    sDateRange = sDateRange & " And (cpttStartDate <= attDropDate)"
    sVehicles = ""
    sStations = ""
    slSuppressNotice = ""
    slNCRFlag = ""
    
    slExportName = "DelqRpt"
    If imRptIndex = NCR_RPT Then
        slExportName = "NCRExport"
    ElseIf imRptIndex = AFFSMISSINGWKS_RPT Then
        slExportName = "MissingWks"
    End If
    
    mDeliverySelected
    
    ' Detrmine what to sort by
    ' Selecting by vehicles
    If optVehAff(0).Value = True Then
        If chkListBox.Value = 0 Then    '= 0 Then                        'User did NOT select all vehicles
            For i = 0 To lbcVehAff.ListCount - 1 Step 1
                If lbcVehAff.Selected(i) Then
                    If Len(sVehicles) = 0 Then
                        sVehicles = "(cpttVefCode = " & lbcVehAff.ItemData(i) & ")"
                    Else
                        sVehicles = sVehicles & " OR (cpttVefCode = " & lbcVehAff.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
    Else                                                               'User selected stations instead of vehicles
        If chkListBox.Value = 0 Then    '= 0 Then                        'User did NOT select all stations
            For i = 0 To lbcVehAff.ListCount - 1 Step 1
                If lbcVehAff.Selected(i) Then
                    If Len(sStations) = 0 Then
                        sStations = "(cpttShfCode = " & lbcVehAff.ItemData(i) & ")"
                    Else
                        sStations = sStations & " OR (cpttShfCode = " & lbcVehAff.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
    End If

    If rbcSuppressNotice(0).Value = True Then       'honor suppress notice?
        ' honor it, test to see if suppress notice set
        slSuppressNotice = " AND attSuppressNotice <> 'Y' "
    End If


    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

'    imSortSelected = cbcSort.ListIndex          '5-18-19
    imSortSelected = cbcSort.ItemData(cbcSort.ListIndex)

    'Generate the report
    If optRptDest(3).Value = True Then  'export the data
        If optVehAff(0).Value = True Then  ' select by Vehicle
            'SQLQuery = "SELECT *"
            SQLQuery = "SELECT DISTINCT  shttCallLetters, shttFax"
            SQLQuery = SQLQuery + " FROM VEF_Vehicles, att, cptt, shtt"
            SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
            SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
            If sVehicles <> "" Then
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
            'D.S. 11/29/01 Added for MAI
            If rbcExpired(1).Value Then  'If True don't show expired agreements
                'SQLQuery = SQLQuery + " AND attOffAir > " & Date
                'D.S. 2/12/02 Added format to Date
                SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
            End If
            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
            SQLQuery = SQLQuery + ")"
        Else          ' select by Stations
            'SQLQuery = "SELECT *"
            SQLQuery = "SELECT DISTINCT shttCallLetters, shttFax"
            SQLQuery = SQLQuery + " FROM VEF_Vehicles, cptt, shtt, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
            SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
            If sStations <> "" Then
                SQLQuery = SQLQuery + " AND (" & sStations & ")"
            End If
            'D.S. 11/29/01 Added for MAI
            If rbcExpired(1).Value Then  'If True don't show expired agreements
                'SQLQuery = SQLQuery + " AND attOffAir > " & Date
                'D.S. 2/12/02 Added format to Date
                SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
            End If
            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
            SQLQuery = SQLQuery + ")"
        End If
        On Error GoTo ErrHand
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            sMail = """" & Trim$(rst!shttCallLetters) & """" & "," & """" & "1-" & Trim$(rst!shttFax) & """"
            Print #hmMail, sMail
            rst.MoveNext
        Wend
        Close hmMail
        On Error GoTo 0
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True
        
        Screen.MousePointer = vbDefault
        gMsgBox "Output Sent To: " & smToFile, vbInformation
        Exit Sub
    ' The code below and to the end of the sub. has been modified modified! D.S. 5/23/00
    Else
        If optVehAff(0).Value = True Then
            ' Select by Vehicle
            SQLQuery = "SELECT cpttshfcode, cpttvefcode, cpttatfcode, cpttstartdate, cpttstatus, cpttpostingstatus, attonair, attoffair, attdropdate, attncr,attformerncr,attsuppressnotice,attcode,"
            SQLQuery = SQLQuery + "attExportToWeb, attExportType, attExportToUnivision, "
            SQLQuery = SQLQuery + "vefcode, vefsort, vefname, shttcallletters,shttfax, shttcode, shttmktcode, shttAudP12Plus, mktName, mktRank "
'            SQLQuery = SQLQuery + " FROM VEF_Vehicles, cptt, shtt, mkt, att, artt "
'            SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
'            SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
'            SQLQuery = SQLQuery + " AND shttmktCode = mktCode"
'            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            '5-21-19 replace above with the joins
            SQLQuery = SQLQuery + " From cptt INNER JOIN VEF_Vehicles on cpttvefcode = vefcode "
            SQLQuery = SQLQuery & " INNER JOIN shtt on cpttshfcode = shttcode "
            SQLQuery = SQLQuery & " INNER JOIN att on cpttatfcode = attcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN mkt on shttMktCode = mktCode "
            '6-2-16 Partial or Unposted weeks do not apply for missing weeks report
            If imRptIndex = AFFSMISSINGWKS_RPT Then     '1-12-10
'                SQLQuery = SQLQuery + " AND cpttStatus = 2 and cpttPostingStatus = 2" + slSuppressNotice
                SQLQuery = SQLQuery + " where ( cpttStatus = 2 and cpttPostingStatus = 2 and attServiceAgreement <> 'Y' " + slSuppressNotice
            Else        'overdue or ncr
                If imRptIndex = NCR_RPT Then
'                    SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                    SQLQuery = SQLQuery + " where ( cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                Else                'overdue cp, select partial posted, completedly unposted, or both
                    If rbcPostType(0).Value = True Then
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and cpttPostingStatus = 1  and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and cpttPostingStatus = 1  and attServiceAgreement <> 'Y' " + slSuppressNotice
                    ElseIf rbcPostType(1).Value = True Then     'nothing posted
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and cpttPostingStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and cpttPostingStatus = 0  and attServiceAgreement <> 'Y' " + slSuppressNotice
                    Else
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                    End If
                End If
            End If
              
            'SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
           
            'D.S. 11/29/01 Added for MAI
            If rbcExpired(1).Value Then  'If True don't show expired agreements
                'SQLQuery = SQLQuery + " AND attOffAir > " & Date
                'D.S. 2/12/02 Added format to Date
                SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
            End If
            If sVehicles <> "" Then
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"

'            If optSortby(0).Value = True Then      'sort by vehicle
'                SQLQuery = SQLQuery + ")" + " ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(1).Value = True Then      'sort by Station
'                SQLQuery = SQLQuery + ")" + " ORDER BY shttCallLetters, vefSort, vefName, cpttStartDate"
'            ElseIf optSortby(2).Value = True Then      'sort by Market Name
'                SQLQuery = SQLQuery + ")" + " ORDER BY mktName, shttCallLetters, vefSort, vefName, mktName, cpttStartDate"
'            ElseIf optSortby(3).Value = True Then      'sort by Market Rank
'                SQLQuery = SQLQuery + ")" + " ORDER BY mktRank, shttCallLetters, vefSort, vefName, mktRank, cpttStartDate"
'            ElseIf optSortby(4).Value = True Then       'sort by mkt rep A/E
'                SQLQuery = SQLQuery + ")" + "  ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(5).Value = True Then        'producer
'                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(6).Value = True Then        'sort by service rep
'                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(7).Value = True Then         'audience
'                SQLQuery = SQLQuery + ")" + " ORDER BY shttAudP12Plus desc, shttCallLetters, vefSort, vefName, cpttStartDate"
'            End If

            '6-7-19 no need to send the sort to crystal
            If imSortSelected = SORT_VEHICLE Then      'sort by vehicle
                SQLQuery = SQLQuery + ")" + " ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_STATION Then       'sort by Station
                SQLQuery = SQLQuery + ")" + " ORDER BY shttCallLetters, vefSort, vefName, cpttStartDate"
            ElseIf imSortSelected = SORT_DMA_MKTNAME Then       'sort by Market Name
                SQLQuery = SQLQuery + ")" + " ORDER BY mktName, shttCallLetters, vefSort, vefName, mktName, cpttStartDate"
            ElseIf imSortSelected = SORT_DMA_MKTRANK Then       'sort by Market Rank
                SQLQuery = SQLQuery + ")" + " ORDER BY mktRank, shttCallLetters, vefSort, vefName, mktRank, cpttStartDate"
            ElseIf imSortSelected = SORT_MKTREP Then       'sort by mkt rep A/E
                SQLQuery = SQLQuery + ")" + "  ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_PRODUCER Then         'producer
                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_SVCREP Then         'sort by service rep
                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_AUDP12 Then          'audience
                SQLQuery = SQLQuery + ")" + " ORDER BY shttAudP12Plus desc, shttCallLetters, vefSort, vefName, cpttStartDate"
           ElseIf imSortSelected = SORT_OWNER Then          'Owner
                SQLQuery = SQLQuery + ")" + " ORDER BY   vefSort, vefName, shttCallLetters, cpttStartDate"
            End If
            

        Else            'select by station
            SQLQuery = "SELECT *"
'            SQLQuery = SQLQuery + " FROM VEF_Vehicles, cptt, shtt, mkt, att, artt "
'            SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
'            SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
'            SQLQuery = SQLQuery + " AND shttmktCode = mktCode"
'            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " From cptt INNER JOIN VEF_Vehicles on cpttvefcode = vefcode "
            SQLQuery = SQLQuery & " INNER JOIN shtt on cpttshfcode = shttcode "
            SQLQuery = SQLQuery & " INNER JOIN att on cpttatfcode = attcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN mkt on shttMktCode = mktCode "
            'SQLQuery = SQLQuery + " AND attarttCode = arttCode"
            'SQLQuery = SQLQuery + " AND attPrintCP = cptt.cpttprintstatus"
            
            If imRptIndex = AFFSMISSINGWKS_RPT Then     '1-12-10
'                SQLQuery = SQLQuery + " AND cpttStatus = 2 and cpttPostingStatus = 2 " + slSuppressNotice
                SQLQuery = SQLQuery + " where ( cpttStatus = 2 and cpttPostingStatus = 2 and attServiceAgreement <> 'Y' " + slSuppressNotice
            Else        'overdue or ncr
                If imRptIndex = NCR_RPT Then
'                    SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                    SQLQuery = SQLQuery + " where ( cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                Else                'overdue cp, select partial posted, completedly unposted, or both
                    If rbcPostType(0).Value = True Then
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and cpttPostingStatus = 1  and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and cpttPostingStatus = 1  and attServiceAgreement <> 'Y' " + slSuppressNotice
                    ElseIf rbcPostType(1).Value = True Then     'nothing posted
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and cpttPostingStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and cpttPostingStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                    Else
'                        SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                        SQLQuery = SQLQuery + " where ( cpttStatus = 0 and attServiceAgreement <> 'Y' " + slSuppressNotice
                    End If
                End If
            End If
            
            'D.S. 11/29/01 Added for MAI
            If rbcExpired(1).Value Then  'If True don't show expired agreements
                'SQLQuery = SQLQuery + " AND attOffAir > " & Date
                'D.S. 2/12/02 Added format to Date
                SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
            End If
            If sStations <> "" Then
                SQLQuery = SQLQuery + " AND (" & sStations & ")"
            End If

            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"

'            If optSortby(0).Value = True Then      'sort by vehicle
'                SQLQuery = SQLQuery + ")" + " ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(1).Value = True Then      'sort by Station
'                SQLQuery = SQLQuery + ")" + " ORDER BY shttCallLetters, vefSort, vefName, cpttStartDate"
'            ElseIf optSortby(2).Value = True Then      'sort by Market Name
'                SQLQuery = SQLQuery + ")" + " ORDER BY mktName, shttCallLetters, vefSort, vefName, cpttStartDate"        'take out shttmarket
'            ElseIf optSortby(3).Value = True Then      'sort by Market Rank
'                SQLQuery = SQLQuery + ")" + "  ORDER BY mktRank, shttCallLetters, vefSort, vefName,  cpttStartDate"       'take out shttrank
'            ElseIf optSortby(4).Value = True Then       'sort by mkt rep A/E
'                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(5).Value = True Then            'sort by producer
'                SQLQuery = SQLQuery + ")" + " ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(6).Value = True Then       'service rep, sort by affiliat A/E
'                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
'            ElseIf optSortby(7).Value = True Then           'audience p12+
'                SQLQuery = SQLQuery + ")" + " ORDER BY shttAudP12Plus desc, shttCallLetters, vefSort, vefName, cpttStartDate"
'            End If

            '6-7-19 no need to send sort in sql call to crystal
            If imSortSelected = SORT_VEHICLE Then      'sort by vehicle
                SQLQuery = SQLQuery + ")" + " ORDER BY vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_STATION Then      'sort by Station
                SQLQuery = SQLQuery + ")" + " ORDER BY shttCallLetters, vefSort, vefName, cpttStartDate"
            ElseIf imSortSelected = SORT_DMA_MKTNAME Then      'sort by Market Name
                SQLQuery = SQLQuery + ")" + " ORDER BY mktName, shttCallLetters, vefSort, vefName, cpttStartDate"        'take out shttmarket
            ElseIf imSortSelected = SORT_DMA_MKTRANK Then      'sort by Market Rank
                SQLQuery = SQLQuery + ")" + "  ORDER BY mktRank, shttCallLetters, vefSort, vefName,  cpttStartDate"       'take out shttrank
            ElseIf imSortSelected = SORT_MKTREP Then       'sort by mkt rep A/E
                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_PRODUCER Then            'sort by producer
                SQLQuery = SQLQuery + ")" + " ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_SVCREP Then       'service rep, sort by affiliat A/E
                SQLQuery = SQLQuery + ")" + "  ORDER by vefSort, vefName, shttCallLetters, cpttStartDate"
            ElseIf imSortSelected = SORT_AUDP12 Then           'audience p12+
                SQLQuery = SQLQuery + ")" + " ORDER BY shttAudP12Plus desc, shttCallLetters, vefSort, vefName, cpttStartDate"
            ElseIf imSortSelected = SORT_OWNER Then           'Owner
                SQLQuery = SQLQuery + ")" + " ORDER BY   vefSort, vefName,shttCallLetters, cpttStartDate"
            End If

        End If
    End If

    blOneFoundToOutput = False  'no records found to output yet
    ' Begin inserting into the pre-pass GRF table for Crystal Reports
    iTtlDatesInRange = 0
    iGotAnyData = False     'assume no data found yet until something written to prepass file
    Set DelinRst = gSQLSelectCall(SQLQuery)
    If Not DelinRst.EOF Then
        iIsFirst = True
        iNumDates = -1      '4-27-17 0
        lArrayIdx = 1
        iTtlUnqDates = 0
        ilAtLeastOneNCR = False
        llConsecutiveWkCount = 0
        sGenDate = Format$(gNow(), "m/d/yyyy")
        sGenTime = Format$(gNow(), sgShowTimeWSecForm)
        sStartDate = Format(sStartDate, "m/d/yyyy")
        While Not DelinRst.EOF
            'determine which export types have been selected
            'For Missing Weeks report, all delivery types have been pre-selected. there is no option
            ilDeliveryTypeFound = False
            'the info in this type defintion will be updated into the GRF for type of delivery options
            tmCurrDeliveryWeek.iCBS = 0
            tmCurrDeliveryWeek.iClearCh = 0
            tmCurrDeliveryWeek.iCumulus = 0
            tmCurrDeliveryWeek.iManual = 0
            tmCurrDeliveryWeek.iMarketron = 0
            tmCurrDeliveryWeek.iNetWeb = 0
            tmCurrDeliveryWeek.iUni = 0
            If DelinRst!attExportType = 0 Then      'client posts manually
                If tmDeliverSelected.iManual Then
                    ilDeliveryTypeFound = True
                    tmCurrDeliveryWeek.iManual = 1
                End If
            '7701
            Else            'some form of export (web, cumulus, uni, etc)
                If tmDeliverSelected.iCBS Then
                    If gIsVendorWithAgreement(DelinRst!attCode, Vendors.cBs) Then
                        tmCurrDeliveryWeek.iCBS = 1
                        ilDeliveryTypeFound = True
                    End If
                End If
                If tmDeliverSelected.iClearCh Then
                    If gIsVendorWithAgreement(DelinRst!attCode, Vendors.iheart) Then
                        ilDeliveryTypeFound = True
                        tmCurrDeliveryWeek.iClearCh = 1
                    End If
                End If
                If tmDeliverSelected.iCumulus Then
                    If gIsVendorWithAgreement(DelinRst!attCode, Vendors.stratus) Then
                        ilDeliveryTypeFound = True
                        tmCurrDeliveryWeek.iCumulus = 1
                    End If
                End If
                If tmDeliverSelected.iMarketron Then
                    If gIsVendorWithAgreement(DelinRst!attCode, Vendors.NetworkConnect) Then
                        ilDeliveryTypeFound = True
                        tmCurrDeliveryWeek.iMarketron = 1
                    End If
                End If
                If tmDeliverSelected.iNetWeb And DelinRst!attExportToWeb = "Y" Then
                    ilDeliveryTypeFound = True
                    tmCurrDeliveryWeek.iNetWeb = 1
                End If
'                If tmDeliverSelected.iCBS And gIfNullInteger(DelinRst!vatWvtIdCodeLog) = Vendors.cBs Then
'                    tmCurrDeliveryWeek.iCBS = 1
'                    ilDeliveryTypeFound = True
'                End If
'                If tmDeliverSelected.iClearCh And gIfNullInteger(DelinRst!vatWvtIdCodeLog) = Vendors.ClearChannel Then
'                    ilDeliveryTypeFound = True
'                    tmCurrDeliveryWeek.iClearCh = 1
'                End If
'                If tmDeliverSelected.iCumulus And gIfNullInteger(DelinRst!vatWvtIdCodeLog) = Vendors.Cumulus Then
'                    ilDeliveryTypeFound = True
'                    tmCurrDeliveryWeek.iCumulus = 1
'                End If
'                If tmDeliverSelected.iMarketron And gIfNullInteger(DelinRst!vatWvtIdCodeLog) = Vendors.NetworkConnect Then
'                    ilDeliveryTypeFound = True
'                    tmCurrDeliveryWeek.iMarketron = 1
'                End If
'                If tmDeliverSelected.iNetWeb And DelinRst!attExportToWeb = "Y" Then
'                    ilDeliveryTypeFound = True
'                    tmCurrDeliveryWeek.iNetWeb = 1
'                End If
'                '7701 removed
''                If tmDeliverSelected.iUni And DelinRst!attExportToUnivision = "Y" Then
''                    ilDeliveryTypeFound = True
''                    tmCurrDeliveryWeek.iUni = 1
''                End If

                '6-7-19 re-establish the test for Uni; vendor method not implemented
                If tmDeliverSelected.iUni And DelinRst!attExportToUnivision = "Y" Then
                    ilDeliveryTypeFound = True
                    tmCurrDeliveryWeek.iUni = 1
                End If

            End If
            If ilDeliveryTypeFound Then
                blOneFoundToOutput = True           'at least one delinquent agreement found to output
                'Build arrary of unique date records
                If iIsFirst = True Then
                    sPrevStartDate = Format(DelinRst!CpttStartDate, "m/d/yyyy")
                    iPrevStnCode = DelinRst!cpttshfcode
                    iPrevVehCode = DelinRst!cpttvefcode
                    lPrevPrintCode = DelinRst!cpttatfCode
                    sPrevNCRFlag = DelinRst!attNCR
                    sPrevFormerNCRFlag = DelinRst!attFormerNCR
                    iPrevPostingStatus = DelinRst!cpttPostingStatus
                    tmPrevDeliveryWeek = tmCurrDeliveryWeek
                    iNumDates = iNumDates + 1
                    iTtlUnqDates = iTtlUnqDates + iNumDates
                    lArrayIdx = 0
                    ReDim tmDlq(0 To 0) As DELINQUENTINFO
                    'For i = 1 To cMaxDates Step 1 ' Init the date field
                    For i = 0 To cMaxDates Step 1 ' Init the date field
                        If igSQLSpec = 0 Then
                            tmDlq(0).sStartDate(i) = "0/0/0000"
                        Else
                            tmDlq(0).sStartDate(i) = "1/1/1970"
                        End If
                        tmDlq(0).iPostingStatus(i) = 0                       'assume unposted (vs partial)
                    Next
                    iIsFirst = False
                    
                'End If         '8-20-09
                Else            '8-20-09
                    If (DelinRst!cpttshfcode = iPrevStnCode) And (DelinRst!cpttvefcode = iPrevVehCode) Then
         
                    ' While the station and veh codes are the same gather them as a group
                        'test to see if the delinquent affidavit (cptt) is within the parameters of the request
                        'For the NCR version, the earliest date has been forced to the beginning of time since
                        'all delinquent cptt dates will be shown if its an non-compliant agreement
                        If ((DateValue(gAdjYear(sPrevStartDate)) >= lSDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                            'for the NCR version, see if this record falls within the requested period
                            If ((DateValue(gAdjYear(sPrevStartDate)) >= lNCRStartDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                                ilAtLeastOneNCR = True
                                llConsecutiveWkCount = llConsecutiveWkCount + 1
                            End If
                            iTtlDatesInRange = iTtlDatesInRange + 1
                            If (imRptIndex = NCR_RPT And (DateValue(gAdjYear(sPrevStartDate)) >= lmOldestNCRDate)) Or imRptIndex <> NCR_RPT Then                'dont include dates prior to oldest delinquent date to include
                                tmDlq(lArrayIdx).sStartDate(iNumDates) = sPrevStartDate
                                tmDlq(lArrayIdx).iPostingStatus(iNumDates) = iPrevPostingStatus
                                tmDlq(lArrayIdx).lPrintStatus = lPrevPrintCode
                                
    '                            lPrevPrintCode = DelinRst!cpttatfCode
    '                            iPrevStnCode = DelinRst!cpttshfCode
    '                            sPrevStartDate = Format(DelinRst!CpttStartDate, "m/d/yyyy")
    '                            iPrevVehCode = DelinRst!cpttvefCode
    '                            sPrevNCRFlag = DelinRst!attNCR
    '                            sPrevFormerNCRFlag = DelinRst!attFormerNCR
            
                                iNumDates = iNumDates + 1
    '                            iTtlUnqDates = iTtlUnqDates + 1
    '                            If iNumDates > cMaxDates Then
    '                                lArrayIdx = lArrayIdx + 1
    '                                iNumDates = 1
    '                                ReDim Preserve tmDlq(0 To lArrayIdx) As DELINQUENTINFO
    '                                For i = 1 To cMaxDates Step 1
    '                                    If igSQLSpec = 0 Then
    '                                        tmDlq(lArrayIdx).sStartDate(i) = "0/0/0000"
    '                                    Else
    '                                        tmDlq(lArrayIdx).sStartDate(i) = "1/1/1970"
    '                                    End If
    '                                Next
    '                            End If
                            End If
                            lPrevPrintCode = DelinRst!cpttatfCode
                            iPrevStnCode = DelinRst!cpttshfcode
                            sPrevStartDate = Format(DelinRst!CpttStartDate, "m/d/yyyy")
                            iPrevVehCode = DelinRst!cpttvefcode
                            sPrevNCRFlag = DelinRst!attNCR
                            sPrevFormerNCRFlag = DelinRst!attFormerNCR
                            iPrevPostingStatus = DelinRst!cpttPostingStatus
                            tmPrevDeliveryWeek = tmCurrDeliveryWeek
                            iTtlUnqDates = iTtlUnqDates + 1
                            If iNumDates > cMaxDates Then
                                lArrayIdx = lArrayIdx + 1
                                'iNumDates = 1
                                iNumDates = 0      '4-27-17 0
                                ReDim Preserve tmDlq(0 To lArrayIdx) As DELINQUENTINFO
                                'For i = 1 To cMaxDates Step 1
                                For i = 0 To cMaxDates Step 1
                                    If igSQLSpec = 0 Then
                                        tmDlq(lArrayIdx).sStartDate(i) = "0/0/0000"
                                    Else
                                        tmDlq(lArrayIdx).sStartDate(i) = "1/1/1970"
                                    End If
                                    tmDlq(lArrayIdx).iPostingStatus(i) = 0
                                Next
                            End If
                        Else
                            sPrevStartDate = Format(DelinRst!CpttStartDate, "m/d/yyyy")
                            iPrevPostingStatus = DelinRst!cpttPostingStatus
                            tmPrevDeliveryWeek = tmCurrDeliveryWeek
                            iTtlUnqDates = iTtlUnqDates + 1
                        End If
                    Else
                    ' We have detetectd a change in either veh or stn. code - write out a GRF rec
                        If ((DateValue(gAdjYear(sPrevStartDate)) >= lSDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                            iTtlDatesInRange = iTtlDatesInRange + 1
                            tmDlq(lArrayIdx).sStartDate(iNumDates) = sPrevStartDate
                            tmDlq(lArrayIdx).lPrintStatus = lPrevPrintCode
                            tmDlq(lArrayIdx).iPostingStatus(iNumDates) = iPrevPostingStatus
                             'for the NCR version, see if this record falls within the requested period
                            If ((DateValue(gAdjYear(sPrevStartDate)) >= lNCRStartDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                                ilAtLeastOneNCR = True
                                llConsecutiveWkCount = llConsecutiveWkCount + 1
                            End If
                        End If
                        
                         If imRptIndex = NCR_RPT Then
                             'if the # weeks delinquent is greater than the # considered non-compliant, show it on report
                             'OR show the station if it already is non-compliant
                             'If ilAtLeastOneNCR Or sPrevNCRFlag = "Y" Then    'vehicle & station already non-compliant, or the # weeks delinquent has caused them to be non-compliant
                             If llConsecutiveWkCount >= imConsecutiveWksNCR Or sPrevNCRFlag = "Y" Then    'vehicle & station already non-compliant, or the # weeks delinquent has caused them to be non-compliant
    
                                 InsertIntoGRF lArrayIdx, sGenDate, sGenTime, iPrevStnCode, iPrevVehCode, lPrevPrintCode, iTtlUnqDates, iTtlDatesInRange, sPrevNCRFlag, sPrevFormerNCRFlag, iPrevPostingStatus
                                 ilAtLeastOneNCR = False
                                 llConsecutiveWkCount = 0
                                 llncragreements(UBound(llncragreements)) = lPrevPrintCode
                                 ReDim Preserve llncragreements(0 To UBound(llncragreements) + 1) As Long
                             Else
                                 ilAtLeastOneNCR = False
                                 llConsecutiveWkCount = 0
                             End If
                         Else
                             InsertIntoGRF lArrayIdx, sGenDate, sGenTime, iPrevStnCode, iPrevVehCode, lPrevPrintCode, iTtlUnqDates, iTtlDatesInRange, sPrevNCRFlag, sPrevFormerNCRFlag, iPrevPostingStatus
                         End If
                        'End If
                        iGotAnyData = True          'at least 1 record found to report
                        iTtlDatesInRange = 0
                        iPrevStnCode = DelinRst!cpttshfcode
                        iPrevVehCode = DelinRst!cpttvefcode
                        lPrevPrintCode = DelinRst!cpttatfCode           '7-16-04 new previous agreement code
                        sPrevStartDate = Format(DelinRst!CpttStartDate, "m/d/yyyy")
                        sPrevNCRFlag = DelinRst!attNCR
                        sPrevFormerNCRFlag = DelinRst!attFormerNCR
                        iPrevPostingStatus = DelinRst!cpttPostingStatus
                        tmPrevDeliveryWeek = tmCurrDeliveryWeek
                        'iNumDates = 1
                        iNumDates = 0
                        lArrayIdx = 0
                        ReDim tmDlq(0 To 0) As DELINQUENTINFO
                        'For i = 1 To cMaxDates Step 1 ' Init the date field
                        For i = 0 To cMaxDates Step 1 ' Init the date field
                            If igSQLSpec = 0 Then
                                tmDlq(0).sStartDate(i) = "0/0/0000"
                            Else
                                tmDlq(0).sStartDate(i) = "1/1/1970"
                            End If
                            tmDlq(0).iPostingStatus(i) = 0
                        Next
                        iTtlUnqDates = 1
                    End If
                End If
            End If                  'ildeliveryTypeFound
            DelinRst.MoveNext
        Wend
    
        If blOneFoundToOutput Then          '4-9-15
            ' We are one record behind - insert the last record before we call Crystal
            If ((DateValue(gAdjYear(sPrevStartDate)) >= lSDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                If ((DateValue(gAdjYear(sPrevStartDate)) >= lNCRStartDate) And (DateValue(gAdjYear(sPrevStartDate)) <= lEDate)) Then
                    ilAtLeastOneNCR = True
                    llConsecutiveWkCount = llConsecutiveWkCount + 1
                End If
                iTtlDatesInRange = iTtlDatesInRange + 1
                tmDlq(lArrayIdx).sStartDate(iNumDates) = sPrevStartDate
                tmDlq(lArrayIdx).lPrintStatus = lPrevPrintCode
                tmDlq(lArrayIdx).iPostingStatus(iNumDates) = iPrevPostingStatus
            End If
            If imRptIndex = NCR_RPT Then
                'if the # weeks delinquent is greater than the # considered non-compliant, show it on report
                'OR show the station if it already is non-compliant
                'If ilAtLeastOneNCR Or sPrevNCRFlag = "Y" Then    'vehicle & station already non-compliant, or the # weeks delinquent has caused them to be non-compliant
                If llConsecutiveWkCount >= imConsecutiveWksNCR Or sPrevNCRFlag = "Y" Then    'vehicle & station already non-compliant, or the # weeks delinquent has caused them to be non-compliant
                    
                    InsertIntoGRF lArrayIdx, sGenDate, sGenTime, iPrevStnCode, iPrevVehCode, lPrevPrintCode, iTtlUnqDates, iTtlDatesInRange, sPrevNCRFlag, sPrevFormerNCRFlag, iPrevPostingStatus
                    iGotAnyData = True
                    llncragreements(UBound(llncragreements)) = lPrevPrintCode
                    ReDim Preserve llncragreements(0 To UBound(llncragreements) + 1) As Long
                End If
            Else
                InsertIntoGRF lArrayIdx, sGenDate, sGenTime, iPrevStnCode, iPrevVehCode, lPrevPrintCode, iTtlUnqDates, iTtlDatesInRange, sPrevNCRFlag, sPrevFormerNCRFlag, iPrevPostingStatus
                iGotAnyData = True          'at least 1 record found to report
            End If
        End If
        
'        If imRptIndex = AFFSMISSINGWKS_RPT Then
'            slRptName = "AfMissWks.rpt"
'        Else
            If imSortSelected = SORT_VEHICLE Then
'            If optSortby(0) = True Then     ' By Vehicle
                slRptName = "afdelqvh.rpt"
                slNCRRptName = "afNCRvh.rpt"
                slMissWksRptName = "AfMissWksVh.rpt"
                
                sgCrystlFormula6 = "V"
            ElseIf imSortSelected = SORT_STATION Then
'            ElseIf optSortby(1) = True Then ' By Station
                slRptName = "afdelqst.rpt"
                slNCRRptName = "afNCRst.rpt"
                slMissWksRptName = "AfMissWks.rpt"
                sgCrystlFormula6 = "S"
                'slRptName = "afdelqst.rpt"
            ElseIf imSortSelected = SORT_DMA_MKTNAME Then
'            ElseIf optSortby(2) = True Then ' By Market
                slRptName = "afdelqst.rpt"
                slNCRRptName = "afNCRst.rpt"
                slMissWksRptName = "AfMissWks.rpt"
                sgCrystlFormula6 = "M"
                'slRptName = "afdelqmn.rpt"
            ElseIf imSortSelected = SORT_DMA_MKTRANK Then
'            ElseIf optSortby(3) = True Then ' By Market Rank
                slRptName = "afdelqst.rpt"
                slNCRRptName = "afNCRst.rpt"
                slMissWksRptName = "AfMissWks.rpt"
                sgCrystlFormula6 = "R"
                'slRptName = "afdelqmr.rpt"
            ElseIf imSortSelected = SORT_MKTREP Then
'            ElseIf optSortby(4) = True Then      'affiliate a/e
                slRptName = "afdelqvh.rpt"
                slNCRRptName = "afNCRvh.rpt"
                slMissWksRptName = "AfMissWksVh.rpt"
                sgCrystlFormula6 = "A"
            ElseIf imSortSelected = SORT_PRODUCER Then
'            ElseIf optSortby(5) = True Then     '3-25-09 Producer
                slRptName = "afdelqPro.rpt"
                slNCRRptName = "afNCRPro.rpt"
                slMissWksRptName = "AfMissWksPro.rpt"
                sgCrystlFormula6 = "P"
            ElseIf imSortSelected = SORT_SVCREP Then
'            ElseIf optSortby(6) = True Then     '8-8-14 Service rep
                slRptName = "afdelqvh.rpt"
                slNCRRptName = "afNCRvh.rpt"
                slMissWksRptName = "AfMissWksVh.rpt"
                sgCrystlFormula6 = "C"
            ElseIf imSortSelected = SORT_AUDP12 Then
'            ElseIf optSortby(7) = True Then     '11-13-14 audience
                slRptName = "afdelqAud.rpt"
                'slNCRRptName = "afNCRAud.rpt"
                slMissWksRptName = "AfMissWksAud.rpt"
                sgCrystlFormula6 = "A"
            ElseIf imSortSelected = SORT_OWNER Then     '5-18-19  Owner
                slRptName = "afdelqPro.rpt"
                slNCRRptName = "afNCRPro.rpt"
                slMissWksRptName = "AfMissWksPro.rpt"
                sgCrystlFormula6 = "O"

            End If
'        End If
        
        'Prepare records (format and formulas) to pass to Crystal
        sGenTime = Format(sGenTime, "hh:mm:ss")
        
        If imRptIndex = NCR_RPT Then
            dFWeek = CDate(sNCRStartDate)
            dOldNCRDate = CDate(smOldestNCRDate)
            sgCrystlFormula10 = "Date(" + Format$(smOldestNCRDate, "yyyy") + "," + Format$(smOldestNCRDate, "mm") + "," + Format$(smOldestNCRDate, "dd") + ")"
        Else
            dFWeek = CDate(sStartDate)
        End If
        
        'CRpt1.Formulas(0) = "StartDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        dFWeek = CDate(sEndDate)
        'CRpt1.Formulas(1) = "EndDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        If optPage(0).Value = True Then
            'CRpt1.Formulas(2) = "NewPage = 'Y'"
            sgCrystlFormula3 = "Y"
        Else
            'CRpt1.Formulas(2) = "NewPage = 'N'"
            sgCrystlFormula3 = "N"
        End If

        If ckcInclComments.Value = vbChecked Then        'include action comments
            sgCrystlFormula4 = "Y"
        Else
            sgCrystlFormula4 = "N"
        End If
        sgCrystlFormula5 = "Date(" + Format$(dActionDate, "yyyy") + "," + Format$(dActionDate, "mm") + "," + Format$(dActionDate, "dd") + ")"

        If rbcPassword(0) = True Then     ' By Vehicle
            sgCrystlFormula7 = "F"          'show fax # (vs password)
        Else
            sgCrystlFormula7 = "P"          'show pasword (vs fax)
        End If
        
        
        '8-20-09 Update flag to show on report
        If ckcUpdateNCR.Value = vbChecked Then
            sgCrystlFormula8 = "Y"
        Else
            sgCrystlFormula8 = "N"
        End If
        
        '8-20-09 Honor Suppress notice
        If rbcSuppressNotice(0).Value = True Then       'honor suppress notice?
            sgCrystlFormula9 = "Y"
        Else
            sgCrystlFormula9 = "N"
        End If
        
        '11-14-14 Partial, unposted or both
        If rbcPostType(0).Value = True Then       'partial only
            sgCrystlFormula11 = "P"
        ElseIf rbcPostType(1).Value = True Then        'unposted only
            sgCrystlFormula11 = "U"
        Else
            sgCrystlFormula11 = "B"                 'include any type of unposted week
        End If

        'Prepare records to pass to Crystal
        SQLQuery = "SELECT * from GRF_Generic_Report "
        SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles on grfsofCode = vefCode "
        SQLQuery = SQLQuery & "INNER JOIN shtt on shttCode = grfCode2 "
        SQLQuery = SQLQuery & "INNER JOIN att on grfPer3 = attCode "
        SQLQuery = SQLQuery & "LEFT OUTER JOIN mkt on shttMktCode = mktCode "
        SQLQuery = SQLQuery & "left outer join artt on grfPer5 = arttcode "
        SQLQuery = SQLQuery & "left outer join ust on grfPer4 = ustcode "

        SQLQuery = SQLQuery & "INNER JOIN VPF_Vehicle_Options on grfsofcode = vpfvefkCode "
        SQLQuery = SQLQuery & "LEFT OUTER JOIN arf_Addresses on vpfProducerArfCode = arfCode "
        
        If ((imRptIndex = OVERDUE_RPT Or imRptIndex = NCR_RPT Or imRptIndex = AFFSMISSINGWKS_RPT) And (imSortSelected = SORT_PRODUCER Or imSortSelected = SORT_OWNER)) Then
            SQLQuery = SQLQuery & " LEFT OUTER JOIN artt arttOwner on shttOwnerArttCode = arttOwner.ArttCode "
        End If


        'SQLQuery = SQLQuery & "GRF_Generic_Report, shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode "
        'SQLQuery = SQLQuery + " WHERE (vefCode = grfsofCode"
        'SQLQuery = SQLQuery + " AND shttCode = grfCode2"
        'SQLQuery = SQLQuery + " AND attCode = grfPer3"
 
        SQLQuery = SQLQuery + " where (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        
        
        'CRpt1.SQLQuery = SQLQuery
        'CRpt1.Action = 1           'make the call to crystal
        
        gUserActivityLog "E", sgReportListName & ": Prepass"
        If imRptIndex = NCR_RPT Then
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slNCRRptName, slExportName
        ElseIf imRptIndex = OVERDUE_RPT Then         'overdue
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
        Else                    'missing weeks
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slMissWksRptName, slExportName
        End If
        ' Kill the Crystal formulas
        'CRpt1.Formulas(0) = ""
        'CRpt1.Formulas(1) = ""
        'CRpt1.Formulas(2) = ""
        
        '7-6-00 update the NCR flags
        If imRptIndex = NCR_RPT And ckcUpdateNCR.Value = vbChecked Then
            For lLoop = 0 To UBound(llncragreements) - 1
                SQLQuery = "Select * from att where attcode = " & llncragreements(lLoop)
                Set UpdateNCRRst = gSQLSelectCall(SQLQuery)
                sOffAir = Format$(UpdateNCRRst!attOffAir, "mm/dd/yyyy")
                sDropDate = Format$(UpdateNCRRst!attDropDate, "mm/dd/yyyy")
                'determine the earliest of 2 dates:  either drop date or off air date
                If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                    sBaseEndDate = sDropDate
                Else
                    sBaseEndDate = sOffAir
                End If
                If Not UpdateNCRRst.EOF Then
                    SQLQuery = "Update att Set AttNCR =" & "'Y'" & " where attcode = " & llncragreements(lLoop)
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DelqRpt-cmdReport_Click"
                        cnn.RollbackTrans
                        Exit Sub
                    End If
                    cnn.CommitTrans
                    
                    'find any agreements in the futures and set the NCR flags in those too
                    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate FROM att"
                    SQLQuery = SQLQuery + " WHERE (attShfCode = " & UpdateNCRRst!attshfcode & " AND attVefCode = " & UpdateNCRRst!attvefCode & ")"
                    Set FutureNCRRst = gSQLSelectCall(SQLQuery)
                    While Not FutureNCRRst.EOF
                    
                        sOffAir = Format$(FutureNCRRst!attOffAir, "mm/dd/yyyy")
                        sDropDate = Format$(FutureNCRRst!attDropDate, "mm/dd/yyyy")
                        'determine the earliest of 2 dates:  either drop date or off air date
                        If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                            sAttEndDate = sDropDate
                        Else
                            sAttEndDate = sOffAir
                        End If
                        
                        If gDateValue(gAdjYear(sAttEndDate)) > gDateValue(gAdjYear(sBaseEndDate)) Then
                            'set the NCR flags for agreements in the future
                            SQLQuery = "Update att Set AttNCR =" & "'" & UpdateNCRRst!attNCR & "', attFormerNCR = " & "'" & UpdateNCRRst!attFormerNCR & "' where attcode = " & FutureNCRRst!attCode
                            cnn.BeginTrans
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "DelqRpt-cmdReport_Click"
                                cnn.RollbackTrans
                                Exit Sub
                            End If
                            cnn.CommitTrans
                        End If
                        FutureNCRRst.MoveNext
                        Wend
                    
                End If
            Next lLoop
            If UBound(llncragreements) > 0 Then
                UpdateNCRRst.Close
                FutureNCRRst.Close
            End If

        End If
        DoEvents            '9-17-09 required for some options because the file gets cleared before crystal can spit out the report
        ' Delete the info we stored in the GRF table
        SQLQuery = "DELETE FROM GRF_Generic_Report"
        'SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Format$(sGenTime, sgSQLTimeForm) & "')"
        SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        'cnn.Execute SQLQuery            ', rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
 
    End If
    If Not iGotAnyData Then             'if not set, nothing was found to output
        gMsgBox "No Data Exists for Requested Period"
    End If
    Erase llncragreements
    DelinRst.Close
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "DelqRpt-cmdReport"
End Sub

Private Sub InsertIntoGRF(ByVal lArrayIdx As Long, ByVal sGenDate As String, ByVal sGenTime As String, ByVal iPrevStnCode As Integer, ByVal iPrevVehCode As Integer, lPrevAtfCode As Long, ByVal iTtlUnqDates As Integer, ByVal iTtlDatesInRange As Integer, sPrevNCR As String, sPrevFormerNCR As String, iPrevPostingStatus As Integer)
    '10/30/09 chnage iArrayIdx & idx from integer to long
    Dim idx As Integer
    Dim ldx As Long
    Dim ContactRst As ADODB.Recordset
    Dim AERst As ADODB.Recordset
    Dim llArttCode1 As Long
    Dim slAEName As String
    Dim llArttCode As Long
    Dim ilMktRepUstCode As Integer     'market rep (affiliate a/e)
    Dim blDoMktRep As Boolean          'true to process mkt rep, vs process service rep
    Dim ilConsecutiveWks As Integer     '4-25-17 # of consecutive weeks found that are still delinquent
    Dim blFirstWk As Boolean
    Dim llWkOf As Long
    Dim llPrevWkOf As Long
    Dim ilWeek As Integer
    ' Note: If you change the value of cMaxDates you must add to or delete from the list of
    '       dates in the array of start dates - tmDlq(ldx).sStartDate(?)
    If lPrevAtfCode = 21 Then
    blFirstWk = blFirstWk
    End If
    '4-25-17 if Critically overdue report, verify that there are "X" consecutive weeks unposted ("X" is the value defined in site
    If imRptIndex = NCR_RPT Then
        ilConsecutiveWks = 0
        blFirstWk = True
        'tmdlq array contains the dates of the delinquent affidavits
        'loop thru the array of dates, each record contains 6 dates.  Need to loop thru each of the six dates within
        'each record to see if there are consecutive # of weeks delinquent across all the dates.  There may be
        'holes between a group of dates.  Once the number of consecutive weeks has been met, can exit this code and continue
        'writing them for prepass.  If the consecutive # of weeks not met, do not create prepass.
        For ldx = LBound(tmDlq) To UBound(tmDlq)
            For ilWeek = 0 To cMaxDates
                'each array contains 6 dates, end is either "0/0/0000" or "1/1/1970"
                If Trim$(tmDlq(ldx).sStartDate(ilWeek)) = "0/0/0000" Or Trim$(tmDlq(ldx).sStartDate(ilWeek)) = "1/1/1970" Then      'reached end of delinquent dates
                    If ilConsecutiveWks < imConsecutiveWksNCR Then          'nothing critically overdue
                        Exit For
                    End If
                Else
                    If blFirstWk Then
                        blFirstWk = False
                        llWkOf = gDateValue(Trim$(tmDlq(ldx).sStartDate(ilWeek)))
                        llPrevWkOf = llWkOf
                        ilConsecutiveWks = ilConsecutiveWks + 1
                    Else
                        llWkOf = gDateValue(Trim$(tmDlq(ldx).sStartDate(ilWeek)))
                        If llWkOf = llPrevWkOf + 7 Then
                            ilConsecutiveWks = ilConsecutiveWks + 1
                            llPrevWkOf = llWkOf
                           
                        Else
                            'start over with consecutive week count, break in the delinquent weeks
                            ilConsecutiveWks = 1
                            llPrevWkOf = llWkOf
                        End If
                    End If
                End If
                If ilConsecutiveWks >= imConsecutiveWksNCR Then
                    Exit For
                End If
            Next ilWeek
            If ilConsecutiveWks >= imConsecutiveWksNCR Then
                Exit For
            End If
        Next ldx
        If ilConsecutiveWks < imConsecutiveWksNCR Then      '
            Exit Sub
        End If
    End If
    
    
    For ldx = 0 To lArrayIdx Step 1
        'If tmDlq(ldx).sStartDate(1) = "0/0/0000  " Then
        If tmDlq(ldx).sStartDate(0) = "0/0/0000  " Then
            Exit For
        End If
        
        'find the applicable affidavit contact
        '9-8-11 change back to using arttaffcontact flag
        '2-12-13 reinstate
        SQLQuery = "SELECT * From artt INNER JOIN shtt on arttshttCode = shttcode  where  arttAffContact = '1' and arttshttcode = " & iPrevStnCode
        '3-8-11 need the affidavit contact, its no longer the same flag
        'SQLQuery = "SELECT * From artt INNER JOIN shtt on arttshttCode = shttcode  where  arttWebEMail = 'Y' and arttshttcode = " & iPrevStnCode
        
        'get the affidavit contact
        Set ContactRst = gSQLSelectCall(SQLQuery)
        llArttCode1 = 0
        If Not ContactRst.EOF Then
            While Not ContactRst.EOF
                '8-31-11 select a personnel that is affwebemail defined with a name in the record
                If Trim$(ContactRst!arttLastName) <> "" Or Trim$(ContactRst!arttFirstName) <> "" Then
                    llArttCode1 = ContactRst!arttCode
                End If
                ContactRst.MoveNext
            Wend
        End If
        
        
'        blDoMktRep = True
'        If imRptIndex <> AFFSMISSINGWKS_RPT Then
'            'if NCR or Overdue , determine if service or mkt rep.  Do Mkt rep by default
'            If imSortSelected = SORT_SVCREP Then            '5-20-19
'                blDoMktRep = False
'            End If
'        End If

'        If blDoMktRep Then
        If imSortSelected = SORT_MKTREP Then
            '1-11-13 Affiliate A/E is now Market Rep for sorting
            SQLQuery = "Select attMktRepUstCode from att where attcode = " & lPrevAtfCode
            'llArttCode1 = 0
            ilMktRepUstCode = 0
            Set ContactRst = gSQLSelectCall(SQLQuery)
            If Not ContactRst.EOF Then
               ' llArttCode1 = ContactRst!attMktRepUstCode
               ilMktRepUstCode = ContactRst!attMktRepUstCode
            End If
            
            'If llArttCode1 > 0 Then         'agreement has a market rep
            If ilMktRepUstCode = 0 Then      'agreement has no rep code, default to station record
                SQLQuery = "Select shttMktRepUstCode from SHTT where shttcode = " & iPrevStnCode
                Set ContactRst = gSQLSelectCall(SQLQuery)
                If Not ContactRst.EOF Then
                    'slAEName = Trim$(ContactRst!ustfirstname) + " " + Trim$(ContactRst!ustLastName)
                    ilMktRepUstCode = ContactRst!shttMktRepUstCode
                End If
            End If
        ElseIf imSortSelected = SORT_SVCREP Then
            SQLQuery = "Select attServRepUstCode from att where attcode = " & lPrevAtfCode
            'llArttCode1 = 0
            ilMktRepUstCode = 0
            Set ContactRst = gSQLSelectCall(SQLQuery)
            If Not ContactRst.EOF Then
               ' llArttCode1 = ContactRst!attMktRepUstCode
               ilMktRepUstCode = ContactRst!attServRepUstCode
            End If
            
            'If llArttCode1 > 0 Then         'agreement has a market rep
            If ilMktRepUstCode = 0 Then      'agreement has no rep code, default to station record
                SQLQuery = "Select shttServRepUstCode from SHTT where shttcode = " & iPrevStnCode
                Set ContactRst = gSQLSelectCall(SQLQuery)
                If Not ContactRst.EOF Then
                    'slAEName = Trim$(ContactRst!ustfirstname) + " " + Trim$(ContactRst!ustLastName)
                    ilMktRepUstCode = ContactRst!shttServRepUstCode
                End If
            End If
        End If
        
'        slAEName = ""
        
'        If optSortby(4).Value = True Then         'sort by AE
'            'get the associated AE name to place into prepass (affiliate system does not
'            'handle alias files in crystal to set the database locations
'            'first get the agreement to get the AE pointer
'            SQLQuery = "Select attarttCode from att where attcode = " & lPrevAtfCode
'            llArttCode = 0
'            Set AERst = gSQLSelectCall(SQLQuery)
'               If Not AERst.EOF Then
'                llArttCode = AERst!attArttCode
'            End If
'
'            'SQLQuery = "Select arttFirstName, arttLastName, attarttcode, arttcode  from att LEFT OUTER JOIN artt on attarttCode = arttCode where( attshfcode = " & iPrevStnCode & " and attvefCode = " & iPrevVehCode & ") "
'            SQLQuery = "Select arttFirstName, arttLastName, attarttcode, arttcode from artt, att where attarttcode = arttcode and arttcode = " & llArttCode
'            Set AERst = gSQLSelectCall(SQLQuery)
'               If Not AERst.EOF Then
'                slAEName = Trim$(AERst!arttFirstName) & " " & Trim$(AERst!arttLastName)
'            End If
'        End If

            'grfSTartDate(1-6) - overdue cps
            'grfCode2 - Station Code (iPrevStnCode)
            'grfsofCode - VEhicle code (iPrevVehCode)
            'grfPer1 - Total dates outstanding (iTtlUnqDates)
            'grfPer2 - Total dates outstanding in range (iTtlDatesInRange)
            'grfPer3 - Print status (tmdlq(ldx).lprintstatus))
            'grfPer4 - Mkt Rep code -ilMktRepUstCode
            'grfPer5 - Contact code (llArttCode1)
            'grfGenDesc - AE Name (slAEName)
            'grfBktType - NCR flag (iPrevNCR)
            'grfgrfDateType - Previous NCR notation (sPrevFormerNCR)
            'grfPer1Genl(1-7) Delivery methods (CBS, ClearChannel, Cumulus, Manual, Marketron, NetWeb, Univision)
            'grfGenDate - Generation date for filtering (sGenDate)
            'grfGenTime - Generation time for filtering (sGenTime)
            SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
            SQLQuery = SQLQuery & " (grfDate1, grfDate2, grfDate3, grfDate4, "
            'SQLQuery = SQLQuery & "  grfDate5, grfDate6, grfDate7, "
            SQLQuery = SQLQuery & "  grfDate5, grfDate6,  "
            SQLQuery = SQLQuery & "  grfCode2, grfSofCode, grfPer1, grfPer2, grfPer3, grfPer4, grfPer5, "
            SQLQuery = SQLQuery & "  grfGenDesc, grfBktType, grfDateType, "
            SQLQuery = SQLQuery & "  grfPer1Genl, grfPer2Genl, grfPer3Genl, grfPer4Genl, grfPer5Genl, grfPer6Genl, grfPer7Genl,  "
            SQLQuery = SQLQuery & "  grfPer8Genl, grfPer9Genl, grfPer10Genl, grfPer11Genl, grfPer12Genl, grfPer13Genl, "
            
            SQLQuery = SQLQuery & "  grfGendate, grfGenTime) "
            'SQLQuery = SQLQuery & " VALUES ('" & Format$(tmDlq(ldx).sStartDate(1), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(2), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & " VALUES ('" & Format$(tmDlq(ldx).sStartDate(0), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(1), sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$(tmDlq(ldx).sStartDate(3), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(4), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(tmDlq(ldx).sStartDate(2), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(3), sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$(tmDlq(ldx).sStartDate(5), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(6), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(tmDlq(ldx).sStartDate(4), sgSQLDateForm) & "', " & "'" & Format$(tmDlq(ldx).sStartDate(5), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & iPrevStnCode & "', " & "'" & iPrevVehCode & "', " & "'" & iTtlUnqDates & "', " & "'" & iTtlDatesInRange & "', " & "'" & tmDlq(ldx).lPrintStatus & "', " & "'" & ilMktRepUstCode & "', " & "'" & llArttCode1 & "', "
            SQLQuery = SQLQuery & "'" & slAEName & "', " & "'" & sPrevNCR & "', " & "'" & sPrevFormerNCR & "', "
            SQLQuery = SQLQuery & "'" & tmPrevDeliveryWeek.iCBS & "', " & "'" & tmPrevDeliveryWeek.iClearCh & "', " & "'" & tmPrevDeliveryWeek.iCumulus & "', "
            SQLQuery = SQLQuery & "'" & tmPrevDeliveryWeek.iManual & "', " & "'" & tmPrevDeliveryWeek.iMarketron & "', " & "'" & tmPrevDeliveryWeek.iNetWeb & "', " & "'" & tmPrevDeliveryWeek.iUni & "', "
            'SQLQuery = SQLQuery & "'" & tmDlq(ldx).iPostingStatus(1) & "', " & "'" & tmDlq(ldx).iPostingStatus(2) & "', " & "'" & tmDlq(ldx).iPostingStatus(3) & "', " & "'" & tmDlq(ldx).iPostingStatus(4) & "', " & "'" & tmDlq(ldx).iPostingStatus(5) & "', " & "'" & tmDlq(ldx).iPostingStatus(6) & "', "
            SQLQuery = SQLQuery & "'" & tmDlq(ldx).iPostingStatus(0) & "', " & "'" & tmDlq(ldx).iPostingStatus(1) & "', " & "'" & tmDlq(ldx).iPostingStatus(2) & "', " & "'" & tmDlq(ldx).iPostingStatus(3) & "', " & "'" & tmDlq(ldx).iPostingStatus(4) & "', " & "'" & tmDlq(ldx).iPostingStatus(5) & "', "
            
            SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"   '", "
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "DelqRpt-cmdReport_Click"
                cnn.RollbackTrans
                Exit Sub
            End If
            cnn.CommitTrans
        Next
        
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPCount-InsertIntoGRF"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmDelqRpt
End Sub

Private Sub CalEffActionDate_GotFocus()
    gCtrlGotFocus CalEffActionDate
    CalEffActionDate.ZOrder (vbBringToFront)
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
Private Sub cmdStationListFile_Click()
    Dim slCurDir As String
    slCurDir = CurDir
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    ' Import from the Selected File
    gSelectiveStationsFromImport lbcVehAff, chkListBox, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub


Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.2
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmDelqRpt
    lacDateDisclaimer.FontSize = 8
    gCenterForm frmDelqRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim slSunday As String

    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    imRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    imNoWksNCR = 0
    If Not rst.EOF Then
        imNoDaysDelq = rst!siteOMNoWeeks         'determine from date by # weeks considered overdue
        imNoWksNCR = rst!siteNCRWks         '6-30-09 # weeks behind considered non-compliant
        imConsecutiveWksNCR = rst!siteNCRWks        '- rst!siteOMNoWeeks
        
    End If
   
    slDate = Format$(gNow(), sgShowDateForm)
    Do While Weekday(slDate) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    
    'Affiliate Missing weeks can use upto today (Sunday date)
    slSunday = slDate
    Do While Weekday(slSunday) <> vbSunday
        slSunday = DateAdd("d", -1, slSunday)
    Loop
    cbcSort.Clear
    cbcSort.AddItem "Vehicle"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_VEHICLE
    cbcSort.AddItem "Station"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_STATION
    cbcSort.AddItem "DMA Market Name"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_DMA_MKTNAME
    cbcSort.AddItem "DMA Market Rank"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_DMA_MKTRANK
    cbcSort.AddItem "Market Rep"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_MKTREP
    cbcSort.AddItem "Producer"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_PRODUCER
    cbcSort.AddItem "Service Rep"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_SVCREP
    If imRptIndex = OVERDUE_RPT Then
        cbcSort.AddItem "Aud P12+"
        cbcSort.ItemData(cbcSort.NewIndex) = SORT_AUDP12
    End If
    cbcSort.AddItem "Owner"
    cbcSort.ItemData(cbcSort.NewIndex) = SORT_OWNER
    For iLoop = 0 To cbcSort.ListCount - 1                 'find the default option in sorted list and set default sort
        If cbcSort.ItemData(iLoop) = SORT_VEHICLE Then
            cbcSort.ListIndex = iLoop
            Exit For
        End If
    Next iLoop
    
    If imRptIndex = NCR_RPT Then
        frmDelqRpt.Caption = "Critically Overdue Report - " & sgClientName

        '4-24-17 default to 52 weeeks back (like Aff mgmt)
        CalToDate.Text = DateAdd("d", -((imNoDaysDelq * 7) + 1), slDate)
        CalFromDate.Text = DateAdd("d", -((51 * 7) - 1), CalToDate.Text)
        'rbcSuppressNotice(1).Value = True
        ckcInclComments.Value = vbChecked   'default include comments on
        smDefaultStartDate = CalFromDate.Text
        smDefaultEndDate = CalToDate.Text
        lmDefaultStartDate = gDateValue(smDefaultStartDate)
        lmDefaultEndDate = gDateValue(smDefaultEndDate)
        OptWks(0).Enabled = False
        lacPostType.Visible = False
        frcPostType.Visible = False
        rbcPostType(2).Value = True

    Else
        If imRptIndex = AFFSMISSINGWKS_RPT Then
            frmDelqRpt.Caption = "Affiliates Missing Weeks -" & sgClientName
'            optSortby(1).Value = True               'default to sort by stations
'            For iLoop = 0 To cbcSort.ListCount - 1                 'find the default option in sorted list and set default sort
'                If cbcSort.ItemData(iLoop) = SORT_STATION Then
'                    cbcSort.ListIndex = iLoop
'                    Exit For
'                End If
'            Next iLoop
'            rbcSuppressNotice(1).Value = True       'ignore suppress notices
            optVehAff(0).Value = True               'select by stations
'            rbcExpired(0).Value = True              'show expired
'            Frame6.Visible = False                  '5-20-19 replaced by cbcsort
'            lacNewPage.Visible = False
'            Frame9.Visible = False                  'new page Y/N
'            lacExpired.Visible = False
'            frcExpired.Visible = False              'include expired agreements
'            lacSuppressNotice.Visible = False
'            Frame4.Visible = False               'suppress notice
            lacAction.Visible = False
            ckcInclComments.Visible = False
            lacEffecCommentDate.Visible = False
            CalEffActionDate.Visible = False
            lacPassword.Visible = False
            frcPassword.Visible = False
            ckcUpdateNCR.Visible = False
            lacOldestNCRDate.Visible = False
            CalOldestNCRDate.Visible = False
            lacDateDisclaimer.Visible = False
'            Frame3.Visible = False
'            chkListBox.Visible = False
'            lbcVehAff.Visible = False
            'default end date to todays date
            CalToDate.Text = slSunday
'            ckcAllDelivery.Visible = False
'            lbcDelivery.Visible = False
            lacPostType.Visible = False
            rbcPostType(0).Visible = False
            rbcPostType(1).Visible = False
            rbcPostType(2).Visible = False
            
        Else
            frmDelqRpt.Caption = "Overdue Affidavits Report - " & sgClientName
            CalToDate.Text = DateAdd("d", -((imNoDaysDelq * 7) + 1), slDate)
            CalEffActionDate.SetEnabled (False)

'            optSortby(7).Visible = True                     'Audience sort for Overdue report only
        End If
        CalFromDate.Text = ""
        'default TO date based on the current Monday (from TODAY), minus the # weeks considered overdue
        'CalToDate.Text = DateAdd("d", -((imNoDaysDelq * 7) + 1), slDate)
        ckcUpdateNCR.Visible = False
        ckcUpdateNCR.Value = vbUnchecked
        lacOldestNCRDate.Visible = False
        CalOldestNCRDate.Visible = False
        lacDateDisclaimer.Visible = False
    End If
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
       lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
       lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop

    mPopDelivery
    imChkListBoxIgnore = False
    
    
'    slDate = Format$(gNow(), sgShowDateForm)
'    Do While Weekday(slDate, vbSunday) <> vbMonday
'        slDate = DateAdd("d", -1, slDate)
'    Loop
'    CalFromDate.text = ""
'    CalToDate.text = DateAdd("d", -((imNoDaysDelq * 7) + 1), slDate)
'    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
'            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
'            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
'    Next iLoop
    gPopExportTypes cboFileType     '3-15-04 populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmDlq
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmDelqRpt = Nothing
End Sub


Private Sub lbcDelivery_Click()
    If imChkDeliveryBoxIgnore Then
        Exit Sub
    End If
    If ckcAllDelivery.Value = 1 Then
        imChkDeliveryBoxIgnore = True
        'chkListBox.Value = False
        ckcAllDelivery.Value = 0    'chged from false to 0 10-22-99
        imChkDeliveryBoxIgnore = False
    End If

End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub
'5-20-19 replace with cbcsort
'Private Sub optSortby_Click(Index As Integer)
'    'default vehicle, affiliate a/e and producer sorts and Service rep to skip to new page
'    If optSortby(0).Value = True Or optSortby(4).Value = True Or optSortby(5).Value = True Or optSortby(6).Value = True Then
'        optPage(0).Value = True
'    Else
'        'default station, market and market name skip to new page off
'        optPage(1).Value = True
'    End If
'End Sub

Private Sub optVehAff_Click(Index As Integer)
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    Screen.MousePointer = vbHourglass
    If optVehAff(0).Value = True Then
        'TTP 9943
        cmdStationListFile.Visible = False
        chkListBox.Caption = "All Vehicles"
        chkListBox.Value = 0    'False
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then    '4-27-09 chged to test only OLA
                lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
            'End If
        Next iLoop
    Else
        'TTP 9943
        cmdStationListFile.Visible = True
        chkListBox.Caption = "All Stations"
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcVehAff.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            End If
        Next iLoop
     End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub OptWks_Click(Index As Integer)
Dim slDate As String
Dim slSunday As String
    If OptWks(Index).Value Then
        If Index = 0 Then
            CalFromDate.Text = ""
            CalToDate.Text = ""
            CalFromDate.SetEnabled (False)
            CalToDate.SetEnabled (False)
            labFrom.Enabled = False
            LabTo.Enabled = False
        Else
            slDate = Format$(gNow(), sgShowDateForm)
            Do While Weekday(slDate, vbSunday) <> vbMonday
                slDate = DateAdd("d", -1, slDate)
            Loop
            
            slSunday = slDate
            Do While Weekday(slSunday) <> vbSunday
                slSunday = DateAdd("d", -1, slSunday)
            Loop
            If imRptIndex = AFFSMISSINGWKS_RPT Then
                CalFromDate.Text = ""
                CalToDate.Text = slSunday
            Else
                CalFromDate.Text = ""
                CalToDate.Text = DateAdd("d", -((imNoDaysDelq * 7) + 1), slDate)
            End If
            CalFromDate.SetEnabled (True)
            CalToDate.SetEnabled (True)
            labFrom.Enabled = True
            LabTo.Enabled = True
        End If
    End If
End Sub

Private Sub CalFromDate_Change()
Dim ilLen As Integer
Dim slDate As String
Dim llFromDate As Long
Dim llToDate As Long

    If lmDefaultStartDate > 0 Then
        ilLen = Len(CalFromDate.Text)
        If ilLen >= 3 Then              'date entered may be x/x (no year)
            slDate = CalFromDate.Text          'retrieve jan thru dec year
            llFromDate = gDateValue(slDate)
            slDate = CalToDate.Text
            llToDate = gDateValue(slDate)
            If llFromDate <> lmDefaultStartDate Or llToDate <> lmDefaultEndDate Then
                ckcUpdateNCR.Value = vbUnchecked
                ckcUpdateNCR.Enabled = False
            Else
                ckcUpdateNCR.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub CalFromDate_GotFocus()
    gCtrlGotFocus CalFromDate
    CalFromDate.ZOrder (vbBringToFront)
End Sub

Private Sub CalToDate_Change()
Dim ilLen As Integer
Dim slDate As String
Dim llFromDate As Long
Dim llToDate As Long

    If lmDefaultEndDate > 0 Then
        ilLen = Len(CalToDate.Text)
        If ilLen >= 3 Then                    'date entered may be x/x (no year)
            slDate = CalToDate.Text          'retrieve jan thru dec year
            llToDate = gDateValue(slDate)
            slDate = CalFromDate.Text
            llFromDate = gDateValue(slDate)
            
            If llFromDate <> lmDefaultStartDate Or llToDate <> lmDefaultEndDate Then
                ckcUpdateNCR.Value = vbUnchecked
                ckcUpdateNCR.Enabled = False
            Else
                ckcUpdateNCR.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub CalToDate_GotFocus()
    gCtrlGotFocus CalToDate
    CalToDate.ZOrder (vbBringToFront)
End Sub
Private Sub mPopDelivery()
    'order of entries same as on the Affiliate Delivery tab
    lbcDelivery.Clear
    lbcDelivery.AddItem "Manual"
    lbcDelivery.AddItem "Counterpoint Affidavit System"   ' JD TTP:10583 12-22-22
    lbcDelivery.AddItem "Univision"
    '9184
    lbcDelivery.AddItem "Stratus"
    lbcDelivery.AddItem "Marketron"
    lbcDelivery.AddItem "CBS"
    lbcDelivery.AddItem "Clear Channel"
    ckcAllDelivery.Value = vbChecked            'force to select all delivery methods
End Sub
'           mDeliverySelected - determine which delivery options user has selected
'
Private Sub mDeliverySelected()
Dim ilLoop As Integer
    tmDeliverSelected.iManual = False
    tmDeliverSelected.iNetWeb = False
    tmDeliverSelected.iCumulus = False
    tmDeliverSelected.iUni = False
    tmDeliverSelected.iMarketron = False
    tmDeliverSelected.iCBS = False
    tmDeliverSelected.iClearCh = False

    For ilLoop = 0 To lbcDelivery.ListCount - 1
        If lbcDelivery.Selected(ilLoop) Then
            If lbcDelivery.List(ilLoop) = "Manual" Then
                tmDeliverSelected.iManual = True
            ElseIf lbcDelivery.List(ilLoop) = "Counterpoint Affidavit System" Then  ' JD TTP:10583 12-22-22
                tmDeliverSelected.iNetWeb = True
            '9184
            ElseIf lbcDelivery.List(ilLoop) = "Stratus" Then
                tmDeliverSelected.iCumulus = True
            ElseIf lbcDelivery.List(ilLoop) = "Univision" Then
                tmDeliverSelected.iUni = True
            ElseIf lbcDelivery.List(ilLoop) = "Marketron" Then
                tmDeliverSelected.iMarketron = True
            ElseIf lbcDelivery.List(ilLoop) = "CBS" Then
                tmDeliverSelected.iCBS = True
            ElseIf lbcDelivery.List(ilLoop) = "Clear Channel" Then
                tmDeliverSelected.iClearCh = True
            End If
        End If
    Next ilLoop
    Exit Sub
End Sub

