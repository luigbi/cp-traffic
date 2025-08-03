VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form DashboardVw 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   20790
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9870
   ScaleWidth      =   20790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSetup 
      Caption         =   "User Dashboard Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   600
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ListView lvcColumns 
         Height          =   3375
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5953
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox chkExcelHeader 
         Caption         =   "Excel Generate Header"
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         ToolTipText     =   "Generate Headers when opening results in Excel"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmcResetColumns 
         Caption         =   "Reset Columns"
         Height          =   375
         Left            =   3480
         TabIndex        =   38
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox CboFontSize 
         Height          =   315
         ItemData        =   "DashboardVw.frx":0000
         Left            =   4560
         List            =   "DashboardVw.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox ckcShowOnStartup 
         Caption         =   "Show Dashboard at Startup"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         ToolTipText     =   "When checked, dashboard will launch at startup for you."
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmcSetupDone 
         Caption         =   "Done"
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblFontSize 
         Caption         =   "Font Size:"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblColumnSetupHeader 
         Caption         =   "Columns"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FDFFD7&
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   3970
   End
   Begin V81Traffic.CSI_Calendar CSI_Calendar1 
      Height          =   315
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   556
      Text            =   "04/15/2024"
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
   Begin VB.ListBox lstFilterSearch 
      Height          =   2400
      Left            =   2520
      TabIndex        =   6
      Top             =   1155
      Visible         =   0   'False
      Width           =   3980
   End
   Begin VB.ComboBox cboItems 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   4255
   End
   Begin VB.ComboBox cboFilterType 
      BackColor       =   &H00FDFFD7&
      Height          =   315
      ItemData        =   "DashboardVw.frx":0004
      Left            =   2530
      List            =   "DashboardVw.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCntr 
      Height          =   7335
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   20775
      _ExtentX        =   36645
      _ExtentY        =   12938
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pbcFooter 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   20715
      TabIndex        =   16
      Top             =   9240
      Width           =   20775
      Begin VB.PictureBox pbcButtons 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   7215
         TabIndex        =   32
         Top             =   120
         Width           =   7215
         Begin VB.CommandButton cmcChgCntr 
            Caption         =   "Change Order"
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
            Height          =   375
            Left            =   1080
            TabIndex        =   37
            Top             =   0
            Width           =   1695
         End
         Begin VB.CommandButton cmcViewCntr 
            Caption         =   "View Order"
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
            Height          =   375
            Left            =   2880
            TabIndex        =   36
            Top             =   0
            Width           =   1695
         End
         Begin VB.CommandButton cmcSchedule 
            Caption         =   "Schedule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   35
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmcClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmcExport 
            Caption         =   "Export..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            TabIndex        =   33
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmcSetup 
         Height          =   495
         Left            =   0
         Picture         =   "DashboardVw.frx":0063
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "User Dashboard Settings"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmcCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13800
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   11160
         Top             =   120
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   13320
         TabIndex        =   22
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblMultiVersionTotalWarn 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross may include totals of multiple proposal versions"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9480
         TabIndex        =   41
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblItemCount 
         Alignment       =   1  'Right Justify
         Caption         =   "Showing ### Items"
         Height          =   255
         Left            =   9480
         TabIndex        =   21
         Top             =   165
         Width           =   3735
      End
   End
   Begin VB.PictureBox pbcHeader 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   20715
      TabIndex        =   13
      Top             =   0
      Width           =   20775
      Begin VB.Frame Frame2 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2400
         TabIndex        =   15
         Top             =   60
         Width           =   18135
         Begin VB.CommandButton cmcResetFilter 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3840
            Picture         =   "DashboardVw.frx":0CA5
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Reset Filters"
            Top             =   165
            Width           =   495
         End
         Begin VB.CommandButton cmcRemoveFilter 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3360
            Picture         =   "DashboardVw.frx":18E7
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Delete Selected Filter"
            Top             =   165
            Width           =   495
         End
         Begin VB.CommandButton cmcUpdateFilter 
            Enabled         =   0   'False
            Height          =   495
            Left            =   2760
            Picture         =   "DashboardVw.frx":2529
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Update Selected Filter"
            Top             =   165
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmcAddFilter 
            Enabled         =   0   'False
            Height          =   495
            Left            =   2760
            Picture         =   "DashboardVw.frx":316B
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Add Filter"
            Top             =   165
            Width           =   615
         End
         Begin RichTextLib.RichTextBox txtCriteria 
            Height          =   1335
            Left            =   4440
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   150
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   2355
            _Version        =   393217
            BackColor       =   16646103
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            MousePointer    =   1
            TextRTF         =   $"DashboardVw.frx":3DAD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Selected Item:"
            Height          =   255
            Left            =   4440
            TabIndex        =   19
            Top             =   165
            Width           =   2415
         End
         Begin VB.Label lblSelectedFilter 
            Caption         =   "[selected item]"
            Height          =   255
            Left            =   4440
            TabIndex        =   18
            Top             =   405
            Width           =   5655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   60
         Width           =   2175
         Begin VB.CheckBox ckcCntrType 
            Caption         =   "Order"
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   11
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox ckcCntrType 
            Caption         =   "Proposal"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox ckcCntrType 
            Caption         =   "Digital"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox ckcCntrType 
            Caption         =   "NTR"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox ckcCntrType 
            Caption         =   "Air Time"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuFilterBy 
         Caption         =   "Filter by"
      End
      Begin VB.Menu mnuFilterBy2 
         Caption         =   "Filter by"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnuExportCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Open in Excel"
      End
      Begin VB.Menu mnuCSV 
         Caption         =   "Export to CSV"
      End
   End
End
Attribute VB_Name = "DashboardVw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bmFirstActivate As Boolean
Dim tmChfDetailList() As CHFDETAILLIST
Dim smLastStartDate As String
Dim smLastEndDate As String
Dim imSortDir As Integer 'Column Sort Direction
Dim imLastSortCol As Integer 'Column Sorted
Dim imLastFilledCol As Integer 'Last special sort column filled
Dim imSlpOffice As Integer
Dim imKeyDelay As Integer
Dim tmAppliedFilters() As APPLIEDFILTER
Dim lmItemCount As Long
Dim dmTotalAmount As Double
Dim imCancelLoading As Integer
Dim bmBuildingFilteredResults As Boolean
Dim imColumnWidths() As Integer
Dim smDefaultFilters As String
Dim tmDefaultFilters() As APPLIEDFILTER
Dim imLastSlpGroup As Integer
Dim smCurrentTitle As String
Dim imCurrentGroup As Integer
Dim imLastGetSPOffice As Integer
Dim imLastGetSPOfficeSlf As Integer

'Sales office
Dim tmSofList() As SOF
Dim tmSof As SOF
Dim hmSof As Integer            'Sale Office file handle
Dim imSofRecLen As Integer      'SOF record length

'Last looked up for performance
Dim imLastAdv As Integer
Dim smLastAdvName As String
Dim imLastAgy As Integer
Dim smLastAgyName As String
Dim imLastSlp As Integer
Dim imLastSlpName As String
Dim imLastSlpOffice As Integer
Dim smLastSlpOfficeName As String

'Contract grid columns
Const C_CHFCODEINDEX = 0
Const C_SORTINDEX = 1
Const C_CNTRUPDATEDATEINDEX = 2
Const C_CNTRNOINDEX = 3
Const C_CNTRTYPEINDEX = 4
Const C_LINETYPEINDEX = 5
Const C_AGYNAMEINDEX = 6
Const C_ADVNAMEINDEX = 7
Const C_PRODUCTINDEX = 8
Const C_STARTDATEINDEX = 9
Const C_ENDDATEINDEX = 10
Const C_GROSSINDEX = 11
Const C_SALESOFFICEINDEX = 12
Const C_SALEPERSONINDEX = 13
Const C_CNTRSTATUSINDEX = 14
Const C_DIGITALDLVYINDEX = 15
Const C_CNTRSCHEDULESTATUSINDEX = 16

'used to narrow the filters down to "Displayed" items
Dim smAdvertiserList() As String
Dim smAgencyList() As String
Dim smSalesOfficeList() As String
Dim smSalespersonList() As String
Dim smContractStatusList() As String
Dim smContractTypeList() As String
Dim smDeliveryStatusList() As String
Dim smScheduleStatusList() As String
    
Private Sub CboFontSize_Click()
    If bmFirstActivate Then Exit Sub
    mPopulate
End Sub

Private Sub cmcResetFilter_Click()
    txtCriteria.Text = smDefaultFilters
    
    ReDim tmAppliedFilters(0 To UBound(tmDefaultFilters)) As APPLIEDFILTER
    For ilTemp = 0 To UBound(tmDefaultFilters)
        tmAppliedFilters(ilTemp).lValue = tmDefaultFilters(ilTemp).lValue
        tmAppliedFilters(ilTemp).sType = tmDefaultFilters(ilTemp).sType
        tmAppliedFilters(ilTemp).sValue = tmDefaultFilters(ilTemp).sValue
    Next ilTemp
    
    cmcResetFilter.Enabled = False
    mPopulate
    
    If cboFilterType = "Sales Group" Then
        txtSearch.Text = imCurrentGroup
        txtSearch_LostFocus
    End If
End Sub

Private Sub cmcSetup_Click()
    'Show column picker fraColumns
    Frame1.Enabled = False
    Frame2.Enabled = False
    pbcFooter.Enabled = False
    Dim ilColumn As Integer
    mUpdateDefaultColumSizes
    fraSetup.Left = cmcSetup.Left + cmcSetup.Width + 100
    fraSetup.Top = pbcFooter.Top - fraSetup.height
    fraSetup.Visible = True
    lvcColumns.ListItems.Clear
    
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        lvcColumns.ListItems.Add , "C:" & ilColumn, grdCntr.TextMatrix(0, ilColumn)
    Next ilColumn
    For ilColumn = 1 To lvcColumns.ListItems.Count
        If grdCntr.ColWidth(ilColumn + C_CNTRUPDATEDATEINDEX - 1) > 0 Then
            lvcColumns.ListItems(ilColumn).Checked = True
        End If
    Next ilColumn
    lvcColumns.SetFocus
    cmcSetup.Enabled = False
End Sub

Private Sub cmcResetColumns_Click()
    Dim ilColumn As Integer
    Dim ilValue As Integer
    For ilColumn = 1 To lvcColumns.ListItems.Count
        If lvcColumns.ListItems(ilColumn).Checked = False Then lvcColumns.ListItems(ilColumn).Checked = True
    Next ilColumn
    
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        imColumnWidths(ilColumn) = -1
    Next ilColumn
    mSetDefaultColumnSizes 0, True
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        ilValue = imColumnWidths(ilColumn)
        grdCntr.ColWidth(ilColumn) = ilValue
    Next ilColumn
End Sub

Private Sub cmcSetupDone_Click()
    'Hide Column picker fraColumns
    cmcSetup.Enabled = True
    fraSetup.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    pbcFooter.Enabled = True
    mSaveDashSettings
End Sub

Private Sub Form_Activate()
    If Not bmFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    
    'Center Form
    Me.Width = Traffic.ScaleWidth - 220
    Me.height = Traffic.ScaleHeight - 400
    'Large screen
    If Me.Width > 24200 Then Me.Width = 24200
    If Me.height > 14000 Then Me.height = 14000
    'Small screen (1024x768)
    If Me.Width < 15140 Then Me.Width = 15140
    If Me.height < 9545 Then Me.height = 9545
    
    Form_Resize
    DashboardVw.Move (Traffic.Width - DashboardVw.Width) \ 2, (Traffic.ScaleHeight + 1700 - DashboardVw.height) \ 2

    'Set column widths
    mSetDefaultColumnSizes
    mLoadDashSettings

    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
    
    DashboardVw.Refresh
    DoEvents
    Me.KeyPreview = True
    
    'Populate grid grdCntr
    mPopulate
    
    mSelectFilterType "Advertiser"
    bmFirstActivate = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    bmFirstActivate = True
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim ilLoop As Integer
    
    ReDim imColumnWidths(0 To 17)
    
    'These hold the list of items from the loaded contracts
    ReDim tmAppliedFilters(0 To 0) As APPLIEDFILTER
    ReDim smAdvertiserList(0 To 0) As String
    ReDim smAgencyList(0 To 0) As String
    ReDim smSalesOfficeList(0 To 0) As String
    ReDim smSalespersonList(0 To 0) As String
    ReDim smContractStatusList(0 To 0) As String
    ReDim smContractTypeList(0 To 0) As String
    ReDim smDeliveryStatusList(0 To 0) As String
    ReDim smScheduleStatusList(0 To 0) As String
    
    'cboFontSize list
    CboFontSize.Clear
    CboFontSize.AddItem ("9")
    CboFontSize.AddItem ("8.25")
    CboFontSize.AddItem ("7")
    CboFontSize.AddItem ("6")
    CboFontSize.ListIndex = 1
    
    imKeyDelay = -1
    lblItemCount.Caption = "Starting Dashboard..."
    
    ckcShowOnStartup.Value = IIF(tgUrf(0).iDashOnStart = 1, vbChecked, vbUnchecked)
    
    'Default Filters
    txtCriteria.Text = ""
    'Mandatory Salesperson filter
    If tgUrf(0).iSlfCode <> 0 Then
        txtCriteria.Text = "[*SP:" & mGetSalespersonName(tgUrf(0).iSlfCode) & "] "
        mManageAppliedFilters "Add", "SP", CLng(tgUrf(0).iSlfCode), mGetSalespersonName(tgUrf(0).iSlfCode)
        ilInx = gBinarySearchSlf(tgUrf(0).iSlfCode)
        If ilInx >= 0 Then              '10-10-18
            smCurrentTitle = tgMSlf(ilInx).sJobTitle
            imCurrentGroup = tgUrf(0).iGroupNo
            'txtCriteria.Text = txtCriteria.Text & "[Group:" & imCurrentGroup & "] "
            'mManageAppliedFilters "Add", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup))
        End If
    End If
    
    'Starting Active Date
    txtCriteria.Text = txtCriteria.Text & "[*Active>:" & Format(gObtainStartStd(gNow()), "mm/dd/yy") & "] "
    mManageAppliedFilters "Add", "*Active>", 0, Format(gObtainStartStd(gNow()), "mm/dd/yy")
    
    'Setup Default filter memory
    smDefaultFilters = txtCriteria.Text
    ReDim tmDefaultFilters(0 To UBound(tmAppliedFilters)) As APPLIEDFILTER
    For ilTemp = 0 To UBound(tmAppliedFilters)
        tmDefaultFilters(ilTemp).lValue = tmAppliedFilters(ilTemp).lValue
        tmDefaultFilters(ilTemp).sType = tmAppliedFilters(ilTemp).sType
        tmDefaultFilters(ilTemp).sValue = tmAppliedFilters(ilTemp).sValue
    Next ilTemp
    
    lblSelectedFilter.Caption = ""
    grdCntr.Enabled = True
    mPopulateFilterType
    
    lstFilterSearch.Top = txtSearch.Top + txtSearch.height
    lstFilterSearch.Left = txtSearch.Left
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = 100
    Me.Refresh
    
    'Build array of Sales offices
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", DashboardVw
    imSofRecLen = Len(tmSof)

    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSofList(0 To ilTemp) As SOF
        tmSofList(ilTemp) = tmSof
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
End Sub

Private Sub Form_Resize()
    pbcHeader.Width = Me.ScaleWidth
    grdCntr.Width = Me.ScaleWidth
    pbcFooter.Width = Me.ScaleWidth
    grdCntr.height = Me.ScaleHeight - grdCntr.Top - pbcFooter.height - 160
    pbcFooter.Top = grdCntr.Top + grdCntr.height + 80
    
    cmcCancel.Left = pbcFooter.Width - cmcCancel.Width - 80
    lblItemCount.Left = pbcFooter.Width - lblItemCount.Width - 120
    lblMultiVersionTotalWarn.Left = lblItemCount.Left
    lblItemCount.Caption = "Showing " & Format(lmItemCount, "#,##0") & " Items / "
    
    Frame2.Width = Me.Width - Frame2.Left - 200
    txtCriteria.Width = Frame2.Width - txtCriteria.Left - 120
    
    pbcButtons.Left = 600 ' ((pbcFooter.Width - pbcButtons.Width) / 2) - lblItemCount.Width / 2
End Sub

Private Sub cboFilterType_Click()
    'Load Items Listbox or show CSI Calendar dropdown
    txtSearch.Text = ""
    lstFilterSearch.Clear
    cboItems.Clear
    CSI_Calendar1.Visible = False
    cboItems.Visible = False
    txtSearch.Enabled = True
    cboItems.Enabled = True
    
    Select Case cboFilterType.Text
        Case "Update Date (Beginning)"
            CSI_Calendar1.Visible = True
            cboItems.Visible = False
            txtSearch.Text = CSI_Calendar1.Text
            txtSearch_LostFocus
        
        Case "Update Date (Ending)"
            CSI_Calendar1.Visible = True
            cboItems.Visible = False
            txtSearch.Text = CSI_Calendar1.Text
            txtSearch_LostFocus
            
        Case "Active on or after"
            CSI_Calendar1.Visible = True
            cboItems.Visible = False
            txtSearch.Text = CSI_Calendar1.Text
            txtSearch_LostFocus
        
        Case "Active on or prior"
            CSI_Calendar1.Visible = True
            cboItems.Visible = False
            txtSearch.Text = CSI_Calendar1.Text
            txtSearch_LostFocus

        Case "Salesperson"
            'Check if Mandatory Salesperson filter already set.  if so, Don't allow Salesperson filtering
            CSI_Calendar1.Visible = False
            If InStr(1, txtCriteria.Text, "[*SP:") > 0 Then
                cboItems.Visible = False
                txtSearch.Enabled = True
            Else
                cboItems.Visible = True
                mPopulateSalespersonList cboItems
            End If
            
        Case "Sales Group"
            CSI_Calendar1.Visible = False
            cboItems.Visible = False
            txtSearch.Enabled = False
            txtSearch.Text = imCurrentGroup
            txtSearch_LostFocus
        
        Case "Sales Office"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateSalesOfficeList cboItems
            
        Case "Advertiser"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateAdvertiserList cboItems
        
        Case "Agency"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateAgencyList cboItems
            'cboItems.AddItem "N/A"
            
        Case "Schedule Status"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateScheduleStatusList cboItems
        
        Case "Contract Type"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateContractTypeList cboItems
        
        Case "Contract Status"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateContractStatusList cboItems
        
        Case "Digital Delivery Status"
            CSI_Calendar1.Visible = False
            cboItems.Visible = True
            mPopulateDigitalDlvyStatusList cboItems
            
    End Select
    
    If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
End Sub

Private Sub cboItems_Change()
    txtSearch.Text = cboItems.Text
    txtSearch_KeyUp vbKeyReturn, 0
    lstFilterSearch.Visible = False
    If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
End Sub

Private Sub cboItems_Click()
    txtSearch.Text = cboItems.Text
    'txtSearch_KeyUp vbKeyReturn, 0
    lstFilterSearch.Visible = False

    If mCheckCanAdd = True Then
        cmcAddFilter.Enabled = True
    Else
        cmcAddFilter.Enabled = False
    End If

    cmcUpdateFilter.Enabled = True
    If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
End Sub

Private Sub ckcCntrType_Click(Index As Integer)
    If bmFirstActivate Then Exit Sub
    ProgressBar1.Value = 0
    mApplyFilter
End Sub

Private Sub cmcAddFilter_Click()
    'Add Filter
    Dim slFilter As String
    slFilter = mMakeFilterString
    If slFilter <> "" Then
        If InStr(1, txtCriteria.Text, slFilter) > 0 Then
            mClearFilterInputs
            Exit Sub
        End If
        If txtCriteria.Text <> "" Then
            If right(txtCriteria.Text, 1) <> " " Then
                txtCriteria.Text = txtCriteria.Text & " "
            End If
        End If
        txtCriteria.Text = txtCriteria.Text & slFilter
    End If
    mCleanupCriteriaExtraSpaces
    
    Select Case cboFilterType
        Case "Advertiser": mManageAppliedFilters "Add", "Adv", mGetItemID("Advertiser", txtSearch.Text), txtSearch.Text
        Case "Agency": mManageAppliedFilters "Add", "Agy", mGetItemID("Agency", txtSearch.Text), txtSearch.Text
        Case "Active on or after": mManageAppliedFilters "Add", "*Active>", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy")
        Case "Active on or prior": mManageAppliedFilters "Add", "Active<", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy")
        Case "Update Date (Beginning)": mManageAppliedFilters "Add", "UDate>", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy")
        Case "Update Date (Ending)": mManageAppliedFilters "Add", "UDate<", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy")
        Case "Contract Number": mManageAppliedFilters "Add", "Cntr", Val(txtSearch.Text), Val(txtSearch.Text)
        Case "Contract Type": mManageAppliedFilters "Add", "CType", 0, txtSearch.Text
        Case "Contract Status": mManageAppliedFilters "Add", "CStat", 0, txtSearch.Text
        Case "Digital Delivery Status": mManageAppliedFilters "Add", "DlvyStat", 0, txtSearch.Text
        Case "Schedule Status": mManageAppliedFilters "Add", "SchStat", 0, txtSearch.Text
        Case "Product": mManageAppliedFilters "Add", "Prod", 0, txtSearch.Text
        Case "Salesperson": mManageAppliedFilters "Add", "SP", mGetItemID("Salesperson", txtSearch.Text), txtSearch.Text
        Case "Sales Group": mManageAppliedFilters "Add", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup))
        Case "Sales Office": mManageAppliedFilters "Add", "SO", mGetItemID("Sales Office", txtSearch.Text), txtSearch.Text
    End Select
    
    mClearFilterInputs
    If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
    cmcAddFilter.Enabled = False
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
    
    mPopulate
    lstFilterSearch.Clear
End Sub

Private Sub cmcCancel_Click()
    imCancelLoading = True
    dmTotalAmount = 0
    lmItemCount = 0
End Sub

Private Sub cmcChgCntr_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String
    Dim llCntrRowSelected As Long
    Dim slCntrStatus As String
    Dim blAllowPropEdit As Boolean
    Dim blAllowPropView As Boolean
    Dim blAllowCntrEdit As Boolean
    Dim blAllowCntrView As Boolean
    
    'Proposal (Edit/View) allowed?
    If tgSpf.sGUsePropSys = "Y" Then
        If igWinStatus(PROPOSALSJOB) = 0 Then
            blAllowPropEdit = False
            blAllowPropView = False
        Else
            If igWinStatus(PROPOSALSJOB) = 1 Then
                blAllowPropEdit = False
                blAllowPropView = True
            Else
                blAllowPropEdit = True
                blAllowPropView = True
            End If
        End If
    Else
        blAllowPropEdit = False
        blAllowPropView = False
    End If
    
    'Order (Edit/View) Allowed
    If igWinStatus(CONTRACTSJOB) = 0 Then
        blAllowCntrEdit = False
        blAllowCntrView = False
    Else
        If igWinStatus(CONTRACTSJOB) = 1 Then
            blAllowCntrEdit = False
            blAllowCntrView = True
        Else
            blAllowCntrEdit = True
            blAllowCntrView = True
        End If
    End If
    
    llCntrRowSelected = grdCntr.Row
    If grdCntr.Row <> grdCntr.RowSel Then llCntrRowSelected = 0
    
    If igJobShowing(CONTRACTSJOB) Then
        '4/5/10:  Enable is always False as Image is not active
        If igOKToCallCntr Then
            igTerminateAndUnload = True
            Unload Contract

            DoEvents
            Sleep 50
        Else
            ilRet = MsgBox("Unable to Change Contract until Current Altered Contract has been Saved", vbInformation + vbOKOnly, "Information")
            Exit Sub
        End If
    End If

    igDashboardCntrStatus = 0

    If llCntrRowSelected >= grdCntr.FixedRows Then
        If Trim$(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX)) = "" Then
            Exit Sub
        End If
        pbcFooter.Enabled = False
        Screen.MousePointer = vbHourglass

        lgAlertCntrNo = Val(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX))
        slCntrStatus = Trim$(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRSTATUSINDEX))
        
        'urfReviseCntr - Allow Revise holds and orders
        If tgUrf(0).sReviseCntr = "N" Then
            If slCntrStatus = "Approved Hold" Or slCntrStatus = "Hold" Or slCntrStatus = "Approved Order" Or slCntrStatus = "Order" Or slCntrStatus = "Rev Completed" Or slCntrStatus = "Rev Unapproved" Or slCntrStatus = "Rev Working" Then
                blAllowCntrEdit = False
                blAllowPropEdit = False
            End If
        End If

        Select Case slCntrStatus
            Case "Completed Proposal", "Unapproved Proposal", "Working Proposal", "Rejected"
                If blAllowPropEdit = False Then
                    GoTo Done
                End If
                igDashboardCntrStatus = 1
                sgCntrScreen = "Proposals"          'Change Complete
                
            Case "Rev Completed", "Rev Unapproved", "Rev Working"
                If blAllowPropEdit = False And blAllowCntrEdit = False Then
                    GoTo Done
                End If
                igDashboardCntrStatus = 2
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8
                If blAllowPropEdit = False And blAllowPropView = False Then
                    sgCntrScreen = "Orders"
                    igDashboardCntrStatus = 5       'View mode
                Else
                    sgCntrScreen = "Proposals"
                End If
                
            Case "Approved Order", "Approved Hold"
                If blAllowPropEdit = False And blAllowCntrEdit = False Then
                    GoTo Done
                End If
                igDashboardCntrStatus = 4
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                If blAllowCntrEdit = False And blAllowPropEdit Then
                    igDashboardCntrStatus = 2       'Revise existing Hold or Order
                    sgCntrScreen = "Proposals"
                Else
                    sgCntrScreen = "Orders"
                End If
                
            Case "Order", "Hold"
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                If blAllowPropEdit = False And blAllowCntrEdit = False Then
                    GoTo Done
                End If
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 6
                'igDashboardCntrStatus = 2
                igDashboardCntrStatus = 4
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                If blAllowCntrEdit = False And blAllowPropEdit Then
                    igDashboardCntrStatus = 2       'Revise existing Hold or Order
                    sgCntrScreen = "Proposals"
                Else
                    sgCntrScreen = "Orders"
                End If
                
        End Select
    End If
    
    Screen.MousePointer = vbDefault
    
    If igDashboardCntrStatus > 0 Then
        mTerminate
        Exit Sub
    End If
Done:
    'Force Reload
    Screen.MousePointer = vbDefault
    lgAlertCntrNo = 0
    smLastStartDate = ""
    smLastEndDate = ""
    mPopulate
    pbcFooter.Enabled = True
    Exit Sub

cmcChgCntrErr:         'VBC NR

    Screen.MousePointer = vbDefault
    ilRet = 1
    Resume Next
End Sub

Private Sub cmcClose_Click()
    igDashboardCntrStatus = -1
    lgAlertCntrNo = 0
    imReturnToDashboard = False
    mTerminate
End Sub

Private Sub cmcExport_Click()
    Call Me.PopupMenu(mnuExport)
End Sub

Private Sub cmcRemoveFilter_Click()
    'Delete selected filter
    Dim slFilter As String
    Dim slOldValue As String
    slOldValue = mGetFilterString(lblSelectedFilter.Caption)
    
    slFilter = mMakeFilterString(slOldValue)
    If slFilter <> "" Then
        txtCriteria.Text = Replace(txtCriteria.Text, lblSelectedFilter.Caption, "")
    End If
    
    mCleanupCriteriaExtraSpaces
    
    Select Case cboFilterType
        Case "Advertiser": mManageAppliedFilters "Remove", "Adv", 0, slOldValue
        Case "Agency": mManageAppliedFilters "Remove", "Agy", 0, slOldValue
        Case "Active on or after": mManageAppliedFilters "Remove", "*Active>", 0, Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Active on or prior": mManageAppliedFilters "Remove", "Active<", 0, Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Update Date (Beginning)": mManageAppliedFilters "Remove", "UDate>", 0, Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Update Date (Ending)": mManageAppliedFilters "Remove", "UDate<", 0, Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Contract Number": mManageAppliedFilters "Remove", "Cntr", Val(slOldValue), Val(slOldValue)
        Case "Contract Type": mManageAppliedFilters "Remove", "CType", 0, slOldValue
        Case "Contract Status": mManageAppliedFilters "Remove", "CStat", 0, slOldValue
        Case "Digital Delivery Status": mManageAppliedFilters "Remove", "DlvyStat", 0, slOldValue
        Case "Schedule Status": mManageAppliedFilters "Remove", "SchStat", 0, slOldValue
        Case "Product": mManageAppliedFilters "Remove", "Prod", 0, slOldValue
        Case "Salesperson": mManageAppliedFilters "Remove", "SP", 0, slOldValue
        Case "Sales Group": mManageAppliedFilters "Remove", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup))
        Case "Sales Office": mManageAppliedFilters "Remove", "SO", 0, slOldValue
    End Select
    
    cmcRemoveFilter.Enabled = False
    cmcUpdateFilter.Enabled = False
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
    
    mClearFilterInputs
    mPopulate
    cboFilterType.Enabled = True
    lstFilterSearch.Clear
    
    If cboFilterType = "Sales Group" Then
        txtSearch.Text = imCurrentGroup
        txtSearch_LostFocus
    End If
    
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
End Sub

Private Sub cmcSchedule_Click()
    Dim ilSchSelCntr As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim llChfCode As Long
    Dim slSchStatus As String
    Dim llCntrRowSelected As Integer
    
    ilSchSelCntr = False
    llCntrRowSelected = grdCntr.RowSel
    If llCntrRowSelected >= grdCntr.FixedRows Then
        If grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX) = "" Then
            Exit Sub
        End If
        llChfCode = grdCntr.TextMatrix(llCntrRowSelected, C_CHFCODEINDEX)
        slSchStatus = grdCntr.TextMatrix(llCntrRowSelected, C_SCHSTATUS)
        'If (slSchStatus = "A") Or (slSchStatus = "N") Then
        '    ilSchSelCntr = True
        'Else
        '    Exit Sub
        'End If
    End If
    
    Screen.MousePointer = vbHourglass
    pbcFooter.Enabled = False

    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & "#" & Trim$(str$(llChfCode))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & "#" & Trim$(str$(llChfCode))
    End If
    'Debug.Print slStr
    sgCommandStr = slStr
    CntrSch.Show vbModal
    slStr = sgDoneMsg
    pbcFooter.Enabled = True
    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To 5
        DoEvents
        Sleep (10)
    Next ilLoop
    
    'Force Reload
    smLastStartDate = ""
    smLastEndDate = ""
    sgCntrForDateStamp = ""
    mPopulate
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
mScheduleErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub cmcUpdateFilter_Click()
    'Update selected filter
    Dim slOldCriteria As String
    slOldCriteria = txtCriteria.Text
    
    Dim slFilter As String
    Dim slOldValue As String
    slOldValue = mGetFilterString(lblSelectedFilter.Caption)
    slFilter = mMakeFilterString()
    If slFilter <> "" Then
        txtCriteria.Text = Replace(txtCriteria.Text, lblSelectedFilter.Caption, slFilter)
    End If
    mCleanupCriteriaExtraSpaces
    
    Select Case cboFilterType
        Case "Advertiser": mManageAppliedFilters "Update", "Adv", mGetItemID("Advertiser", slOldValue), txtSearch.Text, slOldValue
        Case "Agency": mManageAppliedFilters "Update", "Agy", mGetItemID("Agency", txtSearch.Text), txtSearch.Text, slOldValue
        Case "Active on or after": mManageAppliedFilters "Update", "*Active>", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy"), Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Active on or prior": mManageAppliedFilters "Update", "Active<", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy"), Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Update Date (Beginning)": mManageAppliedFilters "Update", "UDate>", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy"), Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Update Date (Ending)": mManageAppliedFilters "Update", "UDate<", 0, Format(DateValue(txtSearch.Text), "mm/dd/yy"), Format(DateValue(slOldValue), "mm/dd/yy")
        Case "Contract Number": mManageAppliedFilters "Update", "Cntr", Val(txtSearch.Text), Val(txtSearch.Text), Val(slOldValue)
        Case "Contract Type": mManageAppliedFilters "Update", "CType", 0, txtSearch.Text, slOldValue
        Case "Contract Status": mManageAppliedFilters "Update", "CStat", 0, txtSearch.Text, slOldValue
        Case "Digital Delivery Status": mManageAppliedFilters "Update", "DlvyStat", 0, txtSearch.Text, slOldValue
        Case "Schedule Status": mManageAppliedFilters "Update", "SchStat", 0, txtSearch.Text, slOldValue
        Case "Product": mManageAppliedFilters "Update", "Prod", 0, txtSearch.Text, slOldValue
        Case "Salesperson": mManageAppliedFilters "Update", "SP", mGetItemID("Salesperson", txtSearch.Text), txtSearch.Text, slOldValue
        Case "Sales Group": mManageAppliedFilters "Update", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup)), Trim(str(imCurrentGroup))
        Case "Sales Office": mManageAppliedFilters "Update", "SO", mGetItemID("Sales Office", txtSearch.Text), txtSearch.Text, slOldValue
    End Select
    
    mClearFilterInputs
    cmcUpdateFilter.Enabled = False
    cmcRemoveFilter.Enabled = False
    mCleanupCriteriaExtraSpaces
    
    If Trim(txtCriteria.Text) <> Trim(slOldCriteria) Then mPopulate
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
    
    cboFilterType.Enabled = True
    lstFilterSearch.Clear
End Sub

Private Sub cmcViewCntr_Click()
    Dim ilRet As Integer
    Dim slStr As String
    Dim llCntrRowSelected As Integer
    Dim slCntrStatus As String
    Dim blAllowPropEdit As Boolean
    Dim blAllowPropView As Boolean
    Dim blAllowCntrEdit As Boolean
    Dim blAllowCntrView As Boolean
    
    'Proposal (Edit/View) allowed?
    If tgSpf.sGUsePropSys = "Y" Then
        If igWinStatus(PROPOSALSJOB) = 0 Then
            blAllowPropEdit = False
            blAllowPropView = False
        Else
            If igWinStatus(PROPOSALSJOB) = 1 Then
                blAllowPropEdit = False
                blAllowPropView = True
            Else
                blAllowPropEdit = True
                blAllowPropView = True
            End If
        End If
    Else
        blAllowPropEdit = False
        blAllowPropView = False
    End If
    
    'Order (Edit/View) Allowed
    If igWinStatus(CONTRACTSJOB) = 0 Then
        blAllowCntrEdit = False
        blAllowCntrView = False
    Else
        If igWinStatus(CONTRACTSJOB) = 1 Then
            blAllowCntrEdit = False
            blAllowCntrView = True
        Else
            blAllowCntrEdit = True
            blAllowCntrView = True
        End If
    End If
    
    llCntrRowSelected = grdCntr.Row
    If grdCntr.Row <> grdCntr.RowSel Then llCntrRowSelected = 0
    
    If igJobShowing(CONTRACTSJOB) Then
        If igOKToCallCntr Then
            igTerminateAndUnload = True
            Unload Contract
            DoEvents
            Sleep 50
        Else
            ilRet = MsgBox("Unable to view Contract until Current Altered Contract has been Saved", vbInformation + vbOKOnly, "Information")
            Exit Sub
        End If
    End If
    igDashboardCntrStatus = 0
    If llCntrRowSelected >= grdCntr.FixedRows Then
        If Trim$(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX)) = "" Then
            Exit Sub
        End If
        
        pbcFooter.Enabled = False
        Screen.MousePointer = vbHourglass
        lgAlertCntrNo = Val(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX))
        slCntrStatus = Trim$(grdCntr.TextMatrix(llCntrRowSelected, C_CNTRSTATUSINDEX))
        
        Select Case slCntrStatus
            Case "Completed Proposal", "Working Proposal", "Rejected", "Rev Completed", "Rev Unapproved", "Rev Working"
                'If blAllowPropView = False Then
                '    GoTo Done
                'End If
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8
                If blAllowPropEdit = False And blAllowPropView = False Then
                    sgCntrScreen = "Orders"
                    igDashboardCntrStatus = 5       'View mode
                Else
                    igDashboardCntrStatus = 3
                    sgCntrScreen = "Proposals"
                End If
                
            Case "Unapproved Proposal"
                If blAllowPropView = False Then
                    GoTo Done
                End If
                igDashboardCntrStatus = 3
                sgCntrScreen = "Proposals"
                
            Case "Order", "Hold", "Approved Order", "Approved Hold"
                If blAllowCntrView = False And blAllowPropView = False Then
                    GoTo Done
                End If
                igDashboardCntrStatus = 5
                 'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                If blAllowCntrView = False And blAllowPropView Then
                    sgCntrScreen = "Proposals"
                    igDashboardCntrStatus = 3
                Else
                    sgCntrScreen = "Orders"
                End If
                
        End Select
    End If
    
    Screen.MousePointer = vbDefault
    
    If igDashboardCntrStatus > 0 Then
        mTerminate
        Exit Sub
    End If
Done:
    Screen.MousePointer = vbDefault
    lgAlertCntrNo = 0
    pbcFooter.Enabled = True
    Exit Sub
cmcViewCntrErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub CSI_Calendar1_Change()
    If CSI_Calendar1.Text = "" Then Exit Sub
    lstFilterSearch.Visible = False
    txtSearch.Text = DateValue(CSI_Calendar1.Text)
    If txtSearch.Enabled = True And txtSearch.Visible = True Then txtSearch.SetFocus
    txtSearch_LostFocus
End Sub

Private Sub CSI_Calendar1_GotFocus()
    lstFilterSearch.Visible = False
End Sub

Private Sub CSI_Calendar1_DateClicked()
    CSI_Calendar1_Change
End Sub

Private Sub grdCntr_Click()
    'Sort
    If grdCntr.MouseRow < grdCntr.FixedRows Then
        mSortByColumn grdCntr.MouseCol
        Exit Sub
    End If
    
    'Don't allow action buttons when multi-rows selected
    If grdCntr.Row <> grdCntr.RowSel Then
        cmcSchedule.Enabled = False
        cmcChgCntr.Enabled = False
        cmcViewCntr.Enabled = False
    End If
End Sub

Private Sub grdCntr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slString As String
    If grdCntr.Row = 0 Then Exit Sub
    If grdCntr.Row <> grdCntr.RowSel Then Exit Sub
    mnuFilterBy2.Visible = False
    
    If grdCntr.MouseRow = grdCntr.Row Then
        If Button = 2 Then
            Select Case grdCntr.MouseCol
                Case C_CHFCODEINDEX: Exit Sub
                Case C_SORTINDEX: Exit Sub
                Case C_CNTRUPDATEDATEINDEX:
                    mnuFilterBy2.Visible = True
                    mnuFilterBy.Caption = "Filter: [UDate>:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                    mnuFilterBy2.Caption = "Filter: [UDate<:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_CNTRNOINDEX: Exit Sub
                Case C_LINETYPEINDEX: Exit Sub
                Case C_CNTRTYPEINDEX:
                    mnuFilterBy.Caption = "Filter: [CType:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_AGYNAMEINDEX:
                    mnuFilterBy.Caption = "Filter: [Agy:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_ADVNAMEINDEX:
                    mnuFilterBy.Caption = "Filter: [Adv:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_PRODUCTINDEX:
                    If Trim(grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol)) = "" Then Exit Sub
                    mnuFilterBy.Caption = "Filter: [Prod:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_STARTDATEINDEX: Exit Sub
                Case C_ENDDATEINDEX: Exit Sub
                Case C_GROSSINDEX: Exit Sub
                Case C_SALESOFFICEINDEX:
                    mnuFilterBy.Caption = "Filter: [SO:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_SALEPERSONINDEX:
                    mnuFilterBy.Caption = "Filter: [SP:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_CNTRSTATUSINDEX:
                    mnuFilterBy.Caption = "Filter: [CStat:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_DIGITALDLVYINDEX:
                    mnuFilterBy.Caption = "Filter: [DlvyStat:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
                Case C_CNTRSCHEDULESTATUSINDEX:
                    mnuFilterBy.Caption = "Filter: [SchStat:" & grdCntr.TextMatrix(grdCntr.MouseRow, grdCntr.MouseCol) & "]"
            End Select
            
            Call Me.PopupMenu(mnuContext)
        End If
    End If
End Sub

Private Sub grdCntr_RowColChange()
    Dim llCntrRowSelected As Integer
    Dim slString As String
    Dim blAllowPropEdit As Boolean
    Dim blAllowPropView As Boolean
    Dim blAllowCntrEdit As Boolean
    Dim blAllowCntrView As Boolean
    cmcSchedule.Enabled = False
    cmcChgCntr.Enabled = False
    cmcViewCntr.Enabled = False
    
    'Proposal (Edit/View) allowed?
    If tgSpf.sGUsePropSys = "Y" Then
        If igWinStatus(PROPOSALSJOB) = 0 Then
            blAllowPropEdit = False
            blAllowPropView = False
        Else
            If igWinStatus(PROPOSALSJOB) = 1 Then
                blAllowPropEdit = False
                blAllowPropView = True
            Else
                blAllowPropEdit = True
                blAllowPropView = True
            End If
        End If
    Else
        blAllowPropEdit = False
        blAllowPropView = False
    End If
    
    'Order (Edit/View) Allowed
    If igWinStatus(CONTRACTSJOB) = 0 Then
        blAllowCntrEdit = False
        blAllowCntrView = False
    Else
        If igWinStatus(CONTRACTSJOB) = 1 Then
            blAllowCntrEdit = False
            blAllowCntrView = True
        Else
            blAllowCntrEdit = True
            blAllowCntrView = True
        End If
    End If
    
    'Select row
    llCntrRowSelected = grdCntr.Row
    
    If llCntrRowSelected > 0 Then
        slCntrStatus = grdCntr.TextMatrix(llCntrRowSelected, C_CNTRSTATUSINDEX)
        
        'urfReviseCntr - Allow Revise holds and orders
        If tgUrf(0).sReviseCntr = "N" Then
            If slCntrStatus = "Approved Hold" Or slCntrStatus = "Hold" Or slCntrStatus = "Approved Order" Or slCntrStatus = "Order" Or slCntrStatus = "Rev Completed" Or slCntrStatus = "Rev Unapproved" Or slCntrStatus = "Rev Working" Then
                blAllowCntrEdit = False
                blAllowPropEdit = False
            End If
        End If
                
        Select Case slCntrStatus
            Case "Working Proposal"
                cmcChgCntr.Caption = "Change Proposal"
                cmcViewCntr.Caption = "View Proposal"
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                'v81 TTP 10937 - testing 2/12/24 9:11 AM - Issue 11
                If blAllowPropView Or blAllowCntrView Then cmcViewCntr.Enabled = True
                
            Case "Rejected"
                cmcChgCntr.Caption = "Change Proposal"
                cmcViewCntr.Caption = "View Proposal"
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowPropView Then cmcViewCntr.Enabled = True
                
            Case "Completed Proposal"
                cmcChgCntr.Caption = "Change Proposal"
                cmcViewCntr.Caption = "View Proposal"
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                'Found another one like Issue 11
                If blAllowPropView Or blAllowCntrView Then cmcViewCntr.Enabled = True
                
            Case "Unapproved Proposal"
                cmcChgCntr.Caption = "Change Proposal"
                cmcViewCntr.Caption = "View Proposal"
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowPropView Then cmcViewCntr.Enabled = True
                
            Case "Approved Hold"
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                If blAllowCntrEdit Or blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
                cmcSchedule.Enabled = True
                
            Case "Hold"
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                If blAllowCntrEdit Or blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
                
            Case "Approved Order":
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                If blAllowCntrEdit Or blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 9
                cmcSchedule.Enabled = True
                
            Case "Order":
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                If blAllowCntrEdit Or blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
                
            Case "Rev Completed":
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8 (Dont allow edit Rev working if Proposals edit is not permitted)
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
            
            Case "Rev Unapproved":
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8 (Dont allow edit Rev working if Proposals edit is not permitted)
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
            
            Case "Rev Working":
                cmcChgCntr.Caption = "Change Order"
                cmcViewCntr.Caption = "View Order"
                'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
                'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8 (Dont allow edit Rev working if Proposals edit is not permitted)
                If blAllowPropEdit Then cmcChgCntr.Enabled = True
                If blAllowCntrView Or blAllowPropView Then cmcViewCntr.Enabled = True
            
        End Select
    End If
End Sub

Private Sub lvcColumns_Click()
    Dim ilLoop As Integer
    Dim ilColumn As Integer
    ilColumn = C_CNTRUPDATEDATEINDEX
    For ilLoop = 1 To lvcColumns.ListItems.Count
        If lvcColumns.ListItems(ilLoop).Checked = True Then
            If grdCntr.ColWidth(ilLoop + 5) < 1 Then
                imColumnWidths(ilLoop + 5) = -1
                mSetDefaultColumnSizes ilLoop + 5, True
                grdCntr.ColWidth(ilLoop + 5) = mGetColumnSize(ilLoop + 5)
            End If
        Else
            If grdCntr.ColWidth(ilLoop + 5) > 0 Then
                grdCntr.ColWidth(ilLoop + 5) = 0
                imColumnWidths(ilLoop + 5) = -1
            End If
        End If
    Next ilLoop
End Sub

Private Sub lstFilterSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If lstFilterSearch.ListIndex = 0 Then
            If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
        End If
    End If
    If KeyCode = vbKeyReturn Then
        txtSearch.Text = lstFilterSearch.Text
        If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
        lstFilterSearch.Visible = False
    End If
End Sub

Private Sub lstFilterSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtSearch.Text = lstFilterSearch.Text
    txtSearch_KeyUp vbKeyReturn, 0
    lstFilterSearch.Visible = False
    If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
    If mCheckCanAdd = True Then
        cmcAddFilter.Enabled = True
    Else
        cmcAddFilter.Enabled = False
    End If
End Sub

Private Sub mnuCSV_Click()
    'Export to CSV
    mExportToCSV
End Sub

Private Sub mnuExcel_Click()
    'Open selection in Excel
    mSendToExcel
End Sub

Private Sub mnuFilterBy2_Click()
    Dim slString As String
    Dim slvalue As String
    slString = Mid(mnuFilterBy2.Caption, 9)
    If InStr(1, txtCriteria.Text, slString) > 0 Then Exit Sub
    
    slvalue = mGetFilterString(slString)
    If Mid(txtCriteria.Text, Len(txtCriteria.Text) - 1, 1) <> "" Then txtCriteria.Text = txtCriteria.Text & " "
    txtCriteria.Text = txtCriteria.Text & slString
    
    If InStr(1, slString, "UDate>:") > 0 Then mManageAppliedFilters "Add", "UDate>", 0, slvalue
    If InStr(1, slString, "UDate<:") > 0 Then mManageAppliedFilters "Add", "UDate<", 0, slvalue
    If InStr(1, slString, "Adv:") > 0 Then mManageAppliedFilters "Add", "Adv", mGetItemID("Advertiser", slvalue), slvalue
    If InStr(1, slString, "Agy:") > 0 Then mManageAppliedFilters "Add", "Agy", mGetItemID("Agency", slvalue), slvalue
    If InStr(1, slString, "CType:") > 0 Then mManageAppliedFilters "Add", "CType", 0, slvalue
    If InStr(1, slString, "CStat:") > 0 Then mManageAppliedFilters "Add", "CStat", 0, slvalue
    If InStr(1, slString, "DlvyStat:") > 0 Then mManageAppliedFilters "Add", "DlvyStat", 0, slvalue
    If InStr(1, slString, "SchStat:") > 0 Then mManageAppliedFilters "Add", "SchStat", 0, slvalue
    If InStr(1, slString, "Prod:") > 0 Then mManageAppliedFilters "Add", "Prod", 0, slvalue
    If InStr(1, slString, "SP:") > 0 Then mManageAppliedFilters "Add", "SP", mGetItemID("Salesperson", slvalue), slvalue
    If InStr(1, slString, "Group:") > 0 Then mManageAppliedFilters "Add", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup))
    If InStr(1, slString, "SO:") > 0 Then mManageAppliedFilters "Add", "SO", mGetItemID("Sales Office", slvalue), slvalue
    mCleanupCriteriaExtraSpaces
    
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True

    mPopulate
End Sub

Private Sub mnuFilterBy_Click()
    Dim slString As String
    Dim slvalue As String
    slString = Mid(mnuFilterBy.Caption, 9)
    If InStr(1, txtCriteria.Text, slString) > 0 Then Exit Sub
    
    slvalue = mGetFilterString(slString)
    If Mid(txtCriteria.Text, Len(txtCriteria.Text) - 1, 1) <> "" Then txtCriteria.Text = txtCriteria.Text & " "
    txtCriteria.Text = txtCriteria.Text & slString
    
    If InStr(1, slString, "UDate>:") > 0 Then mManageAppliedFilters "Add", "UDate>", 0, slvalue
    If InStr(1, slString, "UDate<:") > 0 Then mManageAppliedFilters "Add", "UDate<", 0, slvalue
    If InStr(1, slString, "Adv:") > 0 Then mManageAppliedFilters "Add", "Adv", mGetItemID("Advertiser", slvalue), slvalue
    If InStr(1, slString, "Agy:") > 0 Then mManageAppliedFilters "Add", "Agy", mGetItemID("Agency", slvalue), slvalue
    If InStr(1, slString, "CType:") > 0 Then mManageAppliedFilters "Add", "CType", 0, slvalue
    If InStr(1, slString, "CStat:") > 0 Then mManageAppliedFilters "Add", "CStat", 0, slvalue
    If InStr(1, slString, "DlvyStat:") > 0 Then mManageAppliedFilters "Add", "DlvyStat", 0, slvalue
    If InStr(1, slString, "SchStat:") > 0 Then mManageAppliedFilters "Add", "SchStat", 0, slvalue
    If InStr(1, slString, "Prod:") > 0 Then mManageAppliedFilters "Add", "Prod", 0, slvalue
    If InStr(1, slString, "SP:") > 0 Then mManageAppliedFilters "Add", "SP", mGetItemID("Salesperson", slvalue), slvalue
    If InStr(1, slString, "Group:") > 0 Then mManageAppliedFilters "Add", "Group", CLng(imCurrentGroup), Trim(str(imCurrentGroup))
    If InStr(1, slString, "SO:") > 0 Then mManageAppliedFilters "Add", "SO", mGetItemID("Sales Office", slvalue), slvalue
    mCleanupCriteriaExtraSpaces
    
    cmcResetFilter.Enabled = False
    If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True

    mPopulate
End Sub

Private Sub Timer1_Timer()
    Dim ilFound As Integer
    Dim objForm As Form
    
    'Export button
    If lmItemCount <> 0 Then
        cmcExport.Enabled = True
    Else
        cmcExport.Enabled = False
    End If
    
    'progress bar / label / button position
    ilFound = False
    If ProgressBar1 = 100 Then
        ProgressBar1.Visible = False
        cmcCancel.Visible = False
        cmcClose.Enabled = True
        lblItemCount.Left = pbcFooter.Width - lblItemCount.Width - 120
        lblMultiVersionTotalWarn.Left = lblItemCount.Left
        lblItemCount.Caption = "Showing " & Format(lmItemCount, "#,##0") & " Items / " & Format(dmTotalAmount, "$#,##0.00") & " Gross"
    Else
        ProgressBar1.Visible = True
        cmcCancel.Visible = True
        ProgressBar1.Left = cmcCancel.Left - ProgressBar1.Width - 80
        lblItemCount.Left = ProgressBar1.Left - lblItemCount.Width - 120
        lblMultiVersionTotalWarn.Left = lblItemCount.Left
    End If
    
    'Delay cboItems filtering until keys stopped being pressed
    If imKeyDelay = 0 Then
        imKeyDelay = 0
        mLoadFilterSearch
        imKeyDelay = imKeyDelay - 1
        Exit Sub
    End If
    If imKeyDelay = -1 Then
        Exit Sub
    End If
    imKeyDelay = imKeyDelay - 1
End Sub

Private Sub txtCriteria_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Trim(txtCriteria.Text) = "" Then Exit Sub
    mSelectFilterItem
End Sub

Private Sub txtSearch_Change()
    If txtSearch.Text <> "" Then
        cmcResetFilter.Enabled = False
    Else
        If Trim(txtCriteria.Text) <> Trim(smDefaultFilters) Then cmcResetFilter.Enabled = True
    End If
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 32 And KeyCode <> 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If InStr(1, cboFilterType.Text, "Date") > 0 Then
            If mValidateDate(txtSearch.Text) = True Then
                If cmcUpdateFilter.Enabled = True And cmcUpdateFilter.Visible = True Then
                    cmcUpdateFilter_Click
                End If
                If cmcAddFilter.Enabled = True And cmcAddFilter.Visible = True Then
                    cmcAddFilter_Click
                End If
                Exit Sub
            Else
                cmcAddFilter.Enabled = False
                cmcUpdateFilter.Enabled = True
            End If
        ElseIf InStr(1, cboFilterType.Text, "Contract Number") > 0 Then
            If IsNumeric(txtSearch.Text) Then
                cmcAddFilter_Click
            End If
        ElseIf InStr(1, cboFilterType.Text, "Product") > 0 Then
            If Len(Trim(txtSearch.Text)) > 0 Then
                cmcAddFilter_Click
            End If
        Else
            If mFindInItems(txtSearch.Text) = True Then
                If cmcAddFilter.Enabled = True And cmcAddFilter.Visible = True Then
                    cmcAddFilter_Click
                End If
                If cmcUpdateFilter.Enabled = True And cmcUpdateFilter.Visible = True Then
                    cmcUpdateFilter_Click
                End If
                Exit Sub
            End If
            If lstFilterSearch.Visible = True Then
                If lstFilterSearch.ListIndex > -1 Then
                    txtSearch.Text = lstFilterSearch.Text
                    cmcAddFilter.Enabled = True
                    cmcUpdateFilter.Enabled = True
                End If
            End If
        End If
    End If
    
    If lstFilterSearch.Visible = False Then Exit Sub
    
    If KeyCode = vbKeyDown Then
        If lstFilterSearch.ListCount = 0 Then Exit Sub
            If lstFilterSearch.Visible = True Then
            If lstFilterSearch.ListCount > 0 Then
                lstFilterSearch.Selected(0) = True
            End If
            If lstFilterSearch.Visible = True Then lstFilterSearch.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If lstFilterSearch.ListCount = 0 Then Exit Sub
        If lstFilterSearch.ListIndex = -1 Then Exit Sub
        txtSearch.Text = lstFilterSearch.Text
        lstFilterSearch.Visible = False
        If mCheckCanAdd = True Then
            cmcAddFilter.Enabled = True
        Else
            cmcAddFilter.Enabled = False
        End If
        cmcUpdateFilter.Enabled = True
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSearch_LostFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    If txtSearch.Text = "" Then lstFilterSearch.Visible = False
    If KeyCode < 32 And (KeyCode <> 8 And KeyCode <> 13) Then Exit Sub
    
    If Trim(txtSearch.Text) = "" Then
        lstFilterSearch.Visible = False
        cmcAddFilter.Enabled = False
        cmcUpdateFilter.Enabled = False
        Exit Sub
    Else
        If InStr(1, cboFilterType.Text, "Date") > 0 Or InStr(1, cboFilterType.Text, "Active on") > 0 Then   'Dont' show filter list if Date
            'Check if Valid Date typed in
            If mValidateDate(txtSearch.Text) = True Then
                If mCheckCanAdd = True Then
                    cmcAddFilter.Enabled = True
                Else
                    cmcAddFilter.Enabled = False
                End If
                cmcUpdateFilter.Enabled = True
            Else
                cmcAddFilter.Enabled = False
                cmcUpdateFilter.Enabled = False
            End If
        ElseIf InStr(1, cboFilterType.Text, "Product") > 0 Then
            If Len(Trim(txtSearch.Text)) > 0 Then
                If mCheckCanAdd = True Then
                    cmcAddFilter.Enabled = True
                Else
                    cmcAddFilter.Enabled = False
                End If
                cmcUpdateFilter.Enabled = True
            Else
                cmcAddFilter.Enabled = False
                cmcUpdateFilter.Enabled = False
            End If
        ElseIf InStr(1, cboFilterType.Text, "Contract Number") > 0 Then
            If IsNumeric(txtSearch.Text) Then
                If mCheckCanAdd = True Then
                    cmcAddFilter.Enabled = True
                Else
                    cmcAddFilter.Enabled = False
                End If
                cmcUpdateFilter.Enabled = True
            Else
                cmcAddFilter.Enabled = False
                cmcUpdateFilter.Enabled = False
            End If
        Else
            'Check if txtSearch has a qualified item typed in
            If LCase(txtSearch.Text) <> LCase(lstFilterSearch.Text) Or mFindInItems(txtSearch.Text) = False Then
                cmcAddFilter.Enabled = False
                cmcUpdateFilter.Enabled = False
                If cboItems.Visible = True Then lstFilterSearch.Visible = True
            Else
                If mCheckCanAdd = True Then
                    cmcAddFilter.Enabled = True
                Else
                    cmcAddFilter.Enabled = False
                End If
                cmcUpdateFilter.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        txtSearch_LostFocus
        KeyCode = 0
        Exit Sub
    End If
    
    imKeyDelay = 2 'Delays loading the SearchFilterlist
End Sub

Private Sub txtSearch_LostFocus()
    'Validate Search
    If InStr(1, cboFilterType.Text, "Date") > 0 Or InStr(1, cboFilterType.Text, "Active on") > 0 Then
        If mValidateDate(txtSearch.Text) = True Then
            cmcUpdateFilter.Enabled = True
            If mCheckCanAdd = True Then
                cmcAddFilter.Enabled = True
            Else
                cmcAddFilter.Enabled = False
            End If
            Exit Sub
        Else
            cmcUpdateFilter.Enabled = False
            cmcAddFilter.Enabled = False
        End If
    ElseIf InStr(1, cboFilterType.Text, "Sales Group") > 0 Then
        If imCurrentGroup > 0 Then
            If mCheckCanAdd = True Then
                cmcAddFilter.Enabled = True
            Else
                cmcAddFilter.Enabled = False
            End If
        End If
    Else
        If lstFilterSearch.Visible = False Then Exit Sub
        If DashboardVw.ActiveControl.Name = "lstFilterSearch" Then Exit Sub
        mFindInItems (txtSearch.Text)
    End If
    
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
    lstFilterSearch.Visible = False
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mCheckCanAdd
' Description:       Checks if Add button can be Enabled, False if cboFilter is set to a value already existing in Criteria
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-12:00:45
' Parameters :
'--------------------------------------------------------------------------------
Function mCheckCanAdd() As Boolean
    mCheckCanAdd = True
    'Check if a selected type = Active Date , Make sure Is not already in Filter criteria
    If cboFilterType.Text = "Active on or after" Then
        If InStr(1, txtCriteria.Text, "[*Active>:") > 0 Then
            mCheckCanAdd = False
        End If
    End If
    If cboFilterType.Text = "Active on or after" Then
        If InStr(1, txtCriteria.Text, "[Active>:") > 0 Then
            mCheckCanAdd = False
        End If
    End If

    If cboFilterType.Text = "Active on or prior" Then
        If InStr(1, txtCriteria.Text, "[Active<:") > 0 Then
            mCheckCanAdd = False
        End If
    End If

    'Check if a selected type = Update Date , Make sure Is not already in Filter criteria
    If cboFilterType.Text = "Update Date (Beginning)" Then
        If InStr(1, txtCriteria.Text, "[UDate>:") > 0 Then
            mCheckCanAdd = False
        End If
    End If
    If cboFilterType.Text = "Update Date (Ending)" Then
        If InStr(1, txtCriteria.Text, "[UDate<:") > 0 Then
            mCheckCanAdd = False
        End If
    End If

    Dim slFilter As String
    slFilter = mMakeFilterString
    If slFilter <> "" Then
        If InStr(1, txtCriteria.Text, slFilter) > 0 Then
            mCheckCanAdd = False
        End If
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mCheckIfNoFilter
' Description:       If No filter for the specified filter Type, return true
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:52:58
' Parameters :       slFilterType (String)
'--------------------------------------------------------------------------------
Function mCheckIfNoFilter(slFilterType As String) As Boolean
    mCheckIfNoFilter = True
    Dim ilAppliedFiltersLoop As Integer
    For ilAppliedFiltersLoop = 0 To UBound(tmAppliedFilters)
        If tmAppliedFilters(ilAppliedFiltersLoop).sType = slFilterType Then
            mCheckIfNoFilter = False
            Exit For
        End If
    Next ilAppliedFiltersLoop
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mCleanupCriteriaExtraSpaces
' Description:       Removes extra spaces from txtCriteria at the beginning, between filters, and at the end
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:35:58
' Parameters :
'--------------------------------------------------------------------------------
Sub mCleanupCriteriaExtraSpaces()
    If Trim(txtCriteria.Text) = "" Then txtCriteria.Text = ""
    If Left(txtCriteria.Text, 1) = " " Then txtCriteria.Text = Mid(txtCriteria.Text, 2)
    If Left(txtCriteria.Text, 2) = "  " Then txtCriteria.Text = Mid(txtCriteria.Text, 3)
    txtCriteria.Text = Replace(txtCriteria.Text, "]  [", "] [")
    txtCriteria.Text = Replace(txtCriteria.Text, "]   [", "] [")
    If right(txtCriteria.Text, 2) = "  " Then txtCriteria.Text = Mid(txtCriteria.Text, 1, Len(txtCriteria.Text) - 1)
    If right(txtCriteria.Text, 2) = "  " Then txtCriteria.Text = Mid(txtCriteria.Text, 1, Len(txtCriteria.Text) - 1)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mFillSortColumn
' Description:       Fills C_SORTINDEX with sortable values from the specified Date or Numeric column
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:06:09
' Parameters :       ilSourceColumn (Integer)
'                    slType (String)
'                    blSetProgress (Boolean = False)
'--------------------------------------------------------------------------------
Sub mFillSortColumn(ilSourceColumn As Integer, slType As String, Optional blSetProgress As Boolean = False)
    Dim llLoop As Long
    Dim slString As String
    Dim ilProgress As Integer
    Dim dlProgressStep As Double
    Dim ilDOE As Integer
    dlProgressStep = 90 / grdCntr.rows
    If blSetProgress = True Then ProgressBar1.Value = 0
    
    For llLoop = 1 To grdCntr.rows - 1
        'Progress bar
        If blSetProgress Then
            ilDOE = ilDOE + 1
            If ilDOE > 1000 Then
                dlProgress = dlProgressStep * llLoop
                ProgressBar1.Value = dlProgress
                ilDOE = 0
                DoEvents
            End If
        End If
        If grdCntr.TextMatrix(llLoop, ilSourceColumn) = "" Then
            grdCntr.TextMatrix(llLoop, C_SORTINDEX) = ""
        Else
            If slType = "Date" Then
                grdCntr.TextMatrix(llLoop, C_SORTINDEX) = Format(DateValue(grdCntr.TextMatrix(llLoop, ilSourceColumn)), "yyyymmdd")
            End If
            If slType = "Number" Then
                slString = Replace(grdCntr.TextMatrix(llLoop, ilSourceColumn), ",", "")
                slString = Replace(slString, " ", "")
                slString = Replace(slString, "V", ".")
                slString = Format(Val(slString), "0000000000.00")
                grdCntr.TextMatrix(llLoop, C_SORTINDEX) = slString
            End If
        End If
    Next llLoop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mFindInItems
' Description:       make sure we have a item that is present in the cboItems list, returns True if Found
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:02:19
' Parameters :       slItem (String)
'--------------------------------------------------------------------------------
Function mFindInItems(slItem As String) As Boolean
    Dim ilLoop As Integer
    Dim ilSelPos As Integer
    Dim ilSelLen As Integer
    ilSelPos = txtSearch.SelStart
    ilSelLen = txtSearch.SelLength
    For ilLoop = 0 To cboItems.ListCount - 1
        If LCase(cboItems.List(ilLoop)) = LCase(slItem) Then
            mFindInItems = True
            txtSearch.Text = cboItems.List(ilLoop)
            If ilSelLen = 0 Then
                txtSearch.SelStart = ilSelPos
            End If
            lstFilterSearch.Visible = False
            Exit For
        End If
    Next ilLoop
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetAdvertiserName
' Description:       Get Advertiser Name for the provided ID
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:26:17
' Parameters :       ilNumber (Integer)
'--------------------------------------------------------------------------------
Function mGetAdvertiserName(ilNumber As Integer) As String
    Dim ilInx As Integer
    If ilNumber = imLastAdv Then
        mGetAdvertiserName = smLastAdvName
        Exit Function
    End If
    ilInx = gBinarySearchAdf(ilNumber)
    If ilInx >= 0 Then              '10-10-18
        mGetAdvertiserName = tgCommAdf(ilInx).sName
        imLastAdv = ilNumber
        smLastAdvName = mGetAdvertiserName
    Else
        mGetAdvertiserName = "N/A"
        imLastAdv = ilNumber
        smLastAdvName = mGetAdvertiserName
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetAgencyName
' Description:       Get Agency Name for the provided ID
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:25:23
' Parameters :       ilNumber (Integer)
'--------------------------------------------------------------------------------
Function mGetAgencyName(ilNumber As Integer) As String
    Dim ilInx As Integer
    If ilNumber = imLastAgy Then
        mGetAgencyName = smLastAgyName
        Exit Function
    End If
    ilInx = gBinarySearchAgf(ilNumber)
    If ilInx >= 0 Then              '10-10-18
        mGetAgencyName = Trim(tgCommAgf(ilInx).sName) & ", " & Trim$(tgCommAgf(ilInx).sCityID)
        imLastAgy = ilNumber
        smLastAgyName = mGetAgencyName
    Else
        mGetAgencyName = "N/A"
        imLastAgy = ilNumber
        smLastAgyName = mGetAgencyName
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetCntrStatus
' Description:       Get Contract Status description for the provided code
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:18:42
' Parameters :       slCntrStatusCode (String)
'--------------------------------------------------------------------------------
Private Function mGetCntrStatus(tlCntr As CHFDETAILLIST) As String
    Dim slString As String
    Dim slCntrStatusCode As String
    slCntrStatusCode = tlCntr.sStatus
    
    'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 3
    Select Case slCntrStatusCode
        Case "C": slString = "Completed Proposal"
            If tlCntr.iCntRevNo > 0 And tlCntr.sSchStatus = "P" Then slString = "Rev Completed"
        Case "W": slString = "Working Proposal"
            If tlCntr.iCntRevNo > 0 And tlCntr.sSchStatus = "P" Then slString = "Rev Working"
        Case "I": slString = "Unapproved Proposal"
            If tlCntr.iCntRevNo > 0 And tlCntr.sSchStatus = "P" Then slString = "Rev Unapproved"
        Case "N": slString = "Approved Order"
            'v81 TTP 10937 - testing 2/8/24 3:03 PM - Issue 8 (this is not an Unapproved status):
            'If tlCntr.iCntRevNo > 0 And tlCntr.sSchStatus = "A" Then slString = "Unapproved"
        Case "D": slString = "Rejected"
        Case "G": slString = "Approved Hold"
        Case "H": slString = "Hold"
        Case "O": slString = "Order"
    End Select
    mGetCntrStatus = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetCntrType
' Description:       Get Contract Type description for the provided code
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:19:22
' Parameters :       slCntrTypeCode (String)
'--------------------------------------------------------------------------------
Function mGetCntrType(slCntrTypeCode As String) As String
    Dim slString As String
    Select Case slCntrTypeCode
        Case "C": slString = "Standard"
        Case "V": slString = "Reservation"
        Case "T": slString = "Remnant"
        Case "R": slString = "Direct Response"
        Case "Q": slString = "Per Inquiry"
        Case "S": slString = "PSA"
        Case "M": slString = "Promo"
    End Select
    mGetCntrType = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetDeliveryStatus
' Description:       Get Delivery Status description for provided code
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:17:16
' Parameters :       slDlvyStatusCode (Integer)
'--------------------------------------------------------------------------------
Function mGetDeliveryStatus(slDlvyStatusCode As Integer) As String
    Dim slString As String
    Select Case slDlvyStatusCode
        Case 0: slString = "N/A"
        Case 1: slString = "Not Pushed"
        Case 2: slString = "Pushed"
        Case 3: slString = "Partial"
        Case 4: slString = "Issue Encountered"
        Case 5: slString = "Requires Repush"
    End Select
    mGetDeliveryStatus = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetFilterString
' Description:       Get the string value from a filter "[???:string]"
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:55:48
' Parameters :       slFilter (String) - "[???:string]"
'--------------------------------------------------------------------------------
Function mGetFilterString(slFilter As String) As String
    Dim ilPos As Integer
    Dim ilPos2 As Integer
    ilPos = InStr(1, slFilter, "[")
    If ilPos = 0 Then Exit Function
    ilPos = InStr(ilPos + 1, slFilter, ":")
    If ilPos = 0 Then Exit Function
    
    ilPos2 = InStr(ilPos, slFilter, "]") - 1
    If ilPos2 = 0 Then Exit Function
    
    If ilPos > 0 And ilPos2 > 0 Then
        mGetFilterString = Mid(slFilter, ilPos + 1, (ilPos2 - (ilPos + 1)) + 1)
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetFilterValue
' Description:       Gets value from filter string, providing the filter type; returns a provided DefaultValue if not found
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:58:41
' Parameters :       slFilterName (String) - "Adv, Agy, etc.
'                    slDefaultValue (String = "") - the value to return if filter not found
'--------------------------------------------------------------------------------
Function mGetFilterValue(slFilterName As String, Optional slDefaultValue As String = "")
    Dim ilPos As Integer
    Dim ilPos2 As Integer
    mGetFilterValue = slDefaultValue
    If Trim(txtCriteria.Text) = "" Then Exit Function
    ilPos = InStr(1, txtCriteria.Text, "[" & slFilterName & ":")
    If ilPos = 0 Then Exit Function
    ilPos = ilPos + Len(slFilterName) + 2
    ilPos2 = InStr(ilPos, txtCriteria.Text, "]") - 1
    
    If ilPos > 0 And ilPos2 > 0 Then
        mGetFilterValue = Mid(txtCriteria.Text, ilPos, ilPos2 - ilPos + 1)
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetItemID
' Description:       Looks up the Item ID for the given String Name
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:40:52
' Parameters :       slType (String) - What Type (Like "Advertiser", "Agency", etc.
'                    slString (String) = The String Name of the item to get an ID for
'--------------------------------------------------------------------------------
Function mGetItemID(slType As String, slString As String) As Long
    Dim ilLoop As Long
    Select Case slType
        Case "Advertiser":
            For ilLoop = 0 To UBound(tgCommAdf)
                If Trim(tgCommAdf(ilLoop).sName) = Trim(slString) Then
                    mGetItemID = tgCommAdf(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
            
        Case "Agency":
            For ilLoop = 0 To UBound(tgCommAgf)
                If Trim(tgCommAgf(ilLoop).sName) & ", " & Trim$(tgCommAgf(ilLoop).sCityID) = Trim(slString) Then
                    mGetItemID = tgCommAgf(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
        
        Case "Salesperson":
            For ilLoop = 0 To UBound(tgMSlf)
                If Trim(tgMSlf(ilLoop).sLastName) & ", " & Trim(tgMSlf(ilLoop).sFirstName) = Trim(slString) Then
                    mGetItemID = tgMSlf(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop

        Case "Sales Office":
            For ilLoop = 0 To UBound(tmSofList)
                If Trim(tmSofList(ilLoop).sName) = Trim(slString) Then
                    mGetItemID = tmSofList(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
    End Select
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetSalesOffice
' Description:       Get Sales Office name for the provided ID
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:20:53
' Parameters :       ilSofCode (Integer)
'--------------------------------------------------------------------------------
Function mGetSalesOffice(ilSofCode As Integer) As String
    Dim ilLoopOnEntry As Integer
    Dim ilLoopTemp As Integer
    
    If ilSofCode = imLastSlpOffice Then
        mGetSalesOffice = smLastSlpOfficeName
        Exit Function
    End If
    For ilLoopOnEntry = 0 To UBound(tmSofList)
        'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 5
        If tmSofList(ilLoopOnEntry).iCode = imSlpOffice Then     'matching sales offices
            mGetSalesOffice = Trim(tmSofList(ilLoopOnEntry).sName)
            Exit For
        End If
    Next ilLoopOnEntry
    imLastSlpOffice = ilSofCode
    smLastSlpOfficeName = mGetSalesOffice
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetSalespersonName
' Description:       Get Salesperson Name (Last, First) for the provided ID
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:21:45
' Parameters :       ilNumber (Integer)
'--------------------------------------------------------------------------------
Function mGetSalespersonName(ilNumber As Integer) As String
    Dim ilInx As Integer
    Dim ilLoop As Integer
    
    If ilNumber = imLastSlp Then
        mGetSalespersonName = imLastSlpName
        Exit Function
    End If
    ilInx = gBinarySearchSlf(ilNumber)
    If ilInx >= 0 Then              '10-10-18
        mGetSalespersonName = Trim$(tgMSlf(ilInx).sLastName) & ", " & Trim$(tgMSlf(ilInx).sFirstName)
        imLastSlp = ilNumber
        imLastSlpName = mGetSalespersonName
        imSlpOffice = tgMSlf(ilInx).iSofCode
        'Get Sales person's User Group
        imLastSlpGroup = 0
        For ilLoop = 0 To UBound(tgPopUrf)
            If tgPopUrf(ilLoop).iSlfCode = tgMSlf(ilInx).iCode Then
                imLastSlpGroup = tgPopUrf(ilLoop).iGroupNo
                Exit For
            End If
        Next ilLoop
    Else
        mGetSalespersonName = "Unknown"
        imLastSlp = ilNumber
        imLastSlpName = mGetSalespersonName
        imSlpOffice = 0
        ilSlpURF = 0
        imLastSlpGroup = 0
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetSchdStatus
' Description:       Get Schedule Status description for provided code
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:18:10
' Parameters :       slSchdStatusCode (String)
'--------------------------------------------------------------------------------
Function mGetSchdStatus(slSchdStatusCode As String) As String
    Dim slString As String
    Select Case slSchdStatusCode
        Case "F": slString = "Fully Scheduled"
        Case "M": slString = "Manually Scheduled"
        Case "P": slString = "Prevent Scheduling"
        Case "I": slString = "Interrupted Scheduling"
        Case "N": slString = "New Contract"
        Case "A": slString = "Altered Contract"
    End Select
    mGetSchdStatus = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetLineType
' Description:       Gets the Line Type description for the provided ChfDetail record
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/2/2024-08:59:01
' Parameters :       tlChfDetail (CHFDETAILLIST)
'--------------------------------------------------------------------------------
Private Function mGetLineType(tlChfDetail As CHFDETAILLIST) As String
    Dim slString As String
    slString = IIF(Trim(tlChfDetail.sAirtimeDefined) = "", "N", tlChfDetail.sAirtimeDefined)
    slString = slString & IIF(Trim(tlChfDetail.sNTRDefined) = "", "N", tlChfDetail.sNTRDefined)
    slString = slString & IIF(Trim(tlChfDetail.sAdServerDefined) = "", "N", tlChfDetail.sAdServerDefined)
    If slString = "YNN" Or slString = "NYN" Or slString = "NNY" Then
        'Only one type
        If tlChfDetail.sAirtimeDefined = "Y" Then slString = "Air Time"
        If tlChfDetail.sNTRDefined = "Y" Then slString = "NTR"
        If tlChfDetail.sAdServerDefined = "Y" Then slString = "Digital"
    Else
        'Mixed types
        slString = ""
        If tlChfDetail.sAirtimeDefined = "Y" Then slString = slString & "Air"
        If tlChfDetail.sNTRDefined = "Y" Then
            If slString <> "" Then slString = slString & ","
            slString = slString & "NT"
        End If
        If tlChfDetail.sAdServerDefined = "Y" Then
            If slString <> "" Then slString = slString & ","
            slString = slString & "Dig"
        End If
    End If
    mGetLineType = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mLoadFilterSearch
' Description:       Shows matching string(s) from txtSearch in lstFilterSearch - from cboItems (Multi-term Auto type)
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/7/2024-17:15:33
' Parameters :
'--------------------------------------------------------------------------------
Sub mLoadFilterSearch()
    Dim slSearchParts() As String
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    Dim slSearchpart As String
    Dim ilFindParts As Integer
    
    lstFilterSearch.Clear
    slSearchParts = Split(txtSearch, " ")
    
    For ilLoop = 0 To cboItems.ListCount - 1
        ilFindParts = 0
        For ilLoop2 = 0 To UBound(slSearchParts)
            If InStr(1, LCase(cboItems.List(ilLoop)), LCase(slSearchParts(ilLoop2)), vbTextCompare) > 0 Then
                ilFindParts = ilFindParts + 1
            End If
        Next ilLoop2
        If ilFindParts = UBound(slSearchParts) + 1 Then
            lstFilterSearch.AddItem cboItems.List(ilLoop)
        End If
    Next ilLoop
    'Auto select if only 1 item remains in filter search results
    If lstFilterSearch.ListCount = 1 Then
        lstFilterSearch.Selected(0) = True
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mMakeFilterString
' Description:       makes a filter string from a Filter Type & value String
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:03:19
' Parameters :       slString (String = "") - Optional: If ommited, use txtSearch.Text
'                    slFilterType (String = "") - Optional: Advertiser, Agency, Etc.  If omitted use cboFilterType
'--------------------------------------------------------------------------------
Function mMakeFilterString(Optional slString As String = "", Optional slFilterType As String = "") As String
    Dim slFilterString As String
    If slString = "" Then slString = txtSearch.Text
    If slString = "" Then Exit Function
    If slFilterType = "" Then slFilterType = cboFilterType.Text
    If slFilterType = "" Then Exit Function
    'Makes a Filter string from the selection
    Select Case slFilterType
        Case "Advertiser"
            slFilterString = "[Adv:" & Trim(slString) & "]"
        
        Case "Agency"
            slFilterString = "[Agy:" & Trim(slString) & "]"
        
        Case "Active on or after"
            If mValidateDate(slString) Then
                slString = Format(DateValue(slString), "mm/dd/yy")
                slFilterString = "[*Active>:" & Trim(slString) & "]"
            End If
            
        Case "Active on or prior"
            If mValidateDate(slString) Then
                slString = Format(DateValue(slString), "mm/dd/yy")
                slFilterString = "[Active<:" & Trim(slString) & "]"
            End If
        
        Case "Update Date (Beginning)"
            If mValidateDate(slString) Then
                slString = Format(DateValue(slString), "mm/dd/yy")
                slFilterString = "[UDate>:" & Trim(slString) & "]"
            End If
        
        Case "Update Date (Ending)"
            If mValidateDate(slString) Then
                slString = Format(DateValue(slString), "mm/dd/yy")
                slFilterString = "[UDate<:" & Trim(slString) & "]"
            End If
        
        Case "Contract Number"
            slFilterString = "[Cntr:" & Trim(Val(slString)) & "]"
            
        Case "Contract Status"
            slFilterString = "[CStat:" & Trim(slString) & "]"
        
        Case "Contract Type"
            slFilterString = "[CType:" & Trim(slString) & "]"
            
        Case "Digital Delivery Status"
            slFilterString = "[DlvyStat:" & Trim(slString) & "]"
            
        Case "Schedule Status"
            slFilterString = "[SchStat:" & Trim(slString) & "]"
            
        Case "Product"
            slFilterString = "[Prod:" & Trim(slString) & "]"
            
        Case "Salesperson"
            slFilterString = "[SP:" & Trim(slString) & "]"
        
        Case "Sales Group"
            slFilterString = "[Group:" & Trim(slString) & "]"
        
        Case "Sales Office"
            slFilterString = "[SO:" & Trim(slString) & "]"

    End Select

    mMakeFilterString = slFilterString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mManageAppliedFilters
' Description:       Add, Update or Remove Filters in tmAppliedFilters
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:59:18
' Parameters :       slAction (String) - Add, Update, Remove or Delete
'                    slFilterType (String) - Adv, Agy, etc.
'                    llFilterValue (Long) - AdvertiserID, AgencyID, etc. 0 for Strings and Dates
'                    slFilterValue (String) - Advertiser Name, Agency Name, etc. Date String like mm/dd/yy
'                    slOldFilterValueName (String = "") - Optional: used for Update.  Used to find the old item to update
'--------------------------------------------------------------------------------
Sub mManageAppliedFilters(slAction As String, slFilterType As String, llFilterValue As Long, slFilterValue As String, Optional slOldFilterValueName As String = "")
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Select Case slAction
        Case "Add"
            'Look for an Empty slot
            ilIndex = -1
            For ilLoop = 0 To UBound(tmAppliedFilters)
                If tmAppliedFilters(ilLoop).sType = "" Then
                    ilIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            If ilIndex = -1 Then
                'Add new slot
                ReDim Preserve tmAppliedFilters(0 To UBound(tmAppliedFilters) + 1) As APPLIEDFILTER
                tmAppliedFilters(UBound(tmAppliedFilters)).sType = slFilterType
                tmAppliedFilters(UBound(tmAppliedFilters)).lValue = llFilterValue
                tmAppliedFilters(UBound(tmAppliedFilters)).sValue = slFilterValue
            Else
                'Update empty slot
                tmAppliedFilters(ilIndex).sType = slFilterType
                tmAppliedFilters(ilIndex).lValue = llFilterValue
                tmAppliedFilters(ilIndex).sValue = slFilterValue
            End If
            
        Case "Remove", "Delete"
            ilIndex = -1
            For ilLoop = 0 To UBound(tmAppliedFilters)
                If tmAppliedFilters(ilLoop).sType = slFilterType Then
                    If tmAppliedFilters(ilLoop).sValue = slFilterValue Then
                        ilIndex = ilLoop
                        Exit For
                    End If
                End If
            Next ilLoop
            If ilIndex = -1 Then
                MsgBox "Couldn't find Filter to Delete!"
            Else
                'Update empty slot
                tmAppliedFilters(ilIndex).sType = ""
                tmAppliedFilters(ilIndex).lValue = 0
                tmAppliedFilters(ilIndex).sValue = ""
            End If
            
        Case "Update"
            ilIndex = -1
            For ilLoop = 0 To UBound(tmAppliedFilters)
                If tmAppliedFilters(ilLoop).sType = slFilterType Then
                    If tmAppliedFilters(ilLoop).sValue = slOldFilterValueName Then
                        ilIndex = ilLoop
                        Exit For
                    End If
                End If
            Next ilLoop
            If ilIndex = -1 Then
                MsgBox "Couldn't find Filter to Update!"
            Else
                'Update empty slot
                tmAppliedFilters(ilIndex).sType = slFilterType
                tmAppliedFilters(ilIndex).lValue = llFilterValue
                tmAppliedFilters(ilIndex).sValue = slFilterValue
            End If
    End Select
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateAdvertiserList
' Description:       Populate the provided list with Advertiser names
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:24:52
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateAdvertiserList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smAdvertiserList) > 0 Then
        For llItem = 0 To UBound(smAdvertiserList) - 1
            olList.AddItem smAdvertiserList(llItem)
        Next llItem
    Else
        ilRet = gPopAdvtBox(DashboardVw, olList, tgAdvertiser(), sgAdvertiserTag)
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateAgencyList
' Description:       Populate the provided list with Agency names
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:24:24
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateAgencyList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smAgencyList) > 0 Then
        For llItem = 0 To UBound(smAgencyList) - 1
            olList.AddItem smAgencyList(llItem)
        Next llItem
    Else
        ilRet = gPopAgyBox(DashboardVw, olList, tgAgency(), sgAgencyTag)
        cboItems.AddItem "N/A"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateContractStatusList
' Description:       Populate provided list with Contract Status descriptions
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:11:23
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateContractStatusList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smContractStatusList) > 0 Then
        For llItem = 0 To UBound(smContractStatusList) - 1
            olList.AddItem smContractStatusList(llItem)
        Next llItem
    Else
        olList.AddItem "Working Proposal"
        olList.AddItem "Rejected"
        olList.AddItem "Completed Proposal"
        olList.AddItem "Unapproved Proposal"
        olList.AddItem "Approved Hold"
        olList.AddItem "Hold"
        olList.AddItem "Approved Order"
        olList.AddItem "Order"
        olList.AddItem "Rev Completed"
        olList.AddItem "Rev Working"
        olList.AddItem "Rev Unapproved"
        olList.AddItem "Unapproved"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateContractTypeList
' Description:       Populate provided list with Contract Type descriptions
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:12:00
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateContractTypeList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smContractTypeList) > 0 Then
        For llItem = 0 To UBound(smContractTypeList) - 1
            olList.AddItem smContractTypeList(llItem)
        Next llItem
    Else
        olList.AddItem "Standard"
        olList.AddItem "Reservation"
        olList.AddItem "Remnant"
        olList.AddItem "Direct Response"
        olList.AddItem "Per Inquiry"
        olList.AddItem "PSA"
        olList.AddItem "Promo"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateFilterType
' Description:       Populates cboFilterType with the types of filters available
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:25:44
' Parameters :
'--------------------------------------------------------------------------------
Sub mPopulateFilterType()
    cboFilterType.Clear
    cboFilterType.AddItem "Advertiser"
    cboFilterType.AddItem "Agency"
    cboFilterType.AddItem "Active on or after"
    cboFilterType.AddItem "Active on or prior"
    cboFilterType.AddItem "Update Date (Beginning)"
    cboFilterType.AddItem "Update Date (Ending)"
    cboFilterType.AddItem "Contract Number"
    cboFilterType.AddItem "Contract Type"
    cboFilterType.AddItem "Contract Status"
    cboFilterType.AddItem "Digital Delivery Status"
    cboFilterType.AddItem "Schedule Status"
    cboFilterType.AddItem "Product"
    cboFilterType.AddItem "Salesperson"
    cboFilterType.AddItem "Sales Office"
    
    'if Not a salesperson; which could be a Manager or Planner or Negotiator (AND if they have a Group), then Add Sales Group filter
    If smCurrentTitle <> "S" And imCurrentGroup > 0 Then
        cboFilterType.AddItem "Sales Group"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateSalesOfficeList
' Description:       Populate provided list with Sales Office names
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:21:14
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateSalesOfficeList(olList As ComboBox)
    Dim llItem As Long
    Dim ilLoop
    olList.Clear
    If UBound(smSalesOfficeList) > 0 Then
        For llItem = 0 To UBound(smSalesOfficeList) - 1
            olList.AddItem smSalesOfficeList(llItem)
        Next llItem
    Else
        For ilLoop = 0 To UBound(tmSofList)
            olList.AddItem Trim(tmSofList(ilLoop).sName)
        Next ilLoop
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateSalespersonList
' Description:       Populate the provided list with Salesperson Names
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:22:47
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateSalespersonList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smSalespersonList) > 0 Then
        For llItem = 0 To UBound(smSalespersonList) - 1
            olList.AddItem smSalespersonList(llItem)
        Next llItem
    Else
        ilRet = gPopSalespersonBox(DashboardVw, 0, True, True, olList, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateScheduleStatusList
' Description:       Populate provided list with Schedule Status descriptions
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:20:09
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateScheduleStatusList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smScheduleStatusList) > 0 Then
        For llItem = 0 To UBound(smScheduleStatusList) - 1
            olList.AddItem smScheduleStatusList(llItem)
        Next llItem
    Else
        olList.AddItem "Fully Scheduled"
        olList.AddItem "Manually Scheduled"
        olList.AddItem "Prevent Scheduling"
        olList.AddItem "Interrupted Scheduling"
        olList.AddItem "New Contract"
        olList.AddItem "Altered Contract"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulateDigitalDlvyStatusList
' Description:       Populate provided list with Digital Delivery Status descriptions
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       3/7/2024-8:00:00
' Parameters :       olList (ComboBox)
'--------------------------------------------------------------------------------
Sub mPopulateDigitalDlvyStatusList(olList As ComboBox)
    Dim llItem As Long
    olList.Clear
    If UBound(smDeliveryStatusList) > 0 Then
        For llItem = 0 To UBound(smDeliveryStatusList) - 1
            olList.AddItem smDeliveryStatusList(llItem)
        Next llItem
    Else
        olList.AddItem "N/A"
        olList.AddItem "Not pushed"
        olList.AddItem "Pushed"
        olList.AddItem "Partial"
        olList.AddItem "Issue Encountered"
        olList.AddItem "Requires Repush"
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSelectFilterType
' Description:       Selects the filter cboFilterType if matching the provided filterType string
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:07:06
' Parameters :       slFilterType (String)
'--------------------------------------------------------------------------------
Sub mSelectFilterType(slFilterType As String)
    Dim ilLoop As Integer
    For ilLoop = 0 To cboFilterType.ListCount - 1
        If cboFilterType.List(ilLoop) = slFilterType Then
            cboFilterType.ListIndex = ilLoop
            Exit Sub
        End If
    Next ilLoop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSortByColumn
' Description:       Applies column sorting to specified column.  Keeps track of Ascending/Descending
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-09:09:39
' Parameters :       ilColNum (Integer) - Sort which column?
'                    blSkipRev (Boolean = False) - Skip Reversing sort order (used when calling sort after a populate)
'--------------------------------------------------------------------------------
Sub mSortByColumn(ilColNum As Integer, Optional blSkipRev As Boolean = False)
    lblItemCount.Caption = "Applying Sort..."
    grdCntr.Row = 0
    grdCntr.RowSel = 0
    grdCntr.Redraw = False
    If imSortDir = 0 Then imSortDir = flexSortGenericAscending
    
    If imLastSortCol = ilColNum Then
        'reverse sort the same column
        If blSkipRev = False Then
            If imSortDir = flexSortGenericDescending Then
                imSortDir = flexSortGenericAscending
            Else
                imSortDir = flexSortGenericDescending
            End If
        End If
    Else
        imSortDir = flexSortGenericAscending
        imLastSortCol = ilColNum
    End If
    
    'set which column to sort. If a "Special" column (A date or number), fill C_SORTINDEX with a sortable value, then sort by C_SORTINDEX
    ilSortCol = C_SORTINDEX
    Select Case ilColNum
        Case C_CNTRNOINDEX:
            If imLastFilledCol <> ilColNum Or blSkipRev Then
                mFillSortColumn ilColNum, "Number", Not blSkipRev: imLastFilledCol = ilColNum
            End If
        Case C_STARTDATEINDEX:
            If imLastFilledCol <> ilColNum Or blSkipRev Then
                mFillSortColumn ilColNum, "Date", Not blSkipRev: imLastFilledCol = ilColNum
            End If
        Case C_ENDDATEINDEX:
            If imLastFilledCol <> ilColNum Or blSkipRev Then
                mFillSortColumn ilColNum, "Date", Not blSkipRev: imLastFilledCol = ilColNum
            End If
        Case C_CNTRUPDATEDATEINDEX:
            If imLastFilledCol <> ilColNum Or blSkipRev Then
                mFillSortColumn ilColNum, "Date", Not blSkipRev: imLastFilledCol = ilColNum
            End If
        Case C_GROSSINDEX:
            If imLastFilledCol <> ilColNum Or blSkipRev Then
                mFillSortColumn ilColNum, "Number", Not blSkipRev: imLastFilledCol = ilColNum
            End If
        Case Else: ilSortCol = ilColNum
    End Select
    ProgressBar1.Value = 100
    
    'Sort
    grdCntr.ColSel = ilSortCol
    grdCntr.Col = ilSortCol
    grdCntr.Sort = imSortDir
    grdCntr.Redraw = True
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mValidateDate
' Description:       make sure we have a good date in the provided string
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-12:02:06
' Parameters :       slText (String)
'--------------------------------------------------------------------------------
Function mValidateDate(slText As String) As Boolean
    Dim ilSlashPos1 As Integer
    Dim ilSlashPos2 As Integer
    Dim ilSelPos As Integer
    Dim ilSelLen As Integer
    ilSelPos = txtSearch.SelStart
    ilSelLen = txtSearch.SelLength
    
    ilSlashPos1 = InStr(1, slText, "/")
    If ilSlashPos1 = 0 Then Exit Function
    
    ilSlashPos2 = InStr(ilSlashPos1 + 1, slText, "/")
    If ilSlashPos2 = 0 Then Exit Function
    
    'Make sure at least 2 digits are entered for date
    If Len(slText) - 1 = ilSlashPos2 Then Exit Function
    If Not IsDate(slText) Then Exit Function
    
    If year(DateValue(slText)) < 1970 Then Exit Function
    If year(DateValue(slText)) > 2069 Then Exit Function

    mValidateDate = True
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mAddContractToGrid
' Description:       Adds provided Contract Detail to grdCntr
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:51:47
' Parameters :       tlChfDetail (CHFDETAILLIST)
'--------------------------------------------------------------------------------
Private Sub mAddContractToGrid(tlChfDetail As CHFDETAILLIST)
    Dim llRow As Long
    Dim ilString As String
    Dim llLong As Long
    Dim ilLoop As Integer
    Dim dlFontSize As Double
    
    Select Case CboFontSize.ListIndex
        Case 0: dlFontSize = 9
        Case 1: dlFontSize = 8.25
        Case 2: dlFontSize = 7
        Case 3: dlFontSize = 6
    End Select
    
    'Current Row
    llRow = grdCntr.rows - 1
    
    'Grid Font size
    If llRow = 1 Then
        grdCntr.Row = llRow
        grdCntr.Col = 0
        grdCntr.ColSel = grdCntr.cols - 1
        grdCntr.CellFontSize = dlFontSize
        grdCntr.FillStyle = flexFillRepeat
        grdCntr.Font.Name = "Arial"
    End If
    
    'Check If we need to add a new row
    If grdCntr.TextMatrix(llRow, C_CNTRNOINDEX) <> "" Then
        'Grid Font size

        grdCntr.AddItem ""
        llRow = llRow + 1
        
        grdCntr.Row = llRow
        grdCntr.Col = 0
        If grdCntr.CellFontSize <> dlFontSize Then
            grdCntr.ColSel = grdCntr.cols - 1
            grdCntr.CellFontSize = dlFontSize
            grdCntr.FillStyle = flexFillRepeat
            grdCntr.Font.Name = "Arial"
        End If
                
    End If
    
    'CHFCODE: C_CHFCODEINDEX
    grdCntr.TextMatrix(llRow, C_CHFCODEINDEX) = tlChfDetail.lCode
    
    'Last Update: C_CNTRUPDATEDATEINDEX
    gUnpackDateLong tlChfDetail.iOHDDate(0), tlChfDetail.iOHDDate(1), llLong
    grdCntr.TextMatrix(llRow, C_CNTRUPDATEDATEINDEX) = Format(llLong, "mm/dd/yy")
    
    'CntrNo: C_CNTRNOINDEX
    slString = tlChfDetail.lCntrNo
    'Fix Issue #12; v81 Contract Dashboard 2-14
    Select Case tlChfDetail.sStatus
        Case "W": slString = slString & IIF(tlChfDetail.iPropVer > 0, " V" & Format(tlChfDetail.iPropVer, "00"), "")
        Case "D": slString = slString & IIF(tlChfDetail.iPropVer > 0, " V" & Format(tlChfDetail.iPropVer, "00"), "")
        Case "C": slString = slString & IIF(tlChfDetail.iPropVer > 0, " V" & Format(tlChfDetail.iPropVer, "00"), "")
        Case "I": slString = slString & IIF(tlChfDetail.iPropVer > 0, " V" & Format(tlChfDetail.iPropVer, "00"), "")
        'Case "G": slString = "Approved Hold"
        'Case "H": slString = "Hold"
        'Case "N": slString = "Approved Order"
        'Case "O": slString = "Order"
    End Select
    grdCntr.TextMatrix(llRow, C_CNTRNOINDEX) = slString
    
    'Cntr Type: C_CNTRTYPEINDEX
    grdCntr.TextMatrix(llRow, C_CNTRTYPEINDEX) = Trim(mGetCntrType(tlChfDetail.sType))
    
    'Line Type: C_LINETYPEINDEX
    grdCntr.TextMatrix(llRow, C_LINETYPEINDEX) = Trim(mGetLineType(tlChfDetail))
    
    'Agency: C_AGYNAMEINDEX
    grdCntr.TextMatrix(llRow, C_AGYNAMEINDEX) = Trim(mGetAgencyName(tlChfDetail.iAgfCode))
        
    'Advertiser: C_ADVNAMEINDEX
    grdCntr.TextMatrix(llRow, C_ADVNAMEINDEX) = Trim(mGetAdvertiserName(tlChfDetail.iAdfCode))
    
    'Product: C_PRODUCTINDEX
    grdCntr.TextMatrix(llRow, C_PRODUCTINDEX) = Trim(tlChfDetail.sProduct)
    
    'Start Date: C_STARTDATEINDEX
    gUnpackDateLong tlChfDetail.iStartDate(0), tlChfDetail.iStartDate(1), llLong
    grdCntr.TextMatrix(llRow, C_STARTDATEINDEX) = Format(llLong, "mm/dd/yy")
    
    'End Date: C_ENDDATEINDEX
    gUnpackDateLong tlChfDetail.iEndDate(0), tlChfDetail.iEndDate(1), llLong
    grdCntr.TextMatrix(llRow, C_ENDDATEINDEX) = Format(llLong, "mm/dd/yy")
    
    'Gross: C_GROSSINDEX
    grdCntr.TextMatrix(llRow, C_GROSSINDEX) = Format(tlChfDetail.lInputGross / 100, "#,##0.00")
    
    'Salesperson: C_SALEPERSONINDEX
    grdCntr.TextMatrix(llRow, C_SALEPERSONINDEX) = Trim(mGetSalespersonName(tlChfDetail.iSlfCode(0)))
    
'    'Salesperson Group #
'    grdCntr.TextMatrix(llRow, C_GROUPINDEX) = imLastSlpGroup
    
'    'All Salesperson ID's
'    slString = ","
'    For ilLoop = 0 To 9
'        If tlChfDetail.iSlfCode(ilLoop) <> 0 Then
'            slString = slString & tlChfDetail.iSlfCode(ilLoop) & ","
'        End If
'    Next ilLoop
'    grdCntr.TextMatrix(llRow, C_SPNINDEX) = slString 'SalespersonID's
    
    'Sales Office: C_SALESOFFICEINDEX
    grdCntr.TextMatrix(llRow, C_SALESOFFICEINDEX) = mGetSalesOffice(imSlpOffice)
    
    'Contract Status: C_CNTRSTATUSINDEX
    grdCntr.TextMatrix(llRow, C_CNTRSTATUSINDEX) = mGetCntrStatus(tlChfDetail)
    
    'Digital Delivery Status: C_DIGITALDLVYINDEX
    grdCntr.TextMatrix(llRow, C_DIGITALDLVYINDEX) = mGetDeliveryStatus(tlChfDetail.iDelvyStatus)
    
    'Schedule Status: C_CNTRSCHEDULESTATUSINDEX
    grdCntr.TextMatrix(llRow, C_CNTRSCHEDULESTATUSINDEX) = mGetSchdStatus(tlChfDetail.sSchStatus)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mApplyFilter
' Description:       Read through tmChfDetailList array and check filters.  if Okay, add to grdCntr
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:53:50
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mApplyFilter()
    Dim llContractLoop As Long
    Dim ilAppliedFiltersLoop As Integer
    Dim ilLoop As Integer
    Dim llLong As Long
    Dim ilFoundShowFilter1 As Integer
    Dim ilFoundShowFilter2 As Integer
    Dim ilFoundADV As Integer
    Dim ilFoundAgy As Integer
    Dim ilFoundActiveStart As Integer
    Dim ilFoundActiveEnd As Integer
    Dim ilFoundUpdStart As Integer
    Dim ilFoundUpdEnd As Integer
    Dim ilFoundCntr As Integer
    Dim ilFoundCType As Integer
    Dim ilFoundCStat As Integer
    Dim ilFoundDlvyStat As Integer
    Dim ilFoundSchStat As Integer
    Dim ilFoundProd As Integer
    Dim ilFoundSP As Integer
    Dim ilFoundGroup As Integer
    Dim ilFoundSO As Integer
    Dim ilIncludeAirTime As Integer
    Dim ilIncludeNTR As Integer
    Dim ilIncludeDigital As Integer
    Dim ilIncludeProposals As Integer
    Dim ilIncludeContracts As Integer
    lblMultiVersionTotalWarn.Visible = False
    If ckcCntrType(3).Value = vbChecked Then lblMultiVersionTotalWarn.Visible = True
    
    cmcClose.Enabled = False
    cmcViewCntr.Enabled = False
    cmcChgCntr.Enabled = False
    cmcSchedule.Enabled = False
    
    Dim ilProgress As Integer
    Dim dlProgressStep As Double
    Dim ilDOE As Integer
    If UBound(tmChfDetailList) > 0 Then
        dlProgressStep = 75 / UBound(tmChfDetailList)
    End If
    
    lblItemCount.Caption = "Applying Filter..."
    mClearCntrGrid
    grdCntr.Redraw = False
    
    lmItemCount = 0
    dmTotalAmount = 0
    '------------------------------------------------
    'Loop through all loaded contracts
    For llContractLoop = 0 To UBound(tmChfDetailList) - 1
        'Progress bar
        If imCancelLoading = True Then GoTo CancelLoading
        ilDOE = ilDOE + 1
        If ilDOE >= 1000 Then
            dlProgress = dlProgressStep * llContractLoop
            ProgressBar1.Value = 20 + dlProgress
            ilDOE = 0
            grdCntr.Redraw = True: DoEvents: grdCntr.Redraw = False
        End If
        
        '------------------------------------------------
        'Determine if this contract should show on Grid
        'ilFound = 0
        'If these have No filters applied, then set them as Found=true, so they will pass the Add to grid criteria
        ilFoundADV = mCheckIfNoFilter("Adv")
        ilFoundAgy = mCheckIfNoFilter("Agy")
        ilFoundActiveStart = mCheckIfNoFilter("*Active>")
        ilFoundActiveEnd = mCheckIfNoFilter("Active<")
        ilFoundUpdStart = mCheckIfNoFilter("UDate>")
        ilFoundUpdEnd = mCheckIfNoFilter("UDate<")
        ilFoundCntr = mCheckIfNoFilter("Cntr")
        ilFoundCType = mCheckIfNoFilter("CType")
        ilFoundCStat = mCheckIfNoFilter("CStat")
        ilFoundDlvyStat = mCheckIfNoFilter("DlvyStat")
        ilFoundSchStat = mCheckIfNoFilter("SchStat")
        ilFoundProd = mCheckIfNoFilter("Prod")
        ilFoundSP = mCheckIfNoFilter("SP")
        ilFoundGroup = mCheckIfNoFilter("Group")
        ilFoundSO = mCheckIfNoFilter("SO")
        
        '------------------------------
        'AirTime, NTR, Digital
        ilFoundShowFilter1 = False
        'Contract/Proposal
        ilFoundShowFilter2 = False
        
        On Error GoTo CancelLoading
        'Checkbox Filters
        'AirTime
        If ckcCntrType(0).Value = vbChecked Then
            If tmChfDetailList(llContractLoop).sAirtimeDefined = "Y" Then ilFoundShowFilter1 = True
        End If
        
        'NTR
        If ckcCntrType(1).Value = vbChecked Then
            If tmChfDetailList(llContractLoop).sNTRDefined = "Y" Then ilFoundShowFilter1 = True
        End If
        
        'Digital
        If ckcCntrType(2).Value = vbChecked Then
            If tmChfDetailList(llContractLoop).sAdServerDefined = "Y" Then ilFoundShowFilter1 = True
        End If
        
        '------------------------------
        'Proposals
        If ckcCntrType(3).Value = vbChecked Then
            'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 7
            'If tmChfDetailList(llContractLoop).sStatus = "W" Or tmChfDetailList(llContractLoop).sStatus = "C" Or tmChfDetailList(llContractLoop).sStatus = "I" Then
            If tmChfDetailList(llContractLoop).sStatus = "W" Or tmChfDetailList(llContractLoop).sStatus = "C" Or tmChfDetailList(llContractLoop).sStatus = "I" Or tmChfDetailList(llContractLoop).sStatus = "D" Then
                ilFoundShowFilter2 = True
            End If
        End If
        
        'Contracts
        If ckcCntrType(4).Value = vbChecked Then
            If tmChfDetailList(llContractLoop).sStatus = "H" Or tmChfDetailList(llContractLoop).sStatus = "O" Or tmChfDetailList(llContractLoop).sStatus = "G" Or tmChfDetailList(llContractLoop).sStatus = "N" Then
                'H=Hold, O=Order, G=Approved Hold, N=Approved Order
                ilFoundShowFilter2 = True
            End If
            'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 2
            If tmChfDetailList(llContractLoop).iCntRevNo > 0 And (tmChfDetailList(llContractLoop).sStatus = "C" Or tmChfDetailList(llContractLoop).sStatus = "I" Or tmChfDetailList(llContractLoop).sStatus = "W") Then
                'CntRevNo > 0: C=Rev Completed, I=Rev Unapproved, W=Rev Working
                ilFoundShowFilter2 = True
            End If
        End If
        
        'Check all applied filters
        If ilFoundShowFilter1 And ilFoundShowFilter2 Then
            For ilAppliedFiltersLoop = 0 To UBound(tmAppliedFilters)
                Select Case tmAppliedFilters(ilAppliedFiltersLoop).sType
                    Case "SP" 'Salesperson
                        For ilLoop = 0 To 9
                            If tmChfDetailList(llContractLoop).iSlfCode(ilLoop) = tmAppliedFilters(ilAppliedFiltersLoop).lValue Then
                                ilFoundSP = True
                                Exit For
                            End If
                        Next ilLoop
                        
                    Case "Group" 'Sales group
                        'Load Contracts for other salespeople in same group if logged in Salesperson is a manager (issue 10)
                        If smCurrentTitle <> "S" And imCurrentGroup > 0 Then
                            If mDetermineMatchingSPGroup(tmChfDetailList(llContractLoop), imCurrentGroup) Then
                                ilFoundGroup = True
                                ilFoundSP = True
                            End If
                        End If
                    
                    Case "Adv" 'Advertiser
                        If tmChfDetailList(llContractLoop).iAdfCode = tmAppliedFilters(ilAppliedFiltersLoop).lValue Then
                            ilFoundADV = True
                        End If
                        
                    Case "Agy" 'Agency
                        If tmChfDetailList(llContractLoop).iAgfCode = tmAppliedFilters(ilAppliedFiltersLoop).lValue Then
                            ilFoundAgy = True
                        End If
                        
                    Case "UDate>" 'Update Date (Beginning)
                        gUnpackDateLong tmChfDetailList(llContractLoop).iOHDDate(0), tmChfDetailList(llContractLoop).iOHDDate(1), llLong
                        If DateValue(Format(llLong, "mm/dd/yy")) >= DateValue(tmAppliedFilters(ilAppliedFiltersLoop).sValue) Then
                            ilFoundUpdStart = True
                        End If
                    
                    Case "UDate<" 'Update Date (Ending)
                        gUnpackDateLong tmChfDetailList(llContractLoop).iOHDDate(0), tmChfDetailList(llContractLoop).iOHDDate(1), llLong
                        If DateValue(Format(llLong, "mm/dd/yy")) <= DateValue(tmAppliedFilters(ilAppliedFiltersLoop).sValue) Then
                            ilFoundUpdEnd = True
                        End If
                    
                    Case "Cntr" 'Contract No
                        If tmChfDetailList(llContractLoop).lCntrNo = tmAppliedFilters(ilAppliedFiltersLoop).lValue Then
                            ilFoundCntr = True
                        End If
                        
                    Case "CType" 'Contract Type
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Standard" And tmChfDetailList(llContractLoop).sType = "C" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Reservation" And tmChfDetailList(llContractLoop).sType = "V" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Remnant" And tmChfDetailList(llContractLoop).sType = "T" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Direct Response" And tmChfDetailList(llContractLoop).sType = "R" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Per Inquiry" And tmChfDetailList(llContractLoop).sType = "Q" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "PSA" And tmChfDetailList(llContractLoop).sType = "S" Then ilFoundCType = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Promo" And tmChfDetailList(llContractLoop).sType = "M" Then ilFoundCType = True
                        
                    Case "CStat" 'Contract Status
                        'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 3
                        'Approved Hold          G
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Approved Hold" And tmChfDetailList(llContractLoop).sStatus = "G" Then ilFoundCStat = True
                        'Hold                   H
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Hold" And tmChfDetailList(llContractLoop).sStatus = "H" Then ilFoundCStat = True
                        'Approved Order         N
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Approved Order" And tmChfDetailList(llContractLoop).sStatus = "N" Then ilFoundCStat = True
                        'Order                  O
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Order" And tmChfDetailList(llContractLoop).sStatus = "O" Then ilFoundCStat = True
                        'Rev Completed          C   > 0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Rev Completed" And tmChfDetailList(llContractLoop).sStatus = "C" And tmChfDetailList(llContractLoop).iCntRevNo > 0 Then ilFoundCStat = True
                        'Rev Unapproved         I   > 0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Rev Unapproved" And tmChfDetailList(llContractLoop).sStatus = "I" And tmChfDetailList(llContractLoop).iCntRevNo > 0 Then ilFoundCStat = True
                        'Rev Working            W   > 0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Rev Working" And tmChfDetailList(llContractLoop).sStatus = "W" And tmChfDetailList(llContractLoop).iCntRevNo > 0 Then ilFoundCStat = True
                        'Completed Proposal     C   0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Completed Proposal" And tmChfDetailList(llContractLoop).sStatus = "C" And tmChfDetailList(llContractLoop).iCntRevNo = 0 Then ilFoundCStat = True
                        'Rejected               D
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Rejected" And tmChfDetailList(llContractLoop).sStatus = "D" Then ilFoundCStat = True
                        'Unapproved Proposal    I   0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Unapproved Proposal" And tmChfDetailList(llContractLoop).sStatus = "I" And tmChfDetailList(llContractLoop).iCntRevNo = 0 Then ilFoundCStat = True
                        'Working Proposal       W   0
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Working Proposal" And tmChfDetailList(llContractLoop).sStatus = "W" And tmChfDetailList(llContractLoop).iCntRevNo = 0 Then ilFoundCStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Unapproved" And tmChfDetailList(llContractLoop).sStatus = "N" Then ilFoundCStat = True
                        
                    Case "DlvyStat" 'Delivery Status
                        '0=N/A
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "N/A" And tmChfDetailList(llContractLoop).iDelvyStatus = 0 Then ilFoundDlvyStat = True
                        '1=not pushed
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Not pushed" And tmChfDetailList(llContractLoop).iDelvyStatus = 1 Then ilFoundDlvyStat = True
                        '2=pushed
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Pushed" And tmChfDetailList(llContractLoop).iDelvyStatus = 2 Then ilFoundDlvyStat = True
                        '3=partial
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Partial" And tmChfDetailList(llContractLoop).iDelvyStatus = 3 Then ilFoundDlvyStat = True
                        '4=Issue Encountered
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Issue Encountered" And tmChfDetailList(llContractLoop).iDelvyStatus = 4 Then ilFoundDlvyStat = True
                        '5=Requires Repush
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Requires Repush" And tmChfDetailList(llContractLoop).iDelvyStatus = 5 Then ilFoundDlvyStat = True
                        
                    Case "SchStat" 'Schedule Status
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Fully Scheduled" And tmChfDetailList(llContractLoop).sSchStatus = "F" Then ilFoundSchStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Manually Scheduled" And tmChfDetailList(llContractLoop).sSchStatus = "M" Then ilFoundSchStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Prevent Scheduling" And tmChfDetailList(llContractLoop).sSchStatus = "P" Then ilFoundSchStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Interrupted Scheduling" And tmChfDetailList(llContractLoop).sSchStatus = "I" Then ilFoundSchStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "New Contract" And tmChfDetailList(llContractLoop).sSchStatus = "N" Then ilFoundSchStat = True
                        If tmAppliedFilters(ilAppliedFiltersLoop).sValue = "Altered Contract" And tmChfDetailList(llContractLoop).sSchStatus = "A" Then ilFoundSchStat = True
                    
                    Case "Prod" 'Product
                        If InStr(1, LCase(tmChfDetailList(llContractLoop).sProduct), LCase(tmAppliedFilters(ilAppliedFiltersLoop).sValue)) > 0 Then
                            ilFoundProd = True
                        End If
                        
                    Case "SO" 'Sales Office
                        If mGetSPOffice(tmChfDetailList(llContractLoop).iSlfCode(0)) = tmAppliedFilters(ilAppliedFiltersLoop).lValue Then
                            ilFoundSO = True
                        End If
                End Select
            Next ilAppliedFiltersLoop
        End If
        
        '------------------------------------------------
        If bmBuildingFilteredResults Then
            mBuildFilters tmChfDetailList(llContractLoop)
        End If
        
        '------------------------------------------------
        'Add this contract to grid (we found a match for ALL filters)
        If ilFoundShowFilter1 And ilFoundShowFilter2 Then
            If ilFoundADV And ilFoundAgy And ilFoundUpdStart And ilFoundUpdEnd And ilFoundCntr And ilFoundCType And ilFoundCStat And ilFoundDlvyStat And ilFoundSchStat And ilFoundProd And ilFoundSP And ilFoundGroup And ilFoundSO Then
                lmItemCount = lmItemCount + 1
                dmTotalAmount = dmTotalAmount + tmChfDetailList(llContractLoop).lInputGross / 100
                mAddContractToGrid tmChfDetailList(llContractLoop)
                If lmItemCount = 50 Then grdCntr.Row = 2: grdCntr.RowSel = 2: grdCntr.Redraw = True: DoEvents: grdCntr.Redraw = False: grdCntr.MousePointer = flexArrowHourGlass: Screen.MousePointer = vbArrowHourglass
            End If
        End If
    Next llContractLoop
    
CancelLoading:
    Screen.MousePointer = vbDefault
    grdCntr.MousePointer = flexArrow
    
    ProgressBar1.Value = 95
    mSortByColumn imLastSortCol, True
    
    If imCancelLoading Then
        lblItemCount.Caption = "Canceled"
        lmItemCount = 0
        mClearCntrGrid
    Else
        lblItemCount.Caption = "Showing " & Format(lmItemCount, "#,##0") & " Items / " & Format(dmTotalAmount, "$#,##0.00") & " Gross"
    End If
    grdCntr.Redraw = True
    bmBuildingFilteredResults = False
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mGetSPOffice
' Description:       Get's a provided Salesperson's (ilSlfCode) Office Code
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/13/2024-09:01:05
' Parameters :       ilSlfCode (Integer)
'--------------------------------------------------------------------------------
Private Function mGetSPOffice(ilSlfCode As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilInx As Integer
    If ilSlfCode = imLastGetSPOfficeSlf Then
        mGetSPOffice = imLastGetSPOffice
        Exit Function
    End If
    
    ilInx = gBinarySearchSlf(ilSlfCode)
    If ilInx >= 0 Then              '10-10-18
        mGetSPOffice = tgMSlf(ilInx).iSofCode
        imLastGetSPOffice = mGetSPOffice
        imLastGetSPOfficeSlf = ilSlfCode
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mDetermineMatchingSPGroup
' Description:       Looks up the order's Salespeople and checks each one to see if Group matches provided ilGroupNumber
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/13/2024-08:59:56
' Parameters :       tlChfDetail (CHFDETAILLIST)
'                    ilGroupNumber (Integer)
'--------------------------------------------------------------------------------
Private Function mDetermineMatchingSPGroup(tlChfDetail As CHFDETAILLIST, ilGroupNumber As Integer) As Boolean
    Dim ilFound As Integer
    Dim ilSPLoop As Integer
    Dim ilUserLoop As Integer
    ilFound = False
    'look up each Salesperson in the contract
    For ilSPLoop = 0 To 9
        'Get Sales person's User Group
        For ilUserLoop = 0 To UBound(tgPopUrf) - 1
            If tgPopUrf(ilUserLoop).iSlfCode = tlChfDetail.iSlfCode(ilSPLoop) Then
                If tgPopUrf(ilUserLoop).iGroupNo = ilGroupNumber Then
                    ilFound = True
                    Exit For
                End If
            End If
        Next ilUserLoop
        If ilFound = True Then Exit For
    Next ilSPLoop
    mDetermineMatchingSPGroup = ilFound
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mBuildFilters
' Description:       Populates arrays of Filter lists (Advertiser, Agencies, etc.) From a given Contract Detail record
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:27:50
' Parameters :       tlChfDetail (CHFDETAILLIST)
'--------------------------------------------------------------------------------
Private Sub mBuildFilters(tlChfDetail As CHFDETAILLIST)
    Dim llLoop As Long
    Dim ilFound As Integer
    Dim slString As String
    
    'Advertiser
    slString = Trim(mGetAdvertiserName(tlChfDetail.iAdfCode))
    ilFound = 0
    For llLoop = 0 To UBound(smAdvertiserList) - 1
        If smAdvertiserList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smAdvertiserList(UBound(smAdvertiserList)) = slString
        ReDim Preserve smAdvertiserList(0 To UBound(smAdvertiserList) + 1)
    End If
    
    'Agency
    slString = Trim(mGetAgencyName(tlChfDetail.iAgfCode))
    ilFound = 0
    For llLoop = 0 To UBound(smAgencyList)
        If smAgencyList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smAgencyList(UBound(smAgencyList)) = slString
        ReDim Preserve smAgencyList(0 To UBound(smAgencyList) + 1)
    End If
    
    'Sales Office
    slString = Trim(mGetSalesOffice(imSlpOffice))
    ilFound = 0
    For llLoop = 0 To UBound(smSalesOfficeList)
        If smSalesOfficeList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smSalesOfficeList(UBound(smSalesOfficeList)) = slString
        ReDim Preserve smSalesOfficeList(0 To UBound(smSalesOfficeList) + 1)
    End If
    
    'Sales Person
    slString = Trim(mGetSalespersonName(tlChfDetail.iSlfCode(0)))
    ilFound = 0
    For llLoop = 0 To UBound(smSalespersonList)
        If smSalespersonList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smSalespersonList(UBound(smSalespersonList)) = slString
        ReDim Preserve smSalespersonList(0 To UBound(smSalespersonList) + 1)
    End If
    
    'Contract Status
    slString = Trim(mGetCntrStatus(tlChfDetail))
    ilFound = 0
    For llLoop = 0 To UBound(smContractStatusList)
        If smContractStatusList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smContractStatusList(UBound(smContractStatusList)) = slString
        ReDim Preserve smContractStatusList(0 To UBound(smContractStatusList) + 1)
    End If
    
    'Contract Type
    slString = Trim(mGetCntrType(tlChfDetail.sType))
    ilFound = 0
    For llLoop = 0 To UBound(smContractTypeList)
        If smContractTypeList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smContractTypeList(UBound(smContractTypeList)) = slString
        ReDim Preserve smContractTypeList(0 To UBound(smContractTypeList) + 1)
    End If
    
    'Digital Delivery Status
    slString = Trim(mGetDeliveryStatus(tlChfDetail.iDelvyStatus))
    ilFound = 0
    For llLoop = 0 To UBound(smDeliveryStatusList)
        If smDeliveryStatusList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smDeliveryStatusList(UBound(smDeliveryStatusList)) = slString
        ReDim Preserve smDeliveryStatusList(0 To UBound(smDeliveryStatusList) + 1)
    End If
    
    'Schedule Status
    slString = Trim(mGetSchdStatus(tlChfDetail.sSchStatus))
    ilFound = 0
    For llLoop = 0 To UBound(smScheduleStatusList)
        If smScheduleStatusList(llLoop) = slString Then
            ilFound = 1
            Exit For
        End If
    Next llLoop
    If ilFound = 0 Then
        smScheduleStatusList(UBound(smScheduleStatusList)) = slString
        ReDim Preserve smScheduleStatusList(0 To UBound(smScheduleStatusList) + 1)
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mClearCntrGrid
' Description:       clears grdCntr and sets column headers and column widths
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:08:02
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mClearCntrGrid()
    Dim ilLoop As Integer
    grdCntr.rows = 2
    grdCntr.FixedCols = 0
    grdCntr.FixedRows = 1
    grdCntr.cols = 17
    cmcChgCntr.Enabled = False
    cmcViewCntr.Enabled = False
    cmcSchedule.Enabled = False
    cmcExport.Enabled = False
    
    'ChfCode: C_CHFCODEINDEX
    grdCntr.TextMatrix(0, C_CHFCODEINDEX) = "ChfCode"
    DashboardVw.grdCntr.ColWidth(C_CHFCODEINDEX) = 0
    
    'Sort: C_SORTINDEX
    grdCntr.TextMatrix(0, C_SORTINDEX) = "Sort"
    DashboardVw.grdCntr.ColWidth(C_SORTINDEX) = 0
        
    'Update date: C_CNTRUPDATEDATEINDEX, chfOHDDate
    grdCntr.TextMatrix(0, C_CNTRUPDATEDATEINDEX) = "Upd. Date"
    DashboardVw.grdCntr.ColWidth(C_CNTRUPDATEDATEINDEX) = mGetColumnSize(C_CNTRUPDATEDATEINDEX) ' 870
    
    'Contract number: C_CNTRNOINDEX, if it's a proposal status, append the proposal version number to the contract number. For example, if it's contract 65, proposal version 1, show "65 V1"
    grdCntr.TextMatrix(0, C_CNTRNOINDEX) = "Contract"
    DashboardVw.grdCntr.ColWidth(C_CNTRNOINDEX) = mGetColumnSize(C_CNTRNOINDEX) ' 1110
    
    'Contract type: C_CNTRTYPEINDEX
    grdCntr.TextMatrix(0, C_CNTRTYPEINDEX) = "Contract Type"
    DashboardVw.grdCntr.ColWidth(C_CNTRTYPEINDEX) = mGetColumnSize(C_CNTRTYPEINDEX) ' 1300
    
    'Line Type: C_LINETYPEINDEX
    grdCntr.TextMatrix(0, C_LINETYPEINDEX) = "Line Type"
    DashboardVw.grdCntr.ColWidth(C_LINETYPEINDEX) = mGetColumnSize(C_LINETYPEINDEX)
    
    'Agency: C_AGYNAMEINDEX, show "n/a" or something similar for a direct advertiser
    grdCntr.TextMatrix(0, C_AGYNAMEINDEX) = "Agency"
    DashboardVw.grdCntr.ColWidth(C_AGYNAMEINDEX) = mGetColumnSize(C_AGYNAMEINDEX) '1950
    
    'Advertiser: C_ADVNAMEINDEX
    grdCntr.TextMatrix(0, C_ADVNAMEINDEX) = "Advertiser"
    DashboardVw.grdCntr.ColWidth(C_ADVNAMEINDEX) = mGetColumnSize(C_ADVNAMEINDEX) ' 2280
    
    'Product: C_PRODUCTINDEX
    grdCntr.TextMatrix(0, C_PRODUCTINDEX) = "Product"
    DashboardVw.grdCntr.ColWidth(C_PRODUCTINDEX) = mGetColumnSize(C_PRODUCTINDEX) '2580
    
    'Start date: C_STARTDATEINDEX, (from contract header)
    grdCntr.TextMatrix(0, C_STARTDATEINDEX) = "Start Date"
    DashboardVw.grdCntr.ColWidth(C_STARTDATEINDEX) = mGetColumnSize(C_STARTDATEINDEX) '870
    
    'End date: C_ENDDATEINDEX, (from contract header)
    grdCntr.TextMatrix(0, C_ENDDATEINDEX) = "End Date"
    DashboardVw.grdCntr.ColWidth(C_ENDDATEINDEX) = mGetColumnSize(C_ENDDATEINDEX) '870
    
    'Gross: C_GROSSINDEX, chfInputGross
    grdCntr.TextMatrix(0, C_GROSSINDEX) = "Gross"
    DashboardVw.grdCntr.ColWidth(C_GROSSINDEX) = mGetColumnSize(C_GROSSINDEX) '1110
    
    'Sales Office: C_SALESOFFICEINDEX, from primary salesperson
    grdCntr.TextMatrix(0, C_SALESOFFICEINDEX) = "Office"
    DashboardVw.grdCntr.ColWidth(C_SALESOFFICEINDEX) = mGetColumnSize(C_SALESOFFICEINDEX) '1380
    
    'Primary Salesperson: C_SALEPERSONINDEX, primary salesperson first and last name
    grdCntr.TextMatrix(0, C_SALEPERSONINDEX) = "Salesperson"
    DashboardVw.grdCntr.ColWidth(C_SALEPERSONINDEX) = mGetColumnSize(C_SALEPERSONINDEX) '1800
    
    'Contract Status: C_CNTRSTATUSINDEX (Working, Rejected, Approved Order, Rev Working, etc.)
    'Note: I think a chfStatus of W and a chfSchStatus of P means it's a "Rev Working" status, which is a special status for a contract that has been scheduled and revised on the proposal screen after it was scheduled
    grdCntr.TextMatrix(0, C_CNTRSTATUSINDEX) = "Contract Status"
    DashboardVw.grdCntr.ColWidth(C_CNTRSTATUSINDEX) = mGetColumnSize(C_CNTRSTATUSINDEX) '1530
    
    'Digital delivery status: C_DIGITALDLVYINDEX, new field to be added to the CHF as part of the Megaphone integration to track Pushed, Not Pushed, and N/A (for non-digital line contracts)
    grdCntr.TextMatrix(0, C_DIGITALDLVYINDEX) = "Digital Delivery"
    DashboardVw.grdCntr.ColWidth(C_DIGITALDLVYINDEX) = mGetColumnSize(C_DIGITALDLVYINDEX) '1200
    
    'Schedule status: C_CNTRSCHEDULESTATUSINDEX, scheduled, not scheduled, prevent, interrupted, etc. (chfSchStatus)
    grdCntr.TextMatrix(0, C_CNTRSCHEDULESTATUSINDEX) = "Sched. Status"
    DashboardVw.grdCntr.ColWidth(C_CNTRSCHEDULESTATUSINDEX) = mGetColumnSize(C_CNTRSCHEDULESTATUSINDEX) '1530
    
    'Clear 1st row, Set Alignment, set header row colors
    For ilLoop = 0 To grdCntr.cols - 1
        DashboardVw.grdCntr.ColAlignment(ilLoop) = flexAlignLeftTop
        grdCntr.TextMatrix(1, ilLoop) = ""
        grdCntr.Row = 0
        grdCntr.Col = ilLoop
        grdCntr.CellBackColor = LIGHTBLUE
    Next ilLoop
    DashboardVw.grdCntr.ColAlignment(C_GROSSINDEX) = flexAlignRightTop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mClearFilterInputs
' Description:       Clears user filter input, Unhighlights selected filter in txtCriteria
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:08:39
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mClearFilterInputs()
    'Clear search textbox
    txtSearch.Text = ""
    
    'Unhighlight filters
    txtCriteria.SelStart = 0
    txtCriteria.SelLength = Len(txtCriteria.Text)
    txtCriteria.SelColor = vbBlack
    txtCriteria.SelBold = False
    txtCriteria.SelStart = Len(txtCriteria.Text)
    
    lblSelectedFilter.Caption = ""
    grdCntr.Enabled = True
    
    'Show Add button
    cmcUpdateFilter.Visible = False
    cmcAddFilter.Visible = True
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mExportToCSV
' Description:       creates a .csv file in sgExportPath with the content of grdCntr
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:38:02
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mExportToCSV()
    Dim omBook As Object
    Dim omSheet As Object
    Dim ilRow As Integer
    Dim ilColumn As Integer
    Dim slDelimiter As String
    Dim ilDOE As Integer
    Dim slString As String
    Dim hmExport As Integer
    Dim slFileName As String
    slFileName = sgExportPath & "Dashboard_" & Format(Now, "yyyymmdd_hhMMss") & ".CSV"
    
    slDelimiter = ","
    Screen.MousePointer = vbHourglass

    lblItemCount.Caption = "Exporting CSV..."
    ProgressBar1.Value = 1
    ProgressBar1.Visible = True
    imCancelLoading = False
    ProgressBar1.Visible = True
    cmcCancel.Visible = True
    ProgressBar1.Left = cmcCancel.Left - ProgressBar1.Width - 80
    lblItemCount.Left = ProgressBar1.Left - lblItemCount.Width - 120
    lblMultiVersionTotalWarn.Left = lblItemCount.Left
    
    'Create File
    ilRet = gFileOpen(slFileName, "OUTPUT", hmExport)
    If ilRet <> 0 Then
        MsgBox "Error writing file:" & sgExportPath & slFileName & vbCrLf & "Error:" & ilRet & " - " & Error(ilRet)
        Close #hmExport
        GoTo Done
    End If
    
    'Write Header
    ilRow = 1
    slRecord = ""
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        If grdCntr.ColWidth(ilColumn) > 0 Then
            If slRecord <> "" Then slRecord = slRecord & slDelimiter
            slRecord = slRecord & """" & grdCntr.TextMatrix(ilRow - 1, ilColumn) & """"
        End If
    Next ilColumn
    
    'Write Header
    Print #hmExport, slRecord
    
    'Write rows of Data
    For ilRow = 2 To grdCntr.rows
        ilDOE = ilDOE + 1
        If ilDOE > 10 Then
            ProgressBar1.Value = ilRow * (100 / grdCntr.rows)
            DoEvents
        End If
        
        If imCancelLoading = True Then
            MsgBox "Canceled Excel"
            imCancelLoading = False
            GoTo Done
        End If
        slRecord = ""
        For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
            If grdCntr.ColWidth(ilColumn) > 0 Then
                If slRecord <> "" Then slRecord = slRecord & slDelimiter
                slString = grdCntr.TextMatrix(0, ilColumn)
                If slString = "Gross" Or InStr(1, slString, "Date") > 0 Then
                    'No Quotes (strip commas)
                    slRecord = slRecord & Replace(grdCntr.TextMatrix(ilRow - 1, ilColumn), ",", "")
                Else
                    'Quotes
                    slRecord = slRecord & """" & grdCntr.TextMatrix(ilRow - 1, ilColumn) & """"
                    
                End If
            End If
        Next ilColumn
        Print #hmExport, slRecord
    Next ilRow
    
    'View Excel
    Close #hmExport
    MsgBox "Exported Dashboard: " & vbCrLf & slFileName, vbInformation + vbOKOnly, "Dashboard"
Done:
    ProgressBar1.Value = 100
    ProgressBar1.Visible = False
    cmcCancel.Visible = False
    cmcClose.Enabled = True
    lblItemCount.Caption = "Showing " & Format(lmItemCount, "#,##0") & " Items"

    Screen.MousePointer = vbDefault
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSendToExcel
' Description:       Automates Excel, creates a worksheet of the contents of grdCntr
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:39:07
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mSendToExcel()
    Dim omBook As Object
    Dim omSheet As Object
    Dim imExcelRow As Integer
    Dim ilRow As Integer
    Dim ilColumn As Integer
    Dim slDelimiter As String
    Dim ilDOE As Integer
    Dim ilExcelCol As Integer
    Dim slString As String
    Dim blNeedAnd As Boolean
    Dim ilExcelRow As Integer
    'Dim slString As String
    Dim slString2 As String
    Dim ilFilterRows As Integer
    ilExcelRow = 1
    slDelimiter = Chr$(30) 'ASCII 30 is defined as a "Record Separator" - https://www.asciitable.com/
    Screen.MousePointer = vbHourglass

    lblItemCount.Caption = "Opening in Excel..."
    ProgressBar1.Value = 1
    ProgressBar1.Visible = True
    imCancelLoading = False
    ProgressBar1.Visible = True
    cmcCancel.Visible = True
    ProgressBar1.Left = cmcCancel.Left - ProgressBar1.Width - 80
    lblItemCount.Left = ProgressBar1.Left - lblItemCount.Width - 120
    lblMultiVersionTotalWarn.Left = lblItemCount.Left
        
    'Create an Excel sheet
    ilRet = gExcelOutputGeneration("C")
    If ilRet = False Then
        MsgBox "Error opening Excel!", vbCritical + vbOKOnly
        GoTo Done
    End If
    
    'Open (Book and Sheet): (Parameters: slAction, olBook, olSheet, ilSheetNo)
    ilRet = gExcelOutputGeneration("O", omBook, omSheet, 1)
    
    'Make Report Header
    If chkExcelHeader.Value = vbChecked Then
        slRecord = "Dashboard:" & slDelimiter & Format(Now, "mm/dd/yy hh:MM:ss AMPM ")
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
        ilExcelRow = ilExcelRow + 1
        
        'Proposals/Orders
        slRecord = "For:" & slDelimiter
        blNeedAnd = False
        If ckcCntrType(3).Value = vbChecked Then
            slRecord = slRecord & "Proposals": blNeedAnd = True
        End If
        If ckcCntrType(4).Value = vbChecked Then
            If blNeedAnd = True Then slRecord = slRecord & ", "
            slRecord = slRecord & "Orders"
        End If
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
        ilExcelRow = ilExcelRow + 1
        
        'AT/NTR/Dig
        slRecord = "Showing:" & slDelimiter
        blNeedAnd = False
        If ckcCntrType(0).Value = vbChecked Then
            slRecord = slRecord & "Air Time": blNeedAnd = True
        End If
        If ckcCntrType(1).Value = vbChecked Then
            If blNeedAnd = True Then slRecord = slRecord & ", "
            slRecord = slRecord & "NTR": blNeedAnd = True
        End If
        If ckcCntrType(2).Value = vbChecked Then
            If blNeedAnd = True Then slRecord = slRecord & ", "
            slRecord = slRecord & "Digital"
        End If
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
        ilExcelRow = ilExcelRow + 1
        
        'Filters
        slString = mGetFiltersByType("Cntr")
        If slString <> "" Then
            slRecord = "Contract:" & slDelimiter & slString
            ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
            ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
            ilExcelRow = ilExcelRow + 1
        Else
            'Contract filter not used, get applied Filter type values
            slString = mGetFiltersByType("Adv")
            If slString <> "" Then
                slRecord = "Advertiser:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("Agy")
            If slString <> "" Then
                slRecord = "Agency:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("*Active>")
            slString2 = mGetFiltersByType("Active<")
            If slString <> "" And slString2 <> "" Then
                slRecord = "Active:" & slDelimiter & slString & " to " & slString2
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            Else
                If slString <> "" Then
                    slRecord = "Active on or after:" & slDelimiter & slString
                    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                    ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                    ilExcelRow = ilExcelRow + 1
                End If
                If slString2 <> "" Then
                    slRecord = "Active on or prior:" & slDelimiter & slString
                    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                    ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                    ilExcelRow = ilExcelRow + 1
                End If
            End If
            
            slString = mGetFiltersByType("UDate>")
            slString2 = mGetFiltersByType("UDate<")
            If slString <> "" And slString2 <> "" Then
                slRecord = "Update Date:" & slDelimiter & slString & " to " & slString2
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            Else
                If slString <> "" Then
                    slRecord = "Update Date (Beginning):" & slDelimiter & slString
                    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                    ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                    ilExcelRow = ilExcelRow + 1
                End If
                If slString2 <> "" Then
                    slRecord = "Update Date (Ending):" & slDelimiter & slString
                    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                    ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                    ilExcelRow = ilExcelRow + 1
                End If
            End If
            
            slString = mGetFiltersByType("CStat")
            If slString <> "" Then
                slRecord = "Contract Status:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("DlvyStat")
            If slString <> "" Then
                slRecord = "Digital Delivery Status:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("SchStat")
            If slString <> "" Then
                slRecord = "Schedule Status:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("SP")
            If slString <> "" Then
                slRecord = "Salesperson:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("Group")
            If slString <> "" Then
                slRecord = "Sales Group:" & slDelimiter & imCurrentGroup
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
                        
            slString = mGetFiltersByType("SO")
            If slString <> "" Then
                slRecord = "Sales Office:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
            
            slString = mGetFiltersByType("Prod")
            If slString <> "" Then
                slRecord = "Product:" & slDelimiter & slString
                ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
                ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, 1)
                ilExcelRow = ilExcelRow + 1
            End If
        End If
    End If
    ilFilterRows = ilExcelRow
    ilExcelRow = ilExcelRow + 1
    
    'Make Col Headers
    slRecord = ""
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        If grdCntr.ColWidth(ilColumn) > 0 Then
            If slRecord <> "" Then slRecord = slRecord & slDelimiter
            slRecord = slRecord & grdCntr.TextMatrix(0, ilColumn)
        End If
    Next ilColumn
    
    'Write Header
    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow, 1, slDelimiter)
    'Font bold header
    For ilColumn = 1 To omSheet.UsedRange.Columns.Count
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", ilExcelRow, ilColumn)
    Next ilColumn
    
    'Column Widths
    ilExcelCol = 0
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        If grdCntr.ColWidth(ilColumn) > 0 Then
            ilExcelCol = ilExcelCol + 1
            If grdCntr.ColWidth(ilColumn) > 800 Then
                ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , grdCntr.ColWidth(ilColumn) / 80, , ilExcelCol) 'ColWidth
            End If
        End If
    Next ilColumn
    
    'Write rows of Data
    For ilRow = 2 To grdCntr.rows
        ilDOE = ilDOE + 1
        If ilDOE > 10 Then
            ProgressBar1.Value = ilRow * (100 / grdCntr.rows)
            DoEvents
        End If
        
        If imCancelLoading = True Then
            MsgBox "Canceled Excel"
            imCancelLoading = False
            GoTo Done
        End If
        slRecord = ""
        For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
            If grdCntr.ColWidth(ilColumn) > 0 Then
                If slRecord <> "" Then slRecord = slRecord & slDelimiter
                slRecord = slRecord & grdCntr.TextMatrix(ilRow - 1, ilColumn)
            End If
        Next ilColumn
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilExcelRow + ilRow - 1, 1, slDelimiter)
        'ilExcelRow = ilExcelRow + 1
    Next ilRow
    
    'Column Horizontal Align
    ilExcelCol = 0
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        If grdCntr.ColWidth(ilColumn) > 0 Then
            ilExcelCol = ilExcelCol + 1
            slString = grdCntr.TextMatrix(0, ilColumn)
            If slString = "Contract" Or slString = "Gross" Or InStr(1, slString, "Date") > 0 Then
                ilRet = gExcelOutputGeneration("HA", omBook, omSheet, , str(xlRight), -1, ilExcelCol)
            End If
        End If
    Next ilColumn
    
    'Horizontal Alignment of Filter names & Filter Values
    If chkExcelHeader.Value = vbChecked Then
        For ilRow = 1 To ilFilterRows
            ilRet = gExcelOutputGeneration("HA", omBook, omSheet, , str(xlLeft), ilRow, 1)
            ilRet = gExcelOutputGeneration("HA", omBook, omSheet, , str(xlLeft), ilRow, 2)
        Next ilRow
    End If
    
    'View Excel
    ilRet = gExcelOutputGeneration("V")
    
Done:
    ProgressBar1.Value = 100
    ProgressBar1.Visible = False
    cmcCancel.Visible = False
    cmcClose.Enabled = True
    lblItemCount.Caption = "Showing " & Format(lmItemCount, "#,##0") & " Items"
    Screen.MousePointer = vbDefault
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mPopulate
' Description:       Populate tmChfDetailList with list of Contracts for a Date range
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-11:54:47
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mPopulate()
    lblItemCount.Caption = "Loading contract list..."
    DashboardVw.Refresh
    Dim slStartDate As String
    Dim slEndDate As String
    
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    cmcClose.Enabled = False
    cmcViewCntr.Enabled = False
    cmcChgCntr.Enabled = False
    cmcSchedule.Enabled = False
    
    ProgressBar1.Value = 10
    ProgressBar1.Visible = True
    imCancelLoading = False
    ProgressBar1.Visible = True
    cmcCancel.Visible = True
    ProgressBar1.Left = cmcCancel.Left - ProgressBar1.Width - 80
    lblItemCount.Left = ProgressBar1.Left - lblItemCount.Width - 80
    lblMultiVersionTotalWarn.Left = lblItemCount.Left
    
    Me.Refresh
    DoEvents
    
    Screen.MousePointer = vbHourglass  'Wait
    slStartDate = mGetFilterValue("*Active>", "01/01/1969")
    slEndDate = mGetFilterValue("Active<", "12/31/2070")
    
    If slStartDate = smLastStartDate And slEndDate = smLastEndDate Then
        'This Date range was already loaded in tmChfDetailList
    Else
        'Load Contracts in Date range from Database
        bmBuildingFilteredResults = True
        ReDim smAdvertiserList(0 To 0) As String
        ReDim smAgencyList(0 To 0) As String
        ReDim smSalesOfficeList(0 To 0) As String
        ReDim smSalespersonList(0 To 0) As String
        ReDim smContractStatusList(0 To 0) As String
        ReDim smContractTypeList(0 To 0) As String
        ReDim smDeliveryStatusList(0 To 0) As String
        ReDim smScheduleStatusList(0 To 0) As String
        
        ilRet = gCntrDetailForDateRange(DashboardVw, slStartDate, slEndDate, tmChfDetailList())
        smLastStartDate = slStartDate
        smLastEndDate = slEndDate
    End If
    ProgressBar1.Value = 20
    
    'Filter loaded Contracts and show in Grid
    mApplyFilter
    
    ProgressBar1.Value = 100
    Screen.MousePointer = vbDefault
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSchedule
' Description:       Opens the Scheduler
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-08:39:57
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mSchedule()
    Dim ilSchSelCntr As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim llCntrRowSelected As Long
    llCntrRowSelected = grdCntr.Row
    ilSchSelCntr = False
    
    If llCntrRowSelected >= grdCntr.FixedRows Then
        If grdCntr.TextMatrix(llCntrRowSelected, C_CNTRNOINDEX) = "Altered Contract" Or grdCntr.TextMatrix(llCntrRowSelected, C_CNTRSTATUSINDEX) = "New Contract" Then
            ilSchSelCntr = True
        Else
            Exit Sub
        End If
    End If

    If ilSchSelCntr Then
        CntrSch.Show vbModal
        mPopulate
    End If
    Exit Sub
mScheduleErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSelectFilterItem
' Description:       Highlight the clicked filter in txtCriteria, Select the
'                    filter type and set the SearchText with the Filter text
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:10:09
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mSelectFilterItem()
    Dim ilStartPos As Integer
    Dim ilEndPos As Integer
    Dim ilCurrentPos As Integer
    Dim ilLoop As Integer
    Dim slTerm As String
    cmcRemoveFilter.Enabled = False
    cmcUpdateFilter.Enabled = False
    cmcAddFilter.Enabled = False
    cmcUpdateFilter.Visible = False
    cmcAddFilter.Visible = True
    cboFilterType.Enabled = True
    cboItems.Enabled = True
    
    lblSelectedFilter.Caption = ""
    grdCntr.Enabled = True
    ilCurrentPos = txtCriteria.SelStart
    If ilCurrentPos = 0 Then ilCurrentPos = 1
    'Find starting [
    For ilLoop = ilCurrentPos To 1 Step -1
        If Mid(txtCriteria.Text, ilLoop, 1) = "]" Then
            'if we find an end then user clicked somewhere odd
            mClearFilterInputs
            Exit Sub
        End If
        If Mid(txtCriteria.Text, ilLoop, 1) = "[" Then
            ilStartPos = ilLoop
            Exit For
        End If
    Next ilLoop
    
    'Find ending ]
    For ilLoop = ilCurrentPos To Len(txtCriteria.Text) Step 1
        If Mid(txtCriteria.Text, ilLoop, 1) = "[" Then
            'if we find an start then user clicked somewhere odd
            mClearFilterInputs
            Exit Sub
        End If
        If Mid(txtCriteria.Text, ilLoop, 1) = "]" Then
            ilEndPos = ilLoop
            Exit For
        End If
    Next ilLoop
    
    If ilStartPos = 0 Then
        mClearFilterInputs
        Exit Sub
    End If
    If ilEndPos = 0 Then
        mClearFilterInputs
        Exit Sub
    End If
    
    'Unhighlight
    txtCriteria.SelStart = 0
    txtCriteria.SelLength = Len(txtCriteria.Text)
    txtCriteria.SelColor = vbBlack
    txtCriteria.SelBold = False
    
    'highlight selected filter
    txtCriteria.SelStart = ilStartPos - 1
    txtCriteria.SelLength = ilEndPos - ilStartPos + 1
    txtCriteria.SelColor = vbBlue
    txtCriteria.SelBold = True
    txtCriteria.SelLength = 0
    
    cboFilterType.Enabled = False
    slTerm = Mid(txtCriteria.Text, ilStartPos + 1, ilEndPos - ilStartPos - 1)
    
    'ADV: Advertiser
    If Mid(slTerm, 1, 4) = "Adv:" Then
        mSelectFilterType "Advertiser"
        txtSearch.Text = Trim(Mid(slTerm, 5))
    End If
    
    'Agy: Agency
    If Mid(slTerm, 1, 4) = "Agy:" Then
        mSelectFilterType "Agency"
        txtSearch.Text = Trim(Mid(slTerm, 5))
    End If
        
    'Active Date Beginning
    If Mid(slTerm, 1, 8) = "Active>:" Then
        mSelectFilterType "Active on or after"
        txtSearch.Text = Trim(Mid(slTerm, 9))
        CSI_Calendar1.Text = Trim(txtSearch.Text)
    End If

    'Active Date Ending
    If Mid(slTerm, 1, 8) = "Active<:" Then
        mSelectFilterType "Active on or prior"
        txtSearch.Text = Trim(Mid(slTerm, 9))
        CSI_Calendar1.Text = Trim(txtSearch.Text)
    End If
    
    'Upd: Update Date
    If Mid(slTerm, 1, 7) = "UDate>:" Then
        mSelectFilterType "Update Date (Beginning)"
        txtSearch.Text = Trim(Mid(slTerm, 8))
        CSI_Calendar1.Text = Trim(txtSearch.Text)
    End If
    
    'Upd: Update Date
    If Mid(slTerm, 1, 7) = "UDate<:" Then
        mSelectFilterType "Update Date (Ending)"
        txtSearch.Text = Trim(Mid(slTerm, 8))
        CSI_Calendar1.Text = Trim(txtSearch.Text)
    End If
    
    'Cntr: Contract Number
    If Mid(slTerm, 1, 5) = "Cntr:" Then
        mSelectFilterType "Contract Number"
        txtSearch.Text = Trim(Mid(slTerm, 6))
    End If
    
    'CType: Contract Type
    If Mid(slTerm, 1, 6) = "CType:" Then
        mSelectFilterType "Contract Type"
        txtSearch.Text = Trim(Mid(slTerm, 7))
    End If
    
    'CStat: Contract Status
    If Mid(slTerm, 1, 6) = "CStat:" Then
        mSelectFilterType "Contract Status"
        txtSearch.Text = Trim(Mid(slTerm, 7))
    End If
    
    'DlvyStat: Delivery Status
    If Mid(slTerm, 1, 9) = "DlvyStat:" Then
        mSelectFilterType "Delivery Status"
        txtSearch.Text = Trim(Mid(slTerm, 10))
    End If
    
    'SchStat: Schedule Status
    If Mid(slTerm, 1, 8) = "SchStat:" Then
        mSelectFilterType "Schedule Status"
        txtSearch.Text = Trim(Mid(slTerm, 9))
    End If
    
    'Prod: Product
    If Mid(slTerm, 1, 5) = "Prod:" Then
        mSelectFilterType "Product"
        txtSearch.Text = Trim(Mid(slTerm, 6))
    End If
    
    'SP: SalesPerson
    If Mid(slTerm, 1, 3) = "SP:" Then
        mSelectFilterType "Salesperson"
        txtSearch.Text = Trim(Mid(slTerm, 4))
    End If
    
    'Group
    If Mid(slTerm, 1, 6) = "Group:" Then
        mSelectFilterType "Sales Group"
        txtSearch.Text = imCurrentGroup
        txtSearch.Enabled = False
    End If
    
    'Mandatory SalesPerson
    If Mid(slTerm, 1, 4) = "*SP:" Then
        mSelectFilterType "Salesperson"
        txtSearch.Text = Trim(Mid(slTerm, 5))
        txtSearch.Enabled = False
        cboItems.Enabled = False
        txtCriteria.Enabled = True
        lblSelectedFilter.Caption = "You Cannot change this item"
        Exit Sub
    End If
    
    'SO: Sales Office
    If Mid(slTerm, 1, 3) = "SO:" Then
        mSelectFilterType "Sales Office"
        txtSearch.Text = Trim(Mid(slTerm, 4))
    End If
    
    'Highlight the selected term in the search
    If txtSearch.Enabled = True Then
        If txtSearch.Visible = True And txtSearch.Enabled = True Then txtSearch.SetFocus
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    
    'Show the selected filter
    lblSelectedFilter.Caption = "[" & slTerm & "]"
    grdCntr.Enabled = False
    
    'Mandatory Active Start Date (Can edit, but Can't Delete)
    If Mid(slTerm, 1, 9) = "*Active>:" Then
        mSelectFilterType "Active on or after"
        txtSearch.Text = Trim(Mid(slTerm, 10))
        CSI_Calendar1.Text = Trim(txtSearch.Text)
        cboItems.Enabled = False
        txtCriteria.Enabled = True
        cmcUpdateFilter.Visible = True
        cmcAddFilter.Visible = False
        cmcRemoveFilter.Enabled = False
        cmcUpdateFilter.Enabled = True
        Exit Sub
    End If
    
    'Show Update button
    cmcUpdateFilter.Visible = True
    cmcAddFilter.Visible = False
    cmcRemoveFilter.Enabled = True
    cmcUpdateFilter.Enabled = True
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mTerminate
' Description:       Terminate form
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       1/31/2024-13:27:07
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mTerminate()
    mSaveDashSettings
    Erase tmSofList
    Erase tmChfDetailList
    smLastStartDate = ""
    smLastEndDate = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload DashboardVw
    igManUnload = NO
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSetDefaultColumnSizes
' Description:       Sets the Default column sizes, if Column Number provided - it sets only that column's Default size
'                    If small screen, it will set column sizes more narrow
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/2/2024-09:02:55
' Parameters :       ilColumnNumber (Integer) - optional.  If provided, set only the provided column size
'--------------------------------------------------------------------------------
Private Sub mSetDefaultColumnSizes(Optional ilColumnNumber As Integer = 0, Optional blResetting As Boolean = False)
    Dim ilSmall As Integer 'Small Screen size?
    Dim llExcludedColWidths As Long
    Dim llRemainingWidth As Long
    ilSmall = False
    If Me.Width < 17060 Then ilSmall = True
    
    'These columns dont dynamically size based on screen size (UpdateDt, CntrNo, LineType, StartDate,EndDate,Gross)
    If ilColumnNumber = 0 Or ilColumnNumber = C_CNTRUPDATEDATEINDEX Or blResetting Then
        If imColumnWidths(C_CNTRUPDATEDATEINDEX) < 1 Then imColumnWidths(C_CNTRUPDATEDATEINDEX) = IIF(ilSmall, 840, 880)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_CNTRNOINDEX Or blResetting Then
        If imColumnWidths(C_CNTRNOINDEX) < 1 Then imColumnWidths(C_CNTRNOINDEX) = IIF(ilSmall, 960, 1110)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_LINETYPEINDEX Or blResetting Then
        If imColumnWidths(C_LINETYPEINDEX) < 1 Then imColumnWidths(C_LINETYPEINDEX) = IIF(ilSmall, 810, 870)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_STARTDATEINDEX Or blResetting Then
        If imColumnWidths(C_STARTDATEINDEX) < 1 Then imColumnWidths(C_STARTDATEINDEX) = IIF(ilSmall, 840, 880)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_ENDDATEINDEX Or blResetting Then
        If imColumnWidths(C_ENDDATEINDEX) < 1 Then imColumnWidths(C_ENDDATEINDEX) = IIF(ilSmall, 840, 880)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_GROSSINDEX Or blResetting Then
        If imColumnWidths(C_GROSSINDEX) < 1 Then imColumnWidths(C_GROSSINDEX) = IIF(ilSmall, 1110, 1110)
    End If
    
    'Dynamically sizing columns (Exclude UpdateDt, CntrNo, LineType, StartDate,EndDate,Gross)
    llExcludedColWidths = imColumnWidths(C_CNTRUPDATEDATEINDEX) + imColumnWidths(C_CNTRNOINDEX)
    llExcludedColWidths = llExcludedColWidths + imColumnWidths(C_LINETYPEINDEX)
    llExcludedColWidths = llExcludedColWidths + imColumnWidths(C_STARTDATEINDEX) + imColumnWidths(C_STARTDATEINDEX)
    llExcludedColWidths = llExcludedColWidths + imColumnWidths(C_GROSSINDEX)
    llExcludedColWidths = llExcludedColWidths + 120 'Scrollbar width adjustment
    llExcludedColWidths = llExcludedColWidths + (15 * 20) 'gutter adjustment

    'resize remaining columns based on percent
    llRemainingWidth = (grdCntr.Width - llExcludedColWidths) / 100
    If ilColumnNumber = 0 Or ilColumnNumber = C_CNTRTYPEINDEX Then
        If imColumnWidths(C_CNTRTYPEINDEX) < 1 Then imColumnWidths(C_CNTRTYPEINDEX) = IIF(ilSmall, 825, llRemainingWidth * 7)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_AGYNAMEINDEX Then
        If imColumnWidths(C_AGYNAMEINDEX) < 1 Then imColumnWidths(C_AGYNAMEINDEX) = IIF(ilSmall, 1710, llRemainingWidth * 16)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_ADVNAMEINDEX Then
        If imColumnWidths(C_ADVNAMEINDEX) < 1 Then imColumnWidths(C_ADVNAMEINDEX) = IIF(ilSmall, 1620, llRemainingWidth * 17)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_PRODUCTINDEX Then
        If imColumnWidths(C_PRODUCTINDEX) < 1 Then imColumnWidths(C_PRODUCTINDEX) = IIF(ilSmall, 1380, llRemainingWidth * 18)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_SALESOFFICEINDEX Then
        If imColumnWidths(C_SALESOFFICEINDEX) < 1 Then imColumnWidths(C_SALESOFFICEINDEX) = IIF(ilSmall, 645, llRemainingWidth * 7)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_SALEPERSONINDEX Then
        If imColumnWidths(C_SALEPERSONINDEX) < 1 Then imColumnWidths(C_SALEPERSONINDEX) = IIF(ilSmall, 960, llRemainingWidth * 10)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_CNTRSTATUSINDEX Then
        If imColumnWidths(C_CNTRSTATUSINDEX) < 1 Then imColumnWidths(C_CNTRSTATUSINDEX) = IIF(ilSmall, 840, llRemainingWidth * 9)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_DIGITALDLVYINDEX Then
        If imColumnWidths(C_DIGITALDLVYINDEX) < 1 Then imColumnWidths(C_DIGITALDLVYINDEX) = IIF(ilSmall, 555, llRemainingWidth * 7)
    End If
    If ilColumnNumber = 0 Or ilColumnNumber = C_CNTRSCHEDULESTATUSINDEX Then
        If imColumnWidths(C_CNTRSCHEDULESTATUSINDEX) < 1 Then imColumnWidths(C_CNTRSCHEDULESTATUSINDEX) = IIF(ilSmall, 765, llRemainingWidth * 9)
    End If
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mGetColumnSize
' Description:       Get the column size for a given Column number
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/2/2024-09:03:36
' Parameters :       ilColumn (Integer)
'--------------------------------------------------------------------------------
Private Function mGetColumnSize(ilColumn As Integer) As Integer
    mGetColumnSize = imColumnWidths(ilColumn)
    If mGetColumnSize = -1 Then mGetColumnSize = 0
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mUpdateDefaultColumSizes
' Description:       Remembers the column sizes
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/2/2024-09:02:09
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mUpdateDefaultColumSizes()
    Dim ilColumn As Integer
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        If grdCntr.ColWidth(ilColumn) > 0 Then
            imColumnWidths(ilColumn) = grdCntr.ColWidth(ilColumn)
        Else
            imColumnWidths(ilColumn) = -1
        End If
    Next ilColumn
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSaveDashSettings
' Description:       Saves the current Dashboard settings
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/5/2024-09:50:36
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mSaveDashSettings()
    Dim slSQLQuery As String
    Dim llLongData As Long
    Dim slStringData As String
    Dim llCount As Long
    Dim ilColumn As Integer
    Dim ilLoop As Integer
    
    'Clear Saved Data
    slSQLQuery = "DELETE FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfurfCode = " & tgUrf(0).iCode
    gSQLAndReturn slSQLQuery, False, llCount
    
    'Start Dashboard on Startup?
    tgUrf(0).iDashOnStart = IIF(ckcShowOnStartup.Value = vbChecked, 1, 0)
    slSQLQuery = "Update URF_User_Options set urfDashOnStart = " & tgUrf(0).iDashOnStart & " WHERE urfCode = " & tgUrf(0).iCode
    gSQLAndReturn slSQLQuery, False, llCount
    
    'Excel Headers
    slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'ExcelHeader', " & chkExcelHeader.Value & ", " & tgUrf(0).iCode & ")"
    gSQLAndReturn slSQLQuery, False, llCount

    'Show Checkboxes
    For ilLoop = 0 To 4
        slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfCtrlIndex, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'Show', " & ilLoop & ", " & ckcCntrType(ilLoop).Value & ", " & tgUrf(0).iCode & ")"
        gSQLAndReturn slSQLQuery, False, llCount
    Next ilLoop
    
    'Font Size
    Select Case CboFontSize.ListIndex
        Case 0: llLongData = 900
        Case 1: llLongData = 825
        Case 2: llLongData = 700
        Case 3: llLongData = 600
    End Select
    slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfCtrlIndex, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'FontSize', " & CboFontSize.ListIndex & ", " & llLongData & ", " & tgUrf(0).iCode & ")"
    gSQLAndReturn slSQLQuery, False, llCount
    
    'Sort Column
    slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'SortColumn', " & imLastSortCol & ", " & tgUrf(0).iCode & ")"
    gSQLAndReturn slSQLQuery, False, llCount
    
    'Sort Dir
    slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'SortDir', " & imSortDir & ", " & tgUrf(0).iCode & ")"
    gSQLAndReturn slSQLQuery, False, llCount
    
    'Columns Visible / Column Sizes
    For ilColumn = C_CNTRUPDATEDATEINDEX To grdCntr.cols - 1
        llLongData = grdCntr.ColWidth(ilColumn)
        If llLongData = 0 Then llLongData = -1
        slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfCtrlIndex, fcfLongData, fcfUrfCode) VALUES ('Dashboard', 'ColumnSize', " & ilColumn & ", " & llLongData & ", " & tgUrf(0).iCode & ")"
        gSQLAndReturn slSQLQuery, False, llCount
    Next ilColumn
    
    'Active Filters
    For ilLoop = 0 To UBound(tmAppliedFilters)
        If tmAppliedFilters(ilLoop).sType <> "" Then
            slSQLQuery = "INSERT INTO FCF_Form_Control (fcfName, fcfCtrlName, fcfLongData, fcfStringData, fcfUrfCode) VALUES ('Dashboard', 'AFilter:" & tmAppliedFilters(ilLoop).sType & "', " & tmAppliedFilters(ilLoop).lValue & ", '" & tmAppliedFilters(ilLoop).sValue & "', " & tgUrf(0).iCode & ")"
            gSQLAndReturn slSQLQuery, False, llCount
        End If
    Next ilLoop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mLoadDashSettings
' Description:       Loads the Dashboard setting for the current user
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/5/2024-09:50:59
' Parameters :
'--------------------------------------------------------------------------------
Private Sub mLoadDashSettings()
    Dim slSQLQuery As String
    Dim llLongData As Long
    Dim slStringData As String
    Dim llCount As Long
    Dim ilColumn As Integer
    Dim rst As Recordset
    Dim ilSkip As Integer
    Dim slFilterType As String
    Dim slFilterString As String
    Dim ilLoop As Integer
    
    'Font Size
    slSQLQuery = "SELECT fcfCtrlIndex FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName = 'FontSize' and  fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        CboFontSize.ListIndex = Val(rst!fcfCtrlIndex)
    End If
    
    'Excel Headers
    slSQLQuery = "SELECT fcfCtrlIndex, fcfLongData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName= 'ExcelHeader' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        chkExcelHeader.Value = Val(rst!fcfLongData)
    End If
    
    'Show Checkboxes
    slSQLQuery = "SELECT fcfCtrlIndex, fcfLongData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName= 'Show' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    Do While Not rst.EOF
        If Val(rst!fcfCtrlIndex) > -1 And Val(rst!fcfCtrlIndex) < 5 Then
            ckcCntrType(Val(rst!fcfCtrlIndex)).Value = Val(rst!fcfLongData)
        End If
        rst.MoveNext
    Loop
        
    'Column Visible / ColumnSizes
    slSQLQuery = "SELECT fcfCtrlIndex, fcfLongData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName= 'ColumnSize' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    Do While Not rst.EOF
        If Val(rst!fcfCtrlIndex) <= grdCntr.cols - 1 Then
            imColumnWidths(Val(rst!fcfCtrlIndex)) = Val(rst!fcfLongData)
            grdCntr.ColWidth(Val(rst!fcfCtrlIndex)) = Val(rst!fcfLongData)
        End If
        rst.MoveNext
    Loop
    
    'Column Sort
    slSQLQuery = "SELECT fcfCtrlIndex, fcfLongData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName= 'SortColumn' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        imLastSortCol = Val(rst!fcfLongData)
    End If
    
    'Sort Direction
    slSQLQuery = "SELECT fcfCtrlIndex, fcfLongData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName= 'SortDir' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        imSortDir = Val(rst!fcfLongData)
    End If
    
    'Active Filters
    slSQLQuery = "SELECT fcfCtrlName, fcfLongData, fcfStringData FROM FCF_Form_Control WHERE fcfName = 'Dashboard' and fcfCtrlName like 'AFilter:%' and fcfUrfCode = " & tgUrf(0).iCode
    Set rst = gSQLSelectCall(slSQLQuery)
    Do While Not rst.EOF
        ilSkip = 0
        If InStr(1, rst!fcfCtrlName, "Active>") > 0 And InStr(1, txtCriteria.Text, "[Active>") > 0 Then ilSkip = 1
        If InStr(1, rst!fcfCtrlName, "*Active>") > 0 And InStr(1, txtCriteria.Text, "[*Active>") > 0 Then ilSkip = 1
        
        'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 1
        If InStr(1, rst!fcfCtrlName, "*SP") > 0 And InStr(1, txtCriteria.Text, "[*SP:") > 0 Then ilSkip = 1
        If InStr(1, rst!fcfCtrlName, "SP") > 0 And InStr(1, txtCriteria.Text, "[*SP:") > 0 Then ilSkip = 1
        
        If InStr(1, rst!fcfCtrlName, "Group") > 0 And InStr(1, txtCriteria.Text, "[Group:") > 0 Then ilSkip = 1
        
        'Obsolete terms
        If InStr(1, rst!fcfCtrlName, "ADate>") > 0 Then ilSkip = 1
        If InStr(1, rst!fcfCtrlName, "*ADate>") > 0 Then ilSkip = 1
        If InStr(1, rst!fcfCtrlName, "ADate<") > 0 Then ilSkip = 1
        
        If ilSkip = 0 Then
            slFilterType = Trim(Mid(rst!fcfCtrlName, 9))
            mManageAppliedFilters "Add", slFilterType, Val(rst!fcfLongData), Trim(rst!fcfStringData)
            slFilterType = mGetFilterName(slFilterType)
            slFilterString = mMakeFilterString(Trim(rst!fcfStringData), slFilterType)
            If Mid(txtCriteria.Text, Len(txtCriteria.Text), 1) <> " " Then txtCriteria.Text = txtCriteria.Text & " "
            txtCriteria.Text = txtCriteria.Text & slFilterString & " "
        End If
        rst.MoveNext
    Loop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mGetFilterName
' Description:       Gets the filter name for the provided FilterType (like Adv, returns Advertiser)
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/5/2024-09:51:47
' Parameters :       slFilterType (String) - Adv, Agy, Active>, etc.
'--------------------------------------------------------------------------------
Private Function mGetFilterName(slFilterType As String) As String
    Dim slString As String
    Select Case slFilterType
        Case "Adv": slString = "Advertiser"
        Case "Agy": slString = "Agency"
        Case "*Active>": slString = "Active on or after"
        Case "Active<": slString = "Active on or prior"
        Case "UDate>": slString = "Update Date (Beginning)"
        Case "UDate<": slString = "Update Date (Ending)"
        Case "Cntr": slString = "Contract Number"
        Case "CStat": slString = "Contract Status"
        Case "DlvyStat": slString = "Delivery Status"
        Case "SchStat": slString = "Schedule Status"
        Case "Prod": slString = "Product"
        Case "SP": slString = "Salesperson"
        Case "Group": slString = "Sales Group"
        Case "SO": slString = "Sales Office"
    End Select
    mGetFilterName = slString
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGetFiltersByType
' Description:       Get semicolon separated list of all filters for a given filter type
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       2/5/2024-14:41:29
' Parameters :       slFilerType (String) - Adv, Agy, etc.
'--------------------------------------------------------------------------------
Private Function mGetFiltersByType(slFilerType As String) As String
    Dim ilAppliedFiltersLoop As Integer
    Dim slString As String
    For ilAppliedFiltersLoop = 0 To UBound(tmAppliedFilters)
        If tmAppliedFilters(ilAppliedFiltersLoop).sType = slFilerType Then
            If slString <> "" Then slString = slString & "; "
            slString = slString & Trim(tmAppliedFilters(ilAppliedFiltersLoop).sValue)
        End If
    Next ilAppliedFiltersLoop
    mGetFiltersByType = slString
End Function
