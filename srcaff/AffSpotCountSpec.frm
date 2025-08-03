VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmSpotCountSpec 
   Caption         =   "Spot Count Tie-out"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7725
   Begin V81Affiliate.CSI_Calendar edcFeedStartDate 
      Height          =   270
      Left            =   1500
      TabIndex        =   1
      Top             =   75
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   476
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   2
   End
   Begin V81Affiliate.CSI_Calendar edcFeedEndDate 
      Height          =   300
      Left            =   5430
      TabIndex        =   3
      Top             =   75
      Width           =   1785
      _ExtentX        =   2143
      _ExtentY        =   661
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin VB.CheckBox ckcUserPosting 
      Caption         =   "Mixed User Posting"
      Height          =   240
      Left            =   5835
      TabIndex        =   7
      Top             =   615
      Width           =   1680
   End
   Begin VB.CheckBox ckcPartiallyPosted 
      Caption         =   "Partially Posted"
      Height          =   240
      Left            =   4365
      TabIndex        =   6
      Top             =   615
      Value           =   1  'Checked
      Width           =   1410
   End
   Begin VB.CheckBox ckcNotCompliant 
      Caption         =   "Not Compliant"
      Height          =   240
      Left            =   2985
      TabIndex        =   5
      Top             =   615
      Width           =   1305
   End
   Begin VB.CheckBox ckcOutBalance 
      Caption         =   "Spot Discrepancies"
      Height          =   240
      Left            =   1215
      TabIndex        =   4
      Top             =   615
      Value           =   1  'Checked
      Width           =   1740
   End
   Begin VB.OptionButton rbcSort 
      Caption         =   "Station, Vehicle"
      Height          =   225
      Index           =   1
      Left            =   3255
      TabIndex        =   12
      Top             =   1140
      Width           =   1530
   End
   Begin VB.OptionButton rbcSort 
      Caption         =   "Vehicle, Station"
      Height          =   225
      Index           =   0
      Left            =   1770
      TabIndex        =   11
      Top             =   1140
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CheckBox ckcInBalance 
      Caption         =   "Without any Issues"
      Height          =   240
      Left            =   3210
      TabIndex        =   9
      Top             =   870
      Width           =   1755
   End
   Begin VB.CheckBox ckcAllVehicles 
      Caption         =   "All Vehicles"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1740
      Width           =   1455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffSpotCountSpec.frx":0000
      Left            =   105
      List            =   "AffSpotCountSpec.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2085
      Width           =   4110
   End
   Begin VB.ListBox lbcStations 
      Height          =   3180
      ItemData        =   "AffSpotCountSpec.frx":0004
      Left            =   4800
      List            =   "AffSpotCountSpec.frx":0006
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   2085
      Width           =   2295
   End
   Begin VB.CheckBox ckcAllStations 
      Caption         =   "All Stations"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   1785
      Width           =   1455
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
      Height          =   330
      Left            =   735
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   5490
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcGetCount 
      Caption         =   "Get Counts"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2340
      TabIndex        =   19
      Top             =   5490
      Width           =   1245
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4110
      TabIndex        =   20
      Top             =   5505
      Width           =   1245
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1485
      Top             =   5475
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5910
      FormDesignWidth =   7725
   End
   Begin VB.CheckBox ckcWeb 
      Caption         =   "Include Web Spot Counts"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1425
      Value           =   1  'Checked
      Width           =   2220
   End
   Begin VB.CheckBox ckcCodeRow 
      Caption         =   "Include Extra Row for Internal Codes"
      Height          =   195
      Left            =   3390
      TabIndex        =   14
      Top             =   1425
      Width           =   3090
   End
   Begin VB.CheckBox ckcBreakOutBalance 
      Caption         =   "Break Discrepancies"
      Height          =   240
      Left            =   1215
      TabIndex        =   8
      Top             =   870
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.Label lacShow 
      Caption         =   "Row to Show"
      Height          =   180
      Left            =   135
      TabIndex        =   23
      Top             =   615
      Width           =   1080
   End
   Begin VB.Label lacSort 
      Caption         =   "Sort Major to Minor"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label lacDateComment 
      Caption         =   "Dates must be within same week (Mo-Su)"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   135
      TabIndex        =   22
      Top             =   360
      Width           =   3840
   End
   Begin VB.Label lacFeedStartdate 
      Caption         =   "Feed Start Date"
      Height          =   225
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lacFeedEndDate 
      Caption         =   "Feed End Date"
      Height          =   240
      Left            =   3990
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpotCountSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bmCkcAllIgnore As Boolean


Private Sub ckcAllStations_Click()
    Dim iValue As Integer
    Dim lErr As Long
    Dim lRg As Long
    Dim lRet As Long
    If bmCkcAllIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        bmCkcAllIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        bmCkcAllIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcAllVehicles_Click()
    Dim iValue As Integer
    Dim lErr As Long
    Dim lRg As Long
    Dim lRet As Long
    If bmCkcAllIgnore Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehicles.ListCount > 0 Then
        bmCkcAllIgnore = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        bmCkcAllIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcCancel_Click()
    Unload frmSpotCountSpec
End Sub

Private Sub cmcGetCount_Click()
    If edcFeedStartDate.Text = "" Then
        gMsgBox "Feed Start Date must be specified.", vbOKOnly
        Exit Sub
    End If
    If gIsDate(edcFeedStartDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid feed start date (m/d/yy).", vbCritical
        Exit Sub
    End If
    If edcFeedEndDate.Text = "" Then
        gMsgBox "Feed End Date must be specified.", vbOKOnly
        Exit Sub
    End If
    If gIsDate(edcFeedEndDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid feed end date (m/d/yy).", vbCritical
        Exit Sub
    End If
    If gDateValue(edcFeedEndDate.Text) < gDateValue(edcFeedStartDate.Text) Then
        Beep
        gMsgBox "Feed End Date must be on or after the Feed Start Date.", vbCritical
        Exit Sub
    End If
    If gWeekDayLong(gDateValue(edcFeedEndDate.Text)) < gWeekDayLong(gDateValue(edcFeedStartDate.Text)) Then
        Beep
        gMsgBox "Dates can not cross over into another week.", vbCritical
        Exit Sub
    End If
    If gDateValue(edcFeedEndDate.Text) - gDateValue(edcFeedStartDate.Text) > 6 Then
        Beep
        gMsgBox "Dates can not be more the one week.", vbCritical
        Exit Sub
    End If
    frmSpotCountGrid.Show vbModal
End Sub


Private Sub edcFeedEndDate_Change()
    mSetCommands
End Sub



Private Sub edcFeedStartDate_Change()
    mSetCommands
End Sub

Private Sub Form_Click()
    'cmcGetCount.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 1.5
    Me.Height = (Screen.Height) / 1.5
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    gCenterForm Me
    lacDateComment.Top = edcFeedStartDate.Top + edcFeedStartDate.Height
    lacDateComment.Left = edcFeedStartDate.Left
End Sub

Private Sub Form_Load()
    bmCkcAllIgnore = False
    mPopVehicles
    mPopStations
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSpotCountSpec = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub mPopStations()
    Dim ilLoop As Integer
    lbcStations.Clear
    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If (tgStationInfo(ilLoop).sUsedForATT = "Y") And (tgStationInfo(ilLoop).sAgreementExist = "Y") Then
            If tgStationInfo(ilLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(ilLoop).sCallLetters) & ", " & Trim$(tgStationInfo(ilLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(ilLoop).iCode
            End If
        End If
    Next ilLoop

End Sub

Private Sub mPopVehicles()
    Dim ilLoop
    lbcVehicles.Clear
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(ilLoop).iCode
    Next ilLoop
End Sub

Private Sub lbcStations_Click()
    If bmCkcAllIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        bmCkcAllIgnore = True
        ckcAllStations.Value = vbUnchecked
        bmCkcAllIgnore = False
    End If
    mSetCommands
End Sub

Private Sub lbcVehicles_Click()
    If bmCkcAllIgnore Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = vbChecked Then
        bmCkcAllIgnore = True
        ckcAllVehicles.Value = vbUnchecked
        bmCkcAllIgnore = False
    End If
    mSetCommands
End Sub
Private Sub mSetCommands()
    Dim blEnable As Boolean
    Dim ilLoop As Integer

    blEnable = False
    If (edcFeedStartDate.Text <> "") And (edcFeedEndDate.Text <> "") Then
        blEnable = True
        If lbcVehicles.SelCount <= 0 Then
            blEnable = False
        End If
        If lbcStations.SelCount <= 0 Then
            blEnable = False
        End If
    End If
    
    cmcGetCount.Enabled = blEnable
End Sub
