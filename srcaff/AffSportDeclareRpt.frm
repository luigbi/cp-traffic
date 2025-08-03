VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSportDeclareRpt 
   Caption         =   "Station Sports Declaration Report"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   375
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6480
      FormDesignWidth =   9360
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
      Height          =   4600
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8400
         Picture         =   "AffSportDeclareRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Select Stations from File.."
         Top             =   285
         Width           =   360
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3570
         Index           =   1
         ItemData        =   "AffSportDeclareRpt.frx":056A
         Left            =   4485
         List            =   "AffSportDeclareRpt.frx":056C
         Sorted          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ListBox lbcSeasons 
         Height          =   1035
         ItemData        =   "AffSportDeclareRpt.frx":056E
         Left            =   6855
         List            =   "AffSportDeclareRpt.frx":0570
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Frame frcDeclarations 
         Caption         =   "Declarations"
         Height          =   1140
         Left            =   105
         TabIndex        =   31
         Top             =   2715
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton optDeclare 
            Caption         =   "Summary"
            Height          =   255
            Index           =   2
            Left            =   75
            TabIndex        =   34
            Top             =   795
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.OptionButton optDeclare 
            Caption         =   "All Declarations"
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   33
            Top             =   225
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optDeclare 
            Caption         =   "Delinquent Only"
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   32
            Top             =   510
            Width           =   2895
         End
      End
      Begin VB.Frame frcClearance 
         Caption         =   "Clearance by"
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   1995
         Visible         =   0   'False
         Width           =   3180
         Begin VB.OptionButton optClearance 
            Caption         =   "Potential"
            Height          =   255
            Index           =   1
            Left            =   1410
            TabIndex        =   30
            Top             =   255
            Width           =   1100
         End
         Begin VB.OptionButton optClearance 
            Caption         =   "Delinquent"
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   29
            Top             =   225
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkSuppressDeclaration 
         Caption         =   "Suppress printing of current Carry Status"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1710
         Value           =   1  'Checked
         Width           =   3765
      End
      Begin VB.CheckBox chkPrintables 
         Caption         =   "Printables Only"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1410
         Width           =   1920
      End
      Begin VB.ListBox lbcStations 
         Height          =   3570
         ItemData        =   "AffSportDeclareRpt.frx":0572
         Left            =   6735
         List            =   "AffSportDeclareRpt.frx":0579
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   660
         Width           =   1980
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   6735
         TabIndex        =   20
         Top             =   285
         Width           =   1455
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3570
         Index           =   0
         ItemData        =   "AffSportDeclareRpt.frx":0580
         Left            =   4485
         List            =   "AffSportDeclareRpt.frx":0582
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   660
         Width           =   1995
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sort by"
         Height          =   885
         Left            =   3000
         TabIndex        =   14
         Top             =   3705
         Visible         =   0   'False
         Width           =   1260
         Begin VB.OptionButton optVehAff 
            Caption         =   "Vehicles"
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   15
            Top             =   270
            Width           =   1100
         End
         Begin VB.OptionButton optVehAff 
            Caption         =   "Stations"
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   16
            Top             =   555
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4485
         TabIndex        =   17
         Top             =   285
         Width           =   1320
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   270
         Left            =   3210
         TabIndex        =   11
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/13/2022"
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
      Begin V81Affiliate.CSI_Calendar CalEnterFrom 
         Height          =   270
         Left            =   1575
         TabIndex        =   12
         Top             =   840
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/13/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   1575
         TabIndex        =   10
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/13/2022"
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
      Begin V81Affiliate.CSI_Calendar CalEnterTo 
         Height          =   270
         Left            =   3210
         TabIndex        =   13
         Top             =   840
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/13/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin VB.Frame frcDates 
         Caption         =   "Agreements"
         Height          =   1065
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4230
         Begin VB.Label lacTo 
            Caption         =   "To"
            Height          =   225
            Left            =   2625
            TabIndex        =   25
            Top             =   615
            Width           =   375
         End
         Begin VB.Label lacFrom 
            Caption         =   "Entered- From"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   615
            Width           =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "End"
            Height          =   225
            Left            =   2625
            TabIndex        =   23
            Top             =   255
            Width           =   315
         End
         Begin VB.Label Label3 
            Caption         =   "Active- Start"
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   255
            Width           =   1035
         End
      End
      Begin VB.Label lacSeasons 
         Caption         =   "Seasons"
         Height          =   255
         Left            =   8085
         TabIndex        =   36
         Top             =   150
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lacVehicles 
         Caption         =   "Vehicles"
         Height          =   255
         Left            =   6090
         TabIndex        =   35
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5355
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   5115
      TabIndex        =   7
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
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
      Left            =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   4
         Left            =   2130
         TabIndex        =   5
         Top             =   1140
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         TabIndex        =   4
         Top             =   765
         Width           =   2040
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmSportDeclareRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmSportDeclareRpt - print the Sports Declaration by Station
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'

'****************************************************************************
Option Explicit

Private imChkStationIgnore As Integer
Private imChkListBoxIgnore As Integer
Private rst_att As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset
Private rst_Pet As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset
Dim imRptIndex As Integer                   'report option
Dim imVehAffIndex As Integer                'index into vehicle list box (0 or 1)

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    Dim iLoop As Integer
    
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
    If lbcVehAff(imVehAffIndex).ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff(imVehAffIndex).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff(imVehAffIndex).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False

    End If
    
    If lbcVehAff(imVehAffIndex).SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    End If
    
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllStations_Click()
 Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStationIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStationIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmSportDeclareRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim iRet As Integer
    Dim sCode As String
    Dim sName As String
    Dim sVehicles As String
    Dim sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim slEnterFrom As String
    Dim slEnterTo As String
    Dim slEnteredRange As String
    Dim slGenDate As String
    Dim slGenTime As String
    Dim slPrintSports As String
    Dim llHdVtfCode As Long
    Dim llFtVtfCode As Long
    Dim llLoopOnVtf As Long
    Dim llAttCode As Long
    Dim llLoopOnATT As Long
    ReDim llPrintAtt(0 To 0) As Long            'arry of agreements to update if printables only
    Dim ilVefCode As Integer
    Dim slAirDate As String                       'needed for adjustment of date/time for time zone
    Dim slAirTime As String
    Dim ilTimeAdj As Integer                     '+/- time adjustment for vehicle
    Dim llPetCode As Long
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String
    '12/12/14
    Dim slFed As String
        
        On Error GoTo ErrHand
        
        '11-8-12 Active start dates for user entry is for the Sports Clearance only.
        'Changed to use Season for the Sports Declaration report
        If imRptIndex = SPORTDECLARE_Rpt Then                'sports declaration
            'use the start/end dates of the season selected
            For i = 0 To lbcSeasons.ListCount - 1 Step 1
                If lbcSeasons.Selected(i) Then
                    'get the start/end dates from the season info table
                    CalOnAirDate.Text = Format$(tgSeasonInfo(i).lStartDate, "m/d/yy")
                    calOffAirDate.Text = Format$(tgSeasonInfo(i).lEndDate, "m/d/yy")
                    Exit For
                End If
            Next i
        End If
        'Clearance reports uses the dates entered by user
        sStartDate = Trim$(CalOnAirDate.Text)
        If sStartDate = "" Then
            sStartDate = "1/1/1970"
        End If
        sEndDate = Trim$(calOffAirDate.Text)
        If sEndDate = "" Then
            sEndDate = "12/31/2069"
        End If
        If gIsDate(sStartDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalOnAirDate.SetFocus
            Exit Sub
        End If
        If gIsDate(sEndDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            calOffAirDate.SetFocus
            Exit Sub
        End If
    
        'Validate Entered From/To dates
        slEnterFrom = Trim$(CalEnterFrom.Text)
        If slEnterFrom = "" Then
            slEnterFrom = "1/1/1970"
        End If
        slEnterTo = Trim$(CalEnterTo.Text)
        If slEnterTo = "" Then
            slEnterTo = "12/31/2069"
        End If
        If gIsDate(slEnterFrom) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalEnterFrom.SetFocus
            Exit Sub
        End If
        If gIsDate(slEnterTo) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalEnterTo.SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
      
        If optRptDest(0).Value = True Then
            ilRptDest = 0
        ElseIf optRptDest(1).Value = True Then
            ilRptDest = 1
        ElseIf optRptDest(2).Value = True Then
            ilRptDest = 2
            ilExportType = cboFileType.ListIndex    '3-15-04
        End If
            
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False

        gUserActivityLog "S", sgReportListName & ": Prepass"
        
        sStartDate = Format(sStartDate, "m/d/yyyy")
        sEndDate = Format(sEndDate, "m/d/yyyy")
       
        'create sql query to get agreements active between 2 spans
        sDateRange = " attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attDropDate >=" & "'" + Format$(sStartDate, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'"
        slEnteredRange = " attEnterDate >= " & "'" & Format$(slEnterFrom, sgSQLDateForm) & "'" & " And attEnterDate <= " & "'" & Format$(slEnterTo, sgSQLDateForm) & "'"
        
        sVehicles = ""
        sStations = ""
        
        If chkListBox.Value = vbUnchecked Then
            'User did NOT select all vehicles
            For i = 0 To lbcVehAff(imVehAffIndex).ListCount - 1 Step 1
                If lbcVehAff(imVehAffIndex).Selected(i) Then
                    ilVefCode = lbcVehAff(imVehAffIndex).ItemData(i)                'this applies to Sports Declaration only since its a single vehicle selection
                                                                                    'need this to get the text for home vs away text.  Crystal seems to not be able to handle
                                                                                    'the 2 places needed to show the information with subreports
                    If Len(sVehicles) = 0 Then
                        sVehicles = "(vefCode = " & lbcVehAff(imVehAffIndex).ItemData(i) & ")"
                    Else
                        sVehicles = sVehicles & " OR (vefCode = " & lbcVehAff(imVehAffIndex).ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
    
        'User did NOT select all stations
        If ckcAllStations.Value = vbUnchecked Then    '= 0 Then
            'User did NOT select all stations
            For i = 0 To lbcStations.ListCount - 1 Step 1
                If lbcStations.Selected(i) Then
                    If Len(sStations) = 0 Then
                        sStations = "(attShfCode = " & lbcStations.ItemData(i) & ")"
                    Else
                        sStations = sStations & " OR (attShfCode = " & lbcStations.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
        'End If
        
        slPrintSports = ""          'games only
        If imRptIndex = SPORTDECLARE_Rpt Then                'sports declaration
            If chkPrintables.Value = vbChecked Then         'printables only
                slPrintSports = " and (veftype = 'G' and attPetPrinted <> 'Y') "
            Else
                slPrintSports = " and (vefType = 'G') "
            End If
            
            gGetEventTitles ilVefCode, slEventTitle1, slEventTitle2
            sgCrystlFormula4 = "'" & slEventTitle1 & "'"         'visiting
            sgCrystlFormula5 = "'" & slEventTitle2 & "'"         'home
            
        Else                                            'sports clearance
            slPrintSports = " and (vefType = 'G') "
        End If
    
        
'        'hidden question, currently defaulted to Station
'        If optVehAff(0).Value = True Then   'VEHICLE SORT
'
'            SQLQuery = "SELECT * From shtt INNER JOIN att att ON shttCode = att.attShfCode "
'            'SQLQuery = SQLQuery + "INNER JOIN mkt ON shttMktCode = mktCode "
'            '4-12-12 market is optional
'            SQLQuery = SQLQuery + "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
'            SQLQuery = SQLQuery + " INNER JOIN   VEF_Vehicles ON attVefCode = vefCode "
'
'
'            SQLQuery = SQLQuery + " Where (" & sDateRange & ")" & " and (" & slEnteredRange & ")"
'
'            SQLQuery = SQLQuery + " AND ( shttType = 0 ) " & slPrintSports
'            If sVehicles <> "" Then
'                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
'            End If
'            If sStations <> "" Then '12-13-00
'                SQLQuery = SQLQuery + " AND (" & sStations & ")"
'            End If
'
'            SQLQuery = SQLQuery + " ORDER BY vefName, shttCallLetters"
'            'End If
'        Else            'STATION SORT
               
            SQLQuery = "SELECT * From shtt INNER JOIN att ON shttCode = attShfCode "
            SQLQuery = SQLQuery & "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
            SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles ON attVefCode = vefCode "
    
            SQLQuery = SQLQuery + " Where (" & sDateRange & ")" & " and (" & slEnteredRange & ")"
            SQLQuery = SQLQuery + " AND ( shttType = 0 ) and attServiceAgreement <> 'Y' " & slPrintSports
          
            If sStations <> "" Then
                SQLQuery = SQLQuery + " AND (" & sStations & ")"
            End If
            If sVehicles <> "" Then     '12-13-00
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
'        End If
                             
        slGenDate = Format$(gNow(), "m/d/yyyy")
        slGenTime = Format$(gNow(), sgShowTimeWSecForm)
        Set rst_att = gSQLSelectCall(SQLQuery)
            While Not rst_att.EOF
                ilVefCode = rst_att!attvefCode
                llHdVtfCode = 0
                llFtVtfCode = 0
                ilTimeAdj = gGetTimeAdj(rst_att!attshfcode, rst_att!attvefCode, slFed)

                'dont need the header and footer fields for the sports clearance
                If imRptIndex = SPORTDECLARE_Rpt Then
                    'find the header and footer notes
                    For llLoopOnVtf = LBound(tgVtfInfo) To UBound(tgVtfInfo) - 1
                        If tgVtfInfo(llLoopOnVtf).iVefCode = rst_att!attvefCode Then
                            If tgVtfInfo(llLoopOnVtf).sType = "H" Then
                                llHdVtfCode = tgVtfInfo(llLoopOnVtf).lCode
                            Else
                                If tgVtfInfo(llLoopOnVtf).sType = "F" Then
                                    llFtVtfCode = tgVtfInfo(llLoopOnVtf).lCode
                                End If
                            End If
                        End If
                        If tgVtfInfo(llLoopOnVtf).iVefCode > rst_att!attvefCode Then              'the info has been sorted by vefcode, so exit if entry is a higher internal code
                            Exit For
                        End If
                            
                    Next llLoopOnVtf
                    llAttCode = rst_att!attCode
'                    SQLQuery = "INSERT INTO " & "CPR_Copy_report "
'                    SQLQuery = SQLQuery & " (cprcntrno, cprHd1CefCode, cprFt1CefCode, cprFt2CefCode, "      'attcode, hd vtf code, ft vtf code, sitecode
'                    SQLQuery = SQLQuery & " cprGendate, cprGenTime) "
'
'                    SQLQuery = SQLQuery & " VALUES (" & llAttCode & ", " & llHdVtfCode & ", " & llFtVtfCode & ", " & 1 & ", "
'                    SQLQuery = SQLQuery & "'" & Format$(slGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"   '", "
'
'                    cnn.BeginTrans
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHand:
'                    End If
'                    cnn.CommitTrans

                    SQLQuery = "Select * from pet where petattcode = " & llAttCode
                    Set rst_Pet = gSQLSelectCall(SQLQuery)
                    While Not rst_Pet.EOF           'create a record for each game in this agreement
                        llPetCode = rst_Pet!petCode
                        If imRptIndex = SPORTDECLARE_Rpt Then           '11-8-12
                            SQLQuery = "Select gsfairtime,gsfairdate from gsf_Game_Schd left outer join ghf_game_Header on gsfghfcode = ghfcode  where gsfcode = " & rst_Pet!petGsfCode & " and gsfAirDate >= " & " '" + Format$(sStartDate, sgSQLDateForm) & "'" & " And gsfAirDate <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'"
                        Else                    'sports clearance
                            SQLQuery = "Select gsfairtime,gsfairdate from gsf_Game_Schd where gsfcode = " & rst_Pet!petGsfCode
                        End If
                        Set rst_Gsf = gSQLSelectCall(SQLQuery)
                        
                        While Not rst_Gsf.EOF
                            slAirTime = Format$(CStr(rst_Gsf!gsfAirTime), sgShowTimeWSecForm)
                            slAirDate = Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
                            gAdjustEventTime ilTimeAdj, slAirDate, slAirTime
    
                            slAirDate = Format$(slAirDate, sgSQLDateForm)
                            SQLQuery = "INSERT INTO " & "CPR_Copy_report "
                            SQLQuery = SQLQuery & " (cprcntrno, cprHd1CefCode, cprFt1CefCode, cprFt2CefCode, cprLen, "      'attcode, hd vtf code, ft vtf code, petcode, sitecode
                            SQLQuery = SQLQuery & " cprSpotDate, cprSpotTime, "
                            SQLQuery = SQLQuery & " cprGendate, cprGenTime) "
                            
                            SQLQuery = SQLQuery & " VALUES (" & llAttCode & ", " & llHdVtfCode & ", " & llFtVtfCode & ", " & llPetCode & ", " & 1 & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "',"
    
                            SQLQuery = SQLQuery & "'" & Format$(slGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"   '", "
                            
                            cnn.BeginTrans
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SportDeclareRpt"
                                cnn.RollbackTrans
                                Exit Sub
                            End If
                            cnn.CommitTrans
                        rst_Gsf.MoveNext
                        Wend

                    rst_Pet.MoveNext
                    Wend
                    
                    If chkPrintables.Value = vbChecked Then
                        llPrintAtt(UBound(llPrintAtt)) = llAttCode
                        ReDim Preserve llPrintAtt(LBound(llPrintAtt) To (UBound(llPrintAtt) + 1)) As Long
                    End If
                    
                Else                        'get the game information
                    llAttCode = rst_att!attCode
                    llHdVtfCode = 0
                    llFtVtfCode = 0
                    SQLQuery = "Select * from pet where petattcode = " & llAttCode
                    Set rst_Pet = gSQLSelectCall(SQLQuery)
                    While Not rst_Pet.EOF           'create a record for each game in this agreement
                        'llHdVtfCode = rst_Pet!petCode
                        llPetCode = rst_Pet!petCode
                        '8-1-14 add valid air date selection
                        SQLQuery = "Select gsfairtime,gsfairdate from gsf_Game_Schd where gsfcode = " & rst_Pet!petGsfCode & " and gsfAirDate >= " & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And gsfAirDate <= " & "'" & Format$(sEndDate, sgSQLDateForm) & "'"
                        Set rst_Gsf = gSQLSelectCall(SQLQuery)
                        
                        While Not rst_Gsf.EOF
                            slAirTime = Format$(CStr(rst_Gsf!gsfAirTime), sgShowTimeWSecForm)
                            slAirDate = Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
                            gAdjustEventTime ilTimeAdj, slAirDate, slAirTime
    
                            slAirDate = Format$(slAirDate, sgSQLDateForm)
                            SQLQuery = "INSERT INTO " & "CPR_Copy_report "
                            SQLQuery = SQLQuery & " (cprcntrno, cprFt2CefCode,  "      'attcode & petcode
                            SQLQuery = SQLQuery & " cprSpotDate, cprSpotTime, "
                            SQLQuery = SQLQuery & " cprGendate, cprGenTime) "
                            
                            SQLQuery = SQLQuery & " VALUES (" & llAttCode & ", " & llPetCode & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "',"
    
                            SQLQuery = SQLQuery & "'" & Format$(slGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"   '", "
                            
                            cnn.BeginTrans
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SportDeclareRpt-cmdReport_Click"
                                cnn.RollbackTrans
                                Exit Sub
                            End If
                            cnn.CommitTrans
                        rst_Gsf.MoveNext
                        Wend

                    rst_Pet.MoveNext
                    Wend
                End If
                
'                llAttCode = rst_att!attCode
'                SQLQuery = "INSERT INTO " & "CPR_Copy_report "
'                SQLQuery = SQLQuery & " (cprcntrno, cprHd1CefCode, cprFt1CefCode, cprFt2CefCode, "      'attcode, hd vtf code, ft vtf code, sitecode
'                SQLQuery = SQLQuery & " cprGendate, cprGenTime) "
'
'                SQLQuery = SQLQuery & " VALUES (" & llAttCode & ", " & llHdVtfCode & ", " & llFtVtfCode & ", " & 1 & ", "
'                SQLQuery = SQLQuery & "'" & Format$(slGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"   '", "
'
'                cnn.BeginTrans
'                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                    GoSub ErrHand:
'                End If
'                cnn.CommitTrans
'
'                If chkPrintables.Value = vbChecked And imRptIndex = SPORTDECLARE_Rpt Then
'                    llPrintAtt(UBound(llPrintAtt)) = llAttCode
'                    ReDim Preserve llPrintAtt(LBound(llPrintAtt) To (UBound(llPrintAtt) + 1)) As Long
'                End If
                
            rst_att.MoveNext
        Wend

    gUserActivityLog "E", sgReportListName & ": Prepass"
    
        If imRptIndex = SPORTDECLARE_Rpt Then
            slRptName = "AfSportDeclare.rpt"
            slExportName = "SportDeclare"
            sgCrystlFormula1 = "'N'"
            If chkSuppressDeclaration.Value = vbChecked Then
                sgCrystlFormula1 = "'Y'"
            End If
            
'            SQLQuery = "SELECT shttcallletters, vefname, spfGclient, spfGaddr1, spfGAddr2, spfGAddr3, vefcode, shttState, shttCity, Vtf_Vehicle_Text.vtfText, Vtf_FtVehicle_Text.vtfText "
'            SQLQuery = SQLQuery & " FROM CPR_Copy_report INNER JOIN att on cprcntrno = attcode INNER JOIN spf_Site_Options spf_Site_Options on cprFt2Cefcode = spfcode "
'            SQLQuery = SQLQuery & " LEFT OUTER JOIN VTF_Vehicle_Text VTF_Vehicle_text on cprhd1CefCode = VTF_Vehicle_text.vtfcode LEFT OUTER JOIN  VTF_VEhicle_Text VTF_FtVehicle_Text on cprFt1Cefcode = vtf_FtVEhicle_Text.vtfcode "
'            SQLQuery = SQLQuery & " INNER JOIN shtt on attshfcode = shttcode INNER JOIN VEF_Vehicles on attVefcode = Vefcode "
'            SQLQuery = SQLQuery + " WHERE cprGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND cprGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "'"
'
            SQLQuery = "SELECT shttcallletters, shttState, shttCity, gsfAirDate, gsfAirTime, gsfGameNo, petClearStatus, petDeclaredStatus, mnf_multi_names.mnfName, mnf_VisitMulti_Names.mnfName, mnf_LangMulti_Names.mnfName, vffPledgeClearance,  vefname, vefcode,spfGclient, spfGaddr1, spfGAddr2, spfGAddr3,Vtf_Vehicle_Text.vtfText, Vtf_FtVehicle_Text.vtfText  "
            SQLQuery = SQLQuery & " FROM CPR_Copy_report INNER JOIN att on cprcntrno = attcode INNER JOIN spf_Site_Options spf_Site_Options on cprLen = spfcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN VTF_Vehicle_Text VTF_Vehicle_text on cprhd1CefCode = VTF_Vehicle_text.vtfcode LEFT OUTER JOIN  VTF_VEhicle_Text VTF_FtVehicle_Text on cprFt1Cefcode = vtf_FtVEhicle_Text.vtfcode "
            SQLQuery = SQLQuery & " INNER JOIN vef_vehicles on attvefcode = vefcode "
            SQLQuery = SQLQuery & " INNER JOIN pet on cprft2Cefcode = petcode "
            SQLQuery = SQLQuery & " INNER JOIN  gsf_Game_Schd on petgsfcode = gsfcode "
            SQLQuery = SQLQuery & " INNER JOIN  ghf_Game_Header on gsfghfcode = ghfcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN Vff_Vehicle_Features on vefcode = vffvefcode "
            SQLQuery = SQLQuery & " INNER JOIN shtt on attshfcode = shttcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names mnf_Multi_Names on gsfhomemnfCode = mnf_Multi_Names.mnfcode LEFT OUTER JOIN  mnf_Multi_Names mnf_VisitMulti_Names on gsfVisitmnfCode = mnf_VisitMulti_Names.mnfcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN  mnf_Multi_Names mnf_LangMulti_Names on gsfLangmnfCode = mnf_LangMulti_Names.mnfcode "
            SQLQuery = SQLQuery + " WHERE cprGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND cprGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "'"


        Else
            slRptName = "AfSportClear.rpt"
            slExportName = "SportClear"
            
            'sgCrystlFormula2 is todays date to show diff/discreps in the past
            'Delinquent or Potential option
            If optClearance(0).Value = True Then          'Delinquent
                sgCrystlFormula3 = "'D'"
                sgCrystlFormula1 = "'A'"                        'show all declarations or delinq(difference)
                If optDeclare(1).Value = True Then
                    sgCrystlFormula1 = "'D'"                        'show delinquent or differences only
                End If
            Else
              sgCrystlFormula3 = "'P'"                      'Potential
                If optDeclare(0).Value = True Then            'All
                    sgCrystlFormula1 = "'A'"                        'show all declarations or delinq(difference)
                ElseIf optDeclare(1).Value = True Then
                    sgCrystlFormula1 = "'D'"                        'show delinquent or differences only
                Else
                    sgCrystlFormula1 = "'S'"                        'summary, total line only
                End If
            End If
                        
            SQLQuery = "SELECT shttcallletters,gsfAirDate, gsfAirTime, gsfGameNo, petClearStatus, petDeclaredStatus, mnf_multi_names.mnfName, mnf_VisitMulti_Names.mnfName, vffPledgeClearance,  vefname, vefcode "
            SQLQuery = SQLQuery & " FROM CPR_Copy_report INNER JOIN att on cprcntrno = attcode "
            SQLQuery = SQLQuery & " INNER JOIN vef_vehicles on attvefcode = vefcode "
            SQLQuery = SQLQuery & " INNER JOIN pet on cprft2Cefcode = petcode "
            SQLQuery = SQLQuery & " INNER JOIN  gsf_Game_Schd on petgsfcode = gsfcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN Vff_Vehicle_Features on vefcode = vffvefcode "
            SQLQuery = SQLQuery & " INNER JOIN shtt on attshfcode = shttcode "
            SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names mnf_Multi_Names on gsfhomemnfCode = mnf_Multi_Names.mnfcode LEFT OUTER JOIN  mnf_Multi_Names mnf_VisitMulti_Names on gsfVisitmnfCode = mnf_VisitMulti_Names.mnfcode "
            SQLQuery = SQLQuery + " WHERE cprGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND cprGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "'"

        End If
    
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
        SQLQuery = ""
        SQLQuery = "DELETE FROM CPR_Copy_Report "
        SQLQuery = SQLQuery & " WHERE (cprGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' " & "and cprGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"
        cnn.BeginTrans
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SportDeclareRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        
        'if Printables only, update the flag
        If chkPrintables.Value = vbChecked And imRptIndex = SPORTDECLARE_Rpt Then
            For llLoopOnATT = LBound(llPrintAtt) To UBound(llPrintAtt) - 1
                SQLQuery = "UPDATE Att SET attPetPrinted = " & "'Y'" & " WHERE attcode = " & llPrintAtt(llLoopOnATT)
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHandUpdate:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "SportDeclareRpt-cmdReport_Click"
                    Exit Sub
                End If
            Next llLoopOnATT
        End If
    
        On Error Resume Next
        rst_att.Close
        rst_Shtt.Close
        rst_Pet.Close
        rst_Gsf.Close

        Erase llPrintAtt
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True

        Screen.MousePointer = vbDefault
        Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotDeclareRpt-cmdReport"
    Exit Sub
ErrReturn:
    
    
ErrHandUpdate:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotDeclareRpt-cmdReport"
    Exit Sub
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmSportDeclareRpt
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
    gSelectiveStationsFromImport lbcStations, ckcAllStations, Trim$(CommonDialog1.fileName)
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
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmSportDeclareRpt
    gCenterForm frmSportDeclareRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim dDelinqDate As Date
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    imRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    If imRptIndex = SPORTDECLARE_Rpt Then
        frmSportDeclareRpt.Caption = "Station Sports Declaration Report - " & sgClientName
        chkPrintables.Move 120, 1410
        chkSuppressDeclaration.Move 120, 1710
        frcClearance.Visible = False
        frcDeclarations.Visible = False
        chkListBox.Visible = False
        lacVehicles.Move 4485, 285
        lacVehicles.Visible = True

        lacSeasons.Move 6735, 285
        lacSeasons.Visible = True
        lbcSeasons.Move 6735, lbcVehAff(imVehAffIndex).Top, lbcVehAff(imVehAffIndex).Width, 2000
        lbcSeasons.Visible = True
        ckcAllStations.Top = lbcSeasons.Top + lbcSeasons.Height + 30
        'TTP 9943
        cmdStationListFile.Top = ckcAllStations.Top - 50
        lbcStations.Top = ckcAllStations.Top + ckcAllStations.Height + 30
        lbcStations.Height = 2000
        imVehAffIndex = 1
        lbcVehAff(0).Visible = False            'multi vehicle selection
        lbcVehAff(1).Visible = True             'single vehicle selection
        frcDates.Visible = False
        CalOnAirDate.Visible = False
        calOffAirDate.Visible = False
        CalEnterFrom.Visible = False
        CalEnterTo.Visible = False
        chkPrintables.Move 120, 240
        chkSuppressDeclaration.Move 120, chkPrintables.Top + chkPrintables.Height + 30
    Else
        frmSportDeclareRpt.Caption = "Sports Clearance Report - " & sgClientName
        chkPrintables.Visible = False
        chkSuppressDeclaration.Visible = False
        frcClearance.Move 120, 1410
        frcDeclarations.Move 120, frcClearance.Top + frcClearance.Height + 30
        frcClearance.Visible = True
        frcDeclarations.Visible = True
        imVehAffIndex = 0
        lbcVehAff(0).Visible = True            'multi vehicle selection
        lbcVehAff(1).Visible = False             'single vehicle selection
        
    End If
    
    imChkListBoxIgnore = False
    
    slDate = Format$(gNow(), "m/d/yyyy")
    'send todays date to crystal for Sports Clearance
    dDelinqDate = CDate(slDate)
    sgCrystlFormula2 = "Date(" + Format$(dDelinqDate, "yyyy") + "," + Format$(dDelinqDate, "mm") + "," + Format$(dDelinqDate, "dd") + ")"
    
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = slDate
    calOffAirDate.Text = DateAdd("d", 6, slDate)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        If (tgVehicleInfo(iLoop).sVehType) = "G" Then
            lbcVehAff(imVehAffIndex).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff(imVehAffIndex).ItemData(lbcVehAff(imVehAffIndex).NewIndex) = tgVehicleInfo(iLoop).iCode
        End If
    Next iLoop
    
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    
    CalEnterFrom.ZOrder (0)
    CalOnAirDate.ZOrder (0)
    CalEnterTo.ZOrder (0)
    calOffAirDate.ZOrder (0)

    lbcStations.Clear
    
    'dont show all stations unless more than 1 vehicle is selected; otherwise, only show those stations that have an agreement with the vehicle
'    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
'            lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
'            lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
'        End If
'    Next iLoop
    
    ilRet = gPopVtf()               'obtain the vehicle text info

    gPopExportTypes cboFileType     '3-15-04
    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_att.Close
    rst_Shtt.Close
    rst_Pet.Close
    rst_Gsf.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmSportDeclareRpt = Nothing
End Sub



Private Sub lbcStations_Click()
If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        imChkStationIgnore = True
        'chkListBox.Value = False
        ckcAllStations.Value = 0    'chged from false to 0 10-22-99
        imChkStationIgnore = False
    End If
End Sub

Private Sub lbcVehAff_Click(imVehAff As Integer)
Dim iLoop As Integer
Dim ilVefCode As Integer
Dim llShfCode As Long
Dim ilRet As Integer

    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
    If lbcVehAff(imVehAffIndex).SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    Else
        'only one, show only the ones with agreements
        lbcStations.Clear
        ilVefCode = lbcVehAff(imVehAffIndex).ItemData(lbcVehAff(imVehAffIndex).ListIndex)
        SQLQuery = "Select distinct attshfcode from att where attvefcode = " & ilVefCode
        Set rst_Shtt = gSQLSelectCall(SQLQuery)
        While Not rst_Shtt.EOF
            llShfCode = gBinarySearchStationInfoByCode(rst_Shtt!attshfcode)
            If llShfCode <> -1 Then
                lbcStations.AddItem Trim$(tgStationInfoByCode(llShfCode).sCallLetters) & ", " & Trim$(tgStationInfoByCode(llShfCode).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfoByCode(llShfCode).iCode
            End If

        rst_Shtt.MoveNext
        Wend
    End If
    ckcAllStations.Value = vbUnchecked
    If imRptIndex = SPORTDECLARE_Rpt Then
        ilRet = gPopSeasons(ilVefCode)
        lbcSeasons.Clear
        For iLoop = 0 To UBound(tgSeasonInfo) - 1 Step 1
            lbcSeasons.AddItem Trim$(tgSeasonInfo(iLoop).sName)
            lbcSeasons.ItemData(lbcSeasons.NewIndex) = tgSeasonInfo(iLoop).lGhfCode
        Next iLoop
    End If

End Sub

Private Sub optClearance_Click(Index As Integer)
    If Index = 0 Then
        optDeclare(1).Caption = "Delinquent Only"
        optDeclare(2).Visible = False
    Else
        optDeclare(1).Caption = "Delinquent/Differences Only"
        optDeclare(2).Visible = True
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub
