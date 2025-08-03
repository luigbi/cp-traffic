VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCommentRpt 
   Caption         =   "Contact Comment Report"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
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
      Height          =   5070
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   7545
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6960
         Picture         =   "AffCommentRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Select Stations from File.."
         Top             =   165
         Width           =   360
      End
      Begin VB.CheckBox ckcInclID 
         Caption         =   "Include Internal Comment ID"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   4770
         Width           =   2370
      End
      Begin VB.TextBox txtMatchingComment 
         Height          =   285
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   46
         Top             =   4440
         Width           =   2325
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   3060
         TabIndex        =   30
         Top             =   255
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         backcolor       =   -2147483643
         forecolor       =   -2147483640
         borderstyle     =   1
         csi_showdropdownonfocus=   0
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   51200
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffCommentRpt.frx":056A
         csi_daynamefont =   "AffCommentRpt.frx":0596
         csi_monthnamefont=   "AffCommentRpt.frx":05C4
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         backcolor       =   -2147483643
         forecolor       =   -2147483640
         borderstyle     =   1
         csi_showdropdownonfocus=   0
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   51200
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffCommentRpt.frx":05F2
         csi_daynamefont =   "AffCommentRpt.frx":061E
         csi_monthnamefont=   "AffCommentRpt.frx":064C
      End
      Begin V81Affiliate.CSI_Calendar CalFollowEnd 
         Height          =   285
         Left            =   3060
         TabIndex        =   32
         Top             =   630
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         backcolor       =   -2147483643
         forecolor       =   -2147483640
         borderstyle     =   1
         csi_showdropdownonfocus=   0
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   51200
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffCommentRpt.frx":067A
         csi_daynamefont =   "AffCommentRpt.frx":06A6
         csi_monthnamefont=   "AffCommentRpt.frx":06D4
      End
      Begin V81Affiliate.CSI_Calendar CalFollowStart 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   630
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         backcolor       =   -2147483643
         forecolor       =   -2147483640
         borderstyle     =   1
         csi_showdropdownonfocus=   0
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   51200
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffCommentRpt.frx":0702
         csi_daynamefont =   "AffCommentRpt.frx":072E
         csi_monthnamefont=   "AffCommentRpt.frx":075C
      End
      Begin VB.ListBox lbcDept 
         Height          =   2010
         ItemData        =   "AffCommentRpt.frx":078A
         Left            =   4095
         List            =   "AffCommentRpt.frx":078C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   2760
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame plctotals 
         Caption         =   "Totals by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   39
         Top             =   3675
         Width           =   3360
         Begin VB.OptionButton rvcTotalsBy 
            Caption         =   "Summary"
            Height          =   255
            Index           =   1
            Left            =   1230
            TabIndex        =   41
            Top             =   240
            Width           =   1240
         End
         Begin VB.OptionButton rbctotalsBy 
            Caption         =   "Detail"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame plcFollowUp 
         Caption         =   "Include Follow-Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   3360
         Begin VB.CheckBox ckcFollowUp 
            Caption         =   "Undone"
            Height          =   255
            Index           =   1
            Left            =   1005
            TabIndex        =   38
            Top             =   240
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.CheckBox ckcFollowUp 
            Caption         =   "Done"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   800
         End
      End
      Begin VB.Frame plcSelectWho 
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
         Height          =   570
         Left            =   120
         TabIndex        =   35
         Top             =   2415
         Width           =   3360
         Begin VB.OptionButton rbcWho 
            Caption         =   "All"
            Height          =   225
            Index           =   2
            Left            =   2595
            TabIndex        =   45
            Top             =   240
            Width           =   555
         End
         Begin VB.OptionButton rbcWho 
            Caption         =   "Department"
            Height          =   225
            Index           =   1
            Left            =   1170
            TabIndex        =   44
            Top             =   240
            Width           =   1290
         End
         Begin VB.OptionButton rbcWho 
            Caption         =   "Person"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.CheckBox ckcSkip3 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3210
         TabIndex        =   14
         Top             =   2085
         Width           =   255
      End
      Begin VB.CheckBox ckcSkip2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3210
         TabIndex        =   12
         Top             =   1665
         Width           =   255
      End
      Begin VB.CheckBox ckcSkip1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3210
         TabIndex        =   10
         Top             =   1260
         Width           =   255
      End
      Begin VB.ComboBox cbcSort3 
         Height          =   315
         Left            =   1170
         TabIndex        =   13
         Top             =   2085
         Width           =   1560
      End
      Begin VB.ComboBox cbcSort2 
         Height          =   315
         Left            =   1170
         TabIndex        =   11
         Top             =   1665
         Width           =   1560
      End
      Begin VB.ComboBox cbcSort1 
         Height          =   315
         Left            =   1170
         TabIndex        =   9
         Top             =   1260
         Width           =   1560
      End
      Begin VB.CheckBox chkAllWho 
         Caption         =   "All Persons"
         Height          =   255
         Left            =   4110
         TabIndex        =   15
         Top             =   2415
         Width           =   1575
      End
      Begin VB.CheckBox ckcStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   5820
         TabIndex        =   23
         Top             =   165
         Width           =   1455
      End
      Begin VB.ListBox lbcPersons 
         Height          =   2010
         ItemData        =   "AffCommentRpt.frx":078E
         Left            =   4095
         List            =   "AffCommentRpt.frx":0795
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   2760
         Width           =   3315
      End
      Begin VB.ListBox lbcStations 
         Height          =   1815
         ItemData        =   "AffCommentRpt.frx":079C
         Left            =   5820
         List            =   "AffCommentRpt.frx":079E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1600
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1815
         ItemData        =   "AffCommentRpt.frx":07A0
         Left            =   4095
         List            =   "AffCommentRpt.frx":07A2
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1600
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4095
         TabIndex        =   17
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label lacMatchingComment 
         Caption         =   "Comment contains"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   4455
         Width           =   1530
      End
      Begin VB.Label lacFollowEnd 
         Caption         =   "End:"
         Height          =   255
         Left            =   2625
         TabIndex        =   34
         Top             =   645
         Width           =   525
      End
      Begin VB.Label lacFollowStart 
         Caption         =   "Followup Dates-Start:"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label lacSkip 
         Alignment       =   2  'Center
         Caption         =   "Page skip"
         Height          =   255
         Left            =   2865
         TabIndex        =   28
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lacSort3 
         Caption         =   "Sort Field #3"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2115
         Width           =   1185
      End
      Begin VB.Label lacSort2 
         Caption         =   "Sort Field #2"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1695
         Width           =   1185
      End
      Begin VB.Label lacSort1 
         Caption         =   "Sort Field #1"
         Height          =   255
         Left            =   105
         TabIndex        =   25
         Top             =   1290
         Width           =   1185
      End
      Begin VB.Label lacSortSeq 
         Caption         =   "Enter sort sequence (major to minor):"
         Height          =   255
         Left            =   105
         TabIndex        =   24
         Top             =   990
         Width           =   2880
      End
      Begin VB.Label Label3 
         Caption         =   "Post Dates-Start:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "End:"
         Height          =   255
         Left            =   2625
         TabIndex        =   8
         Top             =   255
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   21
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   20
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   19
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmCommentRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmCommentRpt - List of spots aired for vehicles and/or stations
'*                if the spot doesnt exist in AST, do not go out to
'*                retrieve it from the LST.  Also, include only those
'*                spots that have been imported or posted
'*
'*  Created 7/30/03 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'
'*      8-11-04 Add option to select by stations
'               Fix selectivity by Advertiser
'******************************************************
Option Explicit
Private rst_cct As ADODB.Recordset
Private imChkListBoxIgnore As Integer
Private imChkStnListBoxIgnore As Integer
Private imChkVehListBoxignore As Integer
Private imChkWhoListboxIgnore As Integer
Private imFirstTime As Integer

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkVehListBoxignore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkVehListBoxignore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub
Private Sub chkAllVehicles_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkVehListBoxignore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkVehListBoxignore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub
Private Sub chkAllWho_Click()
Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkWhoListboxIgnore Then
        Exit Sub
    End If
    If chkAllWho.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    
    If rbcWho(1).Value = True Then       'department
        If lbcDept.ListCount > 0 Then
            imChkWhoListboxIgnore = True
            lRg = CLng(lbcDept.ListCount - 1) * &H10000 Or 0
            lRet = SendMessageByNum(lbcDept.hwnd, LB_SELITEMRANGE, iValue, lRg)
            imChkWhoListboxIgnore = False
        End If
    Else                            'persons
        If lbcPersons.ListCount > 0 Then
            imChkWhoListboxIgnore = True
            lRg = CLng(lbcPersons.ListCount - 1) * &H10000 Or 0
            lRet = SendMessageByNum(lbcPersons.hwnd, LB_SELITEMRANGE, iValue, lRg)
            imChkWhoListboxIgnore = False
        End If
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStnListBoxIgnore Then
        Exit Sub
    End If
    If ckcStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStnListBoxIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStnListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmCommentRpt
End Sub
'
'       Contact Comments-
'
Private Sub cmdReport_Click()
    Dim ilTemp As Integer
    Dim sCode As String
    Dim sVehicles, sStations, sWho, sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sFollowUpStart As String
    Dim sFollowUpEnd As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'Dim NewForm As New frmViewReport
    Dim sStartTime As String
    Dim sEndTime As String
    Dim slNow As String
    Dim sGenDate As String
    Dim sGenTime As String
    Dim ilFound As Integer
    Dim slTempDate As String
    Dim llTempDate1 As Long
    Dim llTempDate2 As Long
    Dim slMatchingComment As String
    Dim ilPos As Integer
    Dim slVerify As String
    
    On Error GoTo ErrHand
      
    sGenDate = Format$(gNow(), "m/d/yyyy")
    sGenTime = Format$(gNow(), sgShowTimeWSecForm)

    sStartDate = Trim$(CalOnAirDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If
    sEndDate = Trim$(CalOffAirDate.Text)
    If sEndDate = "" Then
        sEndDate = "12/31/2069"
    End If
    If gIsDate(sStartDate) = False Or (Len(Trim$(sStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Or (Len(Trim$(sEndDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sEndDate = Format(sEndDate, "m/d/yyyy")

    sFollowUpStart = Trim$(CalFollowStart.Text)
    If sFollowUpStart = "" Then
        sFollowUpStart = "1/1/1970"
    End If
    sFollowUpEnd = Trim$(CalFollowEnd.Text)
    If sFollowUpEnd = "" Then
        sFollowUpEnd = "12/31/2069"
    End If
    If gIsDate(sFollowUpStart) = False Or (Len(Trim$(sFollowUpStart)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFollowStart.SetFocus
        Exit Sub
    End If
    If gIsDate(sFollowUpEnd) = False Or (Len(Trim$(sFollowUpEnd)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFollowEnd.SetFocus
        Exit Sub
    End If
    sFollowUpStart = Format(sFollowUpStart, "m/d/yyyy")
    sFollowUpEnd = Format(sFollowUpEnd, "m/d/yyyy")

    slVerify = ""
    If lbcStations.SelCount <= 0 Then
        slVerify = "Station"
    End If
    If lbcVehAff.SelCount <= 0 Then
        If Trim$(slVerify) = "" Then
            slVerify = "Vehicle"
        Else
            slVerify = slVerify & ",Vehicle"
        End If
        
    End If
    If rbcWho(0).Value = True Then
        If lbcPersons.SelCount <= 0 Then
            If Trim$(slVerify) = "" Then
                slVerify = "Person"
            Else
                slVerify = slVerify & ",Person"
            End If
        End If
    ElseIf rbcWho(1).Value = True Then
        If lbcDept.SelCount <= 0 Then
        If Trim$(slVerify) = "" Then
                slVerify = "Dept"
            Else
                slVerify = slVerify & ",Dept"
            End If
        End If
    End If
        
    If Trim$(slVerify) <> "" Then
        slVerify = "At least one " & slVerify & " must be selected"
        MsgBox slVerify, vbOK
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"

    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    slMatchingComment = LCase(Trim$(txtMatchingComment.Text))

    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    sVehicles = ""
    sStations = ""
    sWho = ""
     
    If ckcStations.Value = 0 Then    '= 0 Then                        'User did NOT select all stations
        For ilTemp = 0 To lbcStations.ListCount - 1 Step 1
            If lbcStations.Selected(ilTemp) Then
                If Len(sStations) = 0 Then
                    sStations = " and ((shttCode = " & lbcStations.ItemData(ilTemp) & ")"
                Else
                    sStations = sStations & " OR (shttCode = " & lbcStations.ItemData(ilTemp) & ")"
                End If
            End If
        Next ilTemp
        sStations = sStations & ")"
    End If
    
    If chkAllVehicles.Value = 0 Then    '= 0 Then                        'User did NOT select all stations
        For ilTemp = 0 To lbcVehAff.ListCount - 1 Step 1
            If lbcVehAff.Selected(ilTemp) Then
                If Len(sVehicles) = 0 Then
                    sVehicles = " and ((vefcode = " & lbcVehAff.ItemData(ilTemp) & ")"
                Else
                    sVehicles = sVehicles & " OR (vefCode = " & lbcVehAff.ItemData(ilTemp) & ")"
                End If
            End If
        Next ilTemp
        sVehicles = sVehicles & " OR (cctvefcode = 0))"
        'sVehicles = sVehicles & ")"
    End If

    sWho = ""
    If rbcWho(0).Value = True Then
        If chkAllWho.Value = vbUnchecked Then         'User did NOT select all stations
            For ilTemp = 0 To lbcPersons.ListCount - 1 Step 1
                If lbcPersons.Selected(ilTemp) Then
                    If Len(sWho) = 0 Then
                        sWho = " and ((ustcode = " & lbcPersons.ItemData(ilTemp) & ")"
                    Else
                        sWho = sWho & " OR (ustCode = " & lbcPersons.ItemData(ilTemp) & ")"
                    End If
                End If
            Next ilTemp
            sWho = sWho & ")"
        End If
    ElseIf rbcWho(1).Value = True Then      'dept
        If chkAllWho.Value = vbUnchecked Then         'User did NOT select all stations
            For ilTemp = 0 To lbcDept.ListCount - 1 Step 1
                If lbcDept.Selected(ilTemp) Then
                    If Len(sWho) = 0 Then
                        sWho = " and ((dntcode = " & lbcDept.ItemData(ilTemp) & ")"
                    Else
                        sWho = sWho & " OR (dntCode = " & lbcDept.ItemData(ilTemp) & ")"
                    End If
                End If
            Next ilTemp
            sWho = sWho & ")"
        End If
    End If

  
    'Entered start and end dates to Crystal
'    dFWeek = CDate(sStartDate)
'    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
'    dFWeek = CDate(sEndDAte)
'    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula1 = mDatesToCrystal(sStartDate, sEndDate)
    'Follow-up start and end dates to Crystal
'    dFWeek = CDate(sFollowUpStart)
'    sgCrystlFormula10 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
'    dFWeek = CDate(sFollowUpEnd)
'    sgCrystlFormula11 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula2 = mDatesToCrystal(sFollowUpStart, sFollowUpEnd)
    
    sgCrystlFormula3 = Str$(cbcSort1.ListIndex + 1)         'sort fields, major to minor
    sgCrystlFormula4 = Str$(cbcSort2.ListIndex)
    sgCrystlFormula5 = Str$(cbcSort3.ListIndex)
    
    'determine page skips for each grouping
    If ckcSkip1.Value = vbChecked Then
        sgCrystlFormula7 = "'Y'"
    Else
        sgCrystlFormula7 = "'N'"
    End If
    If ckcSkip2.Value = vbChecked Then
        sgCrystlFormula8 = "'Y'"
    Else
        sgCrystlFormula8 = "'N'"
    End If
    If ckcSkip3.Value = vbChecked Then
        sgCrystlFormula9 = "'Y'"
    Else
        sgCrystlFormula9 = "'N'"
    End If
    
    If rbctotalsBy(0).Value = True Then     'detail
        sgCrystlFormula6 = "'D'"
    Else
        sgCrystlFormula6 = "'S'"        'summary
    End If
    
    
    If ckcFollowUp(0).Value = vbChecked And ckcFollowUp(1).Value = vbChecked Then
        sgCrystlFormula10 = "'" & "Including Follow-up Done and Undone" & "'"
    ElseIf ckcFollowUp(0).Value = vbChecked Then
        sgCrystlFormula10 = "'" & "Including Follow-up Done" & "'"
    ElseIf ckcFollowUp(1).Value = vbChecked Then
        sgCrystlFormula10 = "'" & "Including Follow-up Undone" & "'"
    Else
        sgCrystlFormula10 = "'" & "Done or UnDone not selected" & "'"
    End If
    
    If ckcInclID.Value = vbChecked Then     'include Comment Internal ID
        sgCrystlFormula11 = "'Y'"
    Else
        sgCrystlFormula11 = "'N'"
    End If
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    
    'get the selected comments by station, vehicle, and person (or dept or all comments)
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    'Dan M 1-11-12 all changes to this sql call that I previously made are now removed.  see note below
    'Dan M 9-19-11 added afr for crystal 11....also took pieces from below where afr superseded this code
    SQLQuery = "SELECT * FROM cct "
  '  SQLQuery = "SELECT * FROM afr Left outer join cct on afrastcode = cctcode "
    
    SQLQuery = SQLQuery & "LEFT OUTER JOIN VEF_Vehicles on cctvefcode = vefCode "
    SQLQuery = SQLQuery & "INNER JOIN shtt on cctshfcode = shttCode INNER JOIN cst on cctcstCode = cstCode  "
    SQLQuery = SQLQuery & "INNER JOIN ust on cctustCode = ustCode "
'    'Dan M 9-19-11 added
'    SQLQuery = SQLQuery & " LEFT OUTER JOIN DNT on ustDntCode = dntCode "
    If rbcWho(1).Value = True Then               'by dept
        SQLQuery = SQLQuery & "LEFT OUTER JOIN dnt on ustcode = dntcode "
    End If
    'Entered & followup dates have been set to 1/1/70 and/or 12/31/2069 if not entered, select all dates
    'test the Entered date selectivty
    SQLQuery = SQLQuery & "where (cctEnteredDate >= " & "'" & Format$(sStartDate, sgSQLDateForm) & "' AND cctEnteredDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    'test followup date selectivity
    SQLQuery = SQLQuery & " and (cctActionDate >= " & "'" & Format$(sFollowUpStart, sgSQLDateForm) & "' AND cctActionDate <= '" & Format$(sFollowUpEnd, sgSQLDateForm) & "')"
    If Trim$(sStations) <> "" Then
        SQLQuery = SQLQuery & sStations
    End If
    If Trim$(sVehicles) <> "" Then
        SQLQuery = SQLQuery & sVehicles
    End If
        
    If (rbcWho(0).Value = True) Then    'persons, user codes selected
        SQLQuery = SQLQuery & sWho
    End If
    Set rst_cct = gSQLSelectCall(SQLQuery)
    Do While Not rst_cct.EOF
        If (rst_cct!cctDone = "Y" And ckcFollowUp(0).Value = vbChecked) Or (rst_cct!cctDone <> "Y" And ckcFollowUp(1).Value = vbChecked) Then
            ilFound = False
            'If (Trim$(sgCrystlFormula3) = "1" Or Trim$(sgCrystlFormula4) = "1" Or Trim$(sgCrystlFormula5) = "1") And (DateValue(gAdjYear(rst_cct!cctActionDate)) = DateValue("12/31/2069")) Then
            If (Trim$(sgCrystlFormula3) = "1" Or Trim$(sgCrystlFormula4) = "1" Or Trim$(sgCrystlFormula5) = "1") And (rst_cct!cctActionDate = DateValue("12/31/2069")) Then
                'ignore any comments without an action date if the report is using any sorts with the Follow-up (Action) date
                ilFound = ilFound
            Else
                'see if theres a matching comment that needs to be met
                sStr = LCase(Trim$(rst_cct!cctComment))
                'test all lower case
                ilPos = InStr(sStr, slMatchingComment)
                If (Len(slMatchingComment) = 0) Or (ilPos > 0 And Len(Trim$(slMatchingComment)) > 0) Then     'either no comment entered to match or one was entered and it matches the string
                    ilFound = True
                End If
            End If
            If ilFound Then
                slTempDate = Format(rst_cct!cctActionDate, "m/d/yyyy")
                llTempDate1 = DateValue(slTempDate)
                slTempDate = Format(rst_cct!cctEnteredDate, "m/d/yyyy")
                llTempDate2 = DateValue(slTempDate)
                SQLQuery = "INSERT INTO afr "
                SQLQuery = SQLQuery & " (afrAstCode, afrcrfcsfcode,afrAttCode, afrGenDate, afrGenTime) "
                SQLQuery = SQLQuery & " VALUES (" & rst_cct!cctCode & ", " & llTempDate1 & ", " & llTempDate2 & ", '" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
                 
                cnn.BeginTrans
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "CommentRpt-cmdReport_Click"
                    cnn.RollbackTrans
                    Exit Sub
                End If
                cnn.CommitTrans
            End If
        End If
        rst_cct.MoveNext
    Loop
    'dan m 1-11-12
    ' cr2008 ignored the change of sql, so even though this was not enough, it never showed an error
    ' SQLQuery = "Select * from afr where   afrgenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))
    'when we moved to cr11, this caused an error.  on 9-19, I tried to fix the sql by changing it above and then using it here. Whoops!
    'I changed above back to original, then copied the sql call from the report to here.
    
    SQLQuery = " SELECT cctActionDate, cctComment, cctEnteredDate, shttCallLetters, ustName, vefName, dntName, cctVefCode, afrAstCode, cctDone, afrAttCode, afrCrfCsfcode, cstName, cctCode"
    SQLQuery = SQLQuery & " FROM   {oj (((((afr LEFT OUTER JOIN cct ON afrAstCode =cctCode) INNER JOIN ust ON cctUstCode =ustCode) INNER JOIN shtt ON cctShfCode=shttCode)"
    SQLQuery = SQLQuery & " INNER JOIN cst ON cctCstCode=cstCode) LEFT OUTER JOIN VEF_Vehicles ON cctVefCode= vefCode) LEFT OUTER JOIN dnt ON ustDntCode=dntCode} "
    SQLQuery = SQLQuery & " where afrgenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))

    
    'dan m 9-19-11 for rollback, add afr
   ' SQLQuery = "Select * from afr where   afrgenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))
   ' SQLQuery = SQLQuery & " and afrgenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "afComments.rpt", "AfComments"
     
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime

    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "CommentRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
   
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Comments-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmCommentRpt
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
    gSelectiveStationsFromImport lbcStations, ckcStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
    If imFirstTime = True Then
        mPopSorts cbcSort1, False           'Major group total #1, dont allow NONE for a choice
        mPopSorts cbcSort2, True           ' group total #2, allow  NONE for a choice
        mPopSorts cbcSort3, True           ' group total #3, allow  NONE for a choice
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
Dim ilHalf As Integer
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
'    ckcPersons.Top = 240
'
'    ilHalf = (Frame2.Height - ckcPersons.Height - chkAllVehicles.Height - 120) / 2
'    lbcPerson.Move ckcPersons.Left, ckcPersons.Top + ckcPersons.Height
'    lbcPerson.Height = ilHalf
'    lbcVehAff.Height = ilHalf
'    lbcStations.Height = ilHalf
'    chkAllVehicles.Top = lbcPerson.Top + lbcPerson.Height
'    ckcStations.Top = chkAllVehicles.Top
'    lbcVehAff.Top = chkAllVehicles.Top + chkAllVehicles.Height
'    lbcStations.Top = lbcVehAff.Top
'    lbcVehAff.Height = ilHalf
'    lbcStations.Height = ilHalf
'
    
    gSetFonts frmCommentRpt
    gCenterForm frmCommentRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim ilRet As Integer
    Dim lRg As Long
    Dim lRet As Long
    
    imFirstTime = True
    imChkListBoxIgnore = False
    imChkVehListBoxignore = False
    imChkStnListBoxIgnore = False
    frmCommentRpt.Caption = "Contact Comment Report - " & sgClientName
    
    'ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    gPopDept
    gPopUst
    
    'populate the Stations, Vehicles & Advertisers (currently only advertisers are selectable)
    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
    
    lbcVehAff.Clear

    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    
    lbcPersons.Clear
    For iLoop = 0 To UBound(tgUstInfo) - 1 Step 1
        lbcPersons.AddItem Trim$(tgUstInfo(iLoop).sName)
        lbcPersons.ItemData(lbcPersons.NewIndex) = tgUstInfo(iLoop).iCode
    Next iLoop
    
    lbcDept.Visible = False
    lbcDept.Clear
    For iLoop = 0 To UBound(tgDeptInfo) - 1 Step 1
        lbcDept.AddItem Trim$(tgDeptInfo(iLoop).sName)
        lbcDept.ItemData(lbcDept.NewIndex) = tgDeptInfo(iLoop).iCode
    Next iLoop

    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    rst_cct.Close
    'ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmCommentRpt = Nothing
End Sub

Private Sub grdVehAff_Click()
    If chkAllVehicles.Value = 1 Then
        imChkVehListBoxignore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = 0    'chged from false to 0 10-22-99
        imChkVehListBoxignore = False
    End If
End Sub

Private Sub lbcPerson_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkAllWho.Value = vbChecked Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkAllWho.Value = vbUnchecked
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub lbcDept_Click()
If imChkWhoListboxIgnore Then
        Exit Sub
    End If
    If chkAllWho.Value = vbChecked Then
        imChkWhoListboxIgnore = True
        chkAllWho.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkWhoListboxIgnore = False
    End If
End Sub

Private Sub lbcPersons_Click()
  If imChkWhoListboxIgnore Then
        Exit Sub
    End If
    If chkAllWho.Value = vbChecked Then
        imChkWhoListboxIgnore = True
        chkAllWho.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkWhoListboxIgnore = False
    End If
End Sub

Private Sub lbcStations_Click()
    If imChkStnListBoxIgnore Then
        Exit Sub
    End If
    If ckcStations.Value = 1 Then
        imChkStnListBoxIgnore = True
        ckcStations.Value = 0    'chged from false to 0 10-22-99
        imChkStnListBoxIgnore = False
    End If
End Sub
Private Sub lbcVehAff_Click()
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = 1 Then
        imChkVehListBoxignore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = 0    'chged from false to 0 10-22-99
        imChkVehListBoxignore = False
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

'
'           Populate the drop down with the valid fields to sort:  Vehicle, Market, station and Advertiser (advt/contr# implied)
'           <input> cbcControl as dropdown control
'                   ilShowNone : true - show None as a choice, else false to default to 1st element
'           DH 7-12-04
Public Sub mPopSorts(cbcControl As control, ilShowNone As Integer)
    If ilShowNone Then
        cbcControl.AddItem "None"
    End If
    cbcControl.AddItem "Follow-Up Date"
    cbcControl.AddItem "Person"
    cbcControl.AddItem "Posted Date"
    cbcControl.AddItem "Station"
    cbcControl.AddItem "Vehicle"
    cbcControl.ListIndex = 0
    
   
End Sub

Private Sub rbcWho_Click(Index As Integer)
    If Index = 0 Then       'select by people
        lbcPersons.Visible = True
        lbcDept.Visible = False
        chkAllWho.Caption = "All Persons"
        chkAllWho.Visible = True
    ElseIf Index = 1 Then      'dept
        lbcDept.Visible = True
        lbcPersons.Visible = False
        chkAllWho.Caption = "All Departments"
        chkAllWho.Visible = True
    Else
        lbcPersons.Visible = False
        lbcDept.Visible = False
        chkAllWho.Visible = False
    End If
    
    
End Sub
'       mDatesToCrystal
'       determine dates entered to prepare to send to crystal reports
'       <input> sSTart date - string start date
'               sEndDate - string end date
'       return - string to pass to crystal
'
Private Function mDatesToCrystal(sStartDate As String, sEndDate As String) As String
Dim slStr As String
    slStr = ""
    If Trim$(sStartDate) = "1/1/1970" And sEndDate = "12/31/2069" Then
      slStr = "for all Dates"
    ElseIf Trim$(sStartDate) = "1/1/1970" Then
       slStr = "thru " + Trim$(sEndDate)
    ElseIf Trim$(sEndDate) = "12/31/2069" Then
        slStr = "from " + Trim$(sStartDate)
    Else
        slStr = "for " + Trim$(sStartDate) + "-" + Trim$(sEndDate)
    End If
    mDatesToCrystal = slStr
End Function
