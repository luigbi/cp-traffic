VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSpotUtil 
   Caption         =   "Spot Utility"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AffSpotUtil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin V81Affiliate.CSI_Calendar edcStartDate 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Text            =   "05/24/2024"
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
      CSI_DefaultDateType=   1
   End
   Begin V81Affiliate.CSI_Calendar edcEndDate 
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Top             =   1320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Text            =   "05/24/2024"
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
      CSI_DefaultDateType=   1
   End
   Begin VB.ListBox lbcStations 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Index           =   1
      ItemData        =   "AffSpotUtil.frx":08CA
      Left            =   6120
      List            =   "AffSpotUtil.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   4000
   End
   Begin VB.OptionButton rbcSpots 
      Caption         =   "Delete Network Web Spots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7860
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   2385
   End
   Begin VB.TextBox txtActiveDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4545
      TabIndex        =   13
      Top             =   2925
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Active as"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4545
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdMoveLeft 
      Caption         =   "<"
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveRight 
      Caption         =   ">"
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.ListBox lbcStations 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Index           =   0
      ItemData        =   "AffSpotUtil.frx":08CE
      Left            =   360
      List            =   "AffSpotUtil.frx":08D5
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2280
      Width           =   4000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboVehicle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtNumberOfDays 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton rbcSpots 
      Caption         =   "Delete All Spots For All Weeks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4860
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   2745
   End
   Begin VB.OptionButton rbcSpots 
      Caption         =   "Delete Posted Spots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2610
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton rbcSpots 
      Caption         =   "Clear Posted Spots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   1965
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1140
      TabIndex        =   16
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5700
      TabIndex        =   19
      Top             =   6000
      Width           =   3615
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9960
      Top             =   5280
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6750
      FormDesignWidth =   10530
   End
   Begin VB.Label lbcWebType 
      Caption         =   "Production Website"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label lblEndDate 
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "***  Please Create a Database Backup Before Continuing.  ***"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   "***  Warning:  Clearing or Deleting Spots is a Permanent Action and is NOT Reversible.  ***"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5040
      Width           =   8055
   End
   Begin VB.Label lblInclude 
      Caption         =   "Stations to Include"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblExclude 
      Caption         =   "Stations to Exclude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblStations 
      Caption         =   "Select Vehicle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   705
      Width           =   1335
   End
   Begin VB.Label lblEnd 
      Caption         =   "Number of Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label lblStart 
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1380
      Width           =   975
   End
End
Attribute VB_Name = "frmSpotUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmSpotUtil -
'*
'*
'*  Created February, 2007 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2006
'******************************************************

Option Explicit
Option Compare Text

'support for the type ahead
Private imVehicleInChg As Integer
Private imCreateVehInChg As Integer
Private imStationMarketInChg As Integer
Private imCreateVehBSMode As Integer
Private imStationMarketBSMode As Integer
Private imVehicleBSMode As Integer
Private imNumberDays As Integer
Private imClear As Integer
Private imDelete As Integer

Private tmOverlapInfo() As AGMNTOVERLAPINFO

'misc. vars
Private smAdminEMail As String
Private imVefCode As Integer
Private lmBaseAttCode As Long
Private imBaseShttCode As Integer
Private imAddShttCode As Integer
Private lmAttCode As Long
Private smCurDate As String
Private smCurTime As String
Private imWebType As Integer
Private imUnivisionType As Integer
Private smStationName As String
Private imStationHasTimeZoneDefined As Integer
Private imOKToConvertTimeZones As Integer
Private imAgreeType As Integer


'ADO vars
Private adrst As ADODB.Recordset

Private Function mValidateUserInput() As Integer

    Dim slNowDate As String
    Dim slDate As String
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    If Not rbcSpots(0).Value And Not rbcSpots(1).Value And Not rbcSpots(2).Value And Not rbcSpots(3).Value Then
        mValidateUserInput = False
        gMsgBox "Please select Clear Posted Spots or Delete Posted Spots or Delete All Spots All Weeks or Delete Network Web Spots", vbCritical
        rbcSpots(0).SetFocus
        rbcSpots(0).Value = False
        Exit Function
    End If

    
    'We don't need to validate the date info if we're going to delete ALL spots - rbcSpots(2).value = True
    If Not rbcSpots(2).Value Then
        If edcStartDate.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            edcStartDate.SetFocus
            mValidateUserInput = False
            Exit Function
        End If
        If gIsDate(edcStartDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            edcStartDate.SetFocus
            mValidateUserInput = False
            Exit Function
        Else
            slDate = Format(edcStartDate.Text, sgShowDateForm)
        End If
        
        If Not IsNumeric(txtNumberOfDays.Text) Then
            gMsgBox "Number of days must have a numeric value.", vbOKOnly
            txtNumberOfDays.Text = ""
            txtNumberOfDays.SetFocus
            Exit Function
        End If
        
        imNumberDays = Val(txtNumberOfDays.Text)
        If imNumberDays = 0 Then
            gMsgBox "Number of days must be specified.", vbOKOnly
            txtNumberOfDays.SetFocus
            mValidateUserInput = False
            Exit Function
        End If
        If imNumberDays < 0 Then
            gMsgBox "Start Date must be prior to End Date.", vbOKOnly
            edcEndDate.SetFocus
            mValidateUserInput = False
            Exit Function
        End If
        
        slNowDate = Format$(gNow(), "ddddd")
        If DateValue(gAdjYear(slDate)) <= DateValue(gAdjYear(slNowDate)) Then
            Beep
            ilRet = gMsgBox("Warning: You about to Clear or Delete Spots in the Past.  Do you want to continue?", vbOKCancel)
            If ilRet = vbCancel Then
                edcStartDate.SetFocus
                mValidateUserInput = False
                Exit Function
            End If
        End If
    End If
    mValidateUserInput = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mValidateUserInput"
    Exit Function
End Function


Private Sub mFillVehicle()

    Dim iLoop As Integer
    
    On Error GoTo ErrHand
    
    cboVehicle.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            cboVehicle.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            cboVehicle.ItemData(cboVehicle.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mFillVehicle"
    Exit Sub
End Sub




Private Sub cboVehicle_Change()
    
    Dim llRow As Long
    Dim slName As String
    Dim ilLen As Integer
    
    On Error GoTo ErrHand
'    If imVehicleInChg Then
'        mGenOK
'        Exit Sub
'    End If
    imVehicleInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cboVehicle.Text)
    ilLen = Len(slName)
    If imVehicleBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imVehicleBSMode = False
    End If
    
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboVehicle.ListIndex = llRow
        cboVehicle.SelStart = ilLen
        cboVehicle.SelLength = Len(cboVehicle.Text)
    End If
    Screen.MousePointer = vbDefault
    imVehicleInChg = False
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cboVehicle_Change: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboVehicle_Click()
    
    On Error GoTo ErrHand
    
    lbcStations(0).Clear
    lbcStations(1).Clear
    If cboVehicle.Text <> "" Then
        mShowSelectiveStations
    Else
        Exit Sub
    End If
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cboVehicle_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cboVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    
    imVehicleBSMode = False
    'mGenOK

End Sub

Private Sub cboVehicle_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then
        If cboVehicle.SelLength <> 0 Then
            imVehicleBSMode = True
        End If
    End If
    'mGenOK

End Sub

Private Sub cmdAll_Click()
    
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim slDates As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llActiveDate As Long
    
    On Error GoTo ErrHand
    
    If (txtActiveDate.Visible) And (Trim$(txtActiveDate.Text) <> "") Then
        If Not gIsDate(Trim$(txtActiveDate.Text)) Then
            txtActiveDate.SetFocus
            gMsgBox "Active date is not a valid date", vbCritical
            Exit Sub
        End If
        llActiveDate = DateValue(gAdjYear(txtActiveDate.Text))
        For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
            slStr = lbcStations(0).List(ilLoop)
            ilPos = InStrRev(slStr, " ", -1, vbTextCompare)
            If ilPos > 0 Then
                slStr = Mid(slStr, ilPos + 1)
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slStartDate = Left(slStr, ilPos - 1)
                    slEndDate = Mid(slStr, ilPos + 1)
                    If (slEndDate = "TFN") Then
                        lbcStations(0).Selected(ilLoop) = True
                    Else
                        If llActiveDate <= DateValue(gAdjYear(slEndDate)) Then
                            lbcStations(0).Selected(ilLoop) = True
                        End If
                    End If
                End If
            End If
        Next ilLoop
        cmdMoveRight_Click
    Else
        For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
            lbcStations(1).AddItem lbcStations(0).List(ilLoop)
            lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
        Next ilLoop
        lbcStations(0).Clear
    End If
    
    lbcStations(0).ListIndex = -1
    cmdAll.Visible = False
    txtActiveDate.Visible = False
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cmdAll_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload frmSpotUtil

End Sub

Private Sub cmdUpdate_Click()
        
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    If Not mValidateUserInput Then
        Exit Sub
    End If
    
    '9/12/18: Verify that multicast are aligned
    mAlighMulticast
    
    'UnPost Spots
    If rbcSpots(0).Value Then
        'gLogMsg "Warning you are about to Clear the Posting flags for the selected agreement(s) and time period.  Do you wish to continue? ", "AffUtilsLog.Txt", False
        ilRet = gMsgBox("Warning: This option will reset spots as if they were sent to Network Web but not Posted (removing all Posted status flags from Affiliate Spots, Network Web Spots and Post CP status).  Do you wish to continue?", vbYesNo)
        gLogMsg "Warning: This option will reset spots as if they were sent to Network Web but not Posted (removing all Posted status flags from Affiliate Spots, Network Web Spots and Post CP status).   Do you wish to continue? ", "AffUtilsLog.Txt", False
        If ilRet = vbNo Then
            gLogMsg "   Replied NO", "AffUtilsLog.Txt", False
            Exit Sub
        Else
            gLogMsg "   Replied YES", "AffUtilsLog.Txt", False
        End If
    

        If lbcStations(1).ListCount > 0 Then
            mSpotsClearPosting
            cmdCancel.Caption = "Done"
        Else
            gLogMsg "Nothing to Process", "AffUtilsLog.Txt", False
            gMsgBox "Nothing to Process"
        End If

    End If
    
    'Delete a range of Spots for an Agreement
    If rbcSpots(1).Value Then
        'gLogMsg "Warning you are about to Delete Spots for the selected agreement(s) and time period.  Do you wish to continue? ", "AffUtilsLog.Txt", False
        ilRet = gMsgBox("Warning: This option will remove spots as if they were never generated on the Affiliate system (removing Affiliate Spots, Network Web Spots, and resetting the affiliate affidavit status).  Do you wish to continue?", vbYesNo)
        gLogMsg "   Warning: This option will remove spots as if they were never generated on the Affiliate system (removing Affiliate Spots, Network Web Spots, and resetting the affiliate affidavit status).   Do you wish to continue? ", "AffUtilsLog.Txt", False
        If ilRet = vbNo Then
            gLogMsg "   Replied NO", "AffUtilsLog.Txt", False
            Exit Sub
        Else
            gLogMsg "   Replied YES", "AffUtilsLog.Txt", False
        End If
    
        If lbcStations(1).ListCount > 0 Then
            mSpotsDelete
            cmdCancel.Caption = "Done"
        Else
            gLogMsg "Nothing to Process", "AffUtilsLog.Txt", False
            gMsgBox "Nothing to Process"
        End If

    End If
    
    'Delete All Spots for an Agreement
    If rbcSpots(2).Value Then
        ilRet = gMsgBox("Warning: This option will Remove all spots as if Traffic Logs were never generated (removing Affiliate Spots, Network Web Spots, and Post CP).  Do you wish to continue?", vbYesNo)
        gLogMsg " Warning: This option will Remove all spots as if Traffic Logs were never generated (removing Affiliate Spots, Network Web Spots, and Post CP).  Do you wish to continue? ", "AffUtilsLog.Txt", False
        If ilRet = vbNo Then
            gLogMsg "Replied NO", "AffUtilsLog.Txt", False
            Exit Sub
        Else
            gLogMsg "Replied YES", "AffUtilsLog.Txt", False
        End If

        If lbcStations(1).ListCount > 0 Then
            mSpotsDeleteAll
            cmdCancel.Caption = "Done"
        Else
            gLogMsg "Nothing to Process", "AffUtilsLog.Txt", False
            gMsgBox "Nothing to Process"
        End If

    End If
    
    'Delete Network Web Spots for an Agreement
    If rbcSpots(3).Value Then
        ilRet = gMsgBox("Warning: This option will remove Network Web spots as if they were never Exported (removing Network Web Spots).  Do you wish to continue?", vbYesNo)
        gLogMsg " Warning: This option will remove Network Web spots as if they were never Exported (removing Web Network Spots).  Do you wish to continue? ", "AffUtilsLog.Txt", False
        If ilRet = vbNo Then
            gLogMsg "Replied NO", "AffUtilsLog.Txt", False
            Exit Sub
        Else
            gLogMsg "Replied YES", "AffUtilsLog.Txt", False
        End If

        If lbcStations(1).ListCount > 0 Then
            mSpotsDeleteWeb
            cmdCancel.Caption = "Done"
        Else
            gLogMsg "Nothing to Process", "AffUtilsLog.Txt", False
            gMsgBox "Nothing to Process"
        End If

    End If
    
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cmdUpdate_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub

Private Sub cmdMoveLeft_Click()
    
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    
    On Error GoTo ErrHand
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        If lbcStations(1).Selected(ilLoop) Then
            lbcStations(0).AddItem lbcStations(1).List(ilLoop)
            lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(ilLoop)
        End If
    Next ilLoop
    
    ilInclCount = lbcStations(0).ListCount
    ilExclCount = lbcStations(1).ListCount
    For ilLoop = 0 To ilInclCount - 1 Step 1
        For ilIdx = 0 To ilExclCount - 1 Step 1
            If lbcStations(0).List(ilLoop) = lbcStations(1).List(ilIdx) Then
                lbcStations(1).RemoveItem (ilIdx)
                ilExclCount = ilExclCount - 1
                'mGenOK
                Exit For
            End If
        Next ilIdx
    Next ilLoop
    
    lbcStations(1).ListIndex = -1
    lbcStations(0).ListIndex = -1
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cmdMoveLeft_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdMoveRight_Click()

    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    
    On Error GoTo ErrHand
    
    If rbcSpots(2).Value = True Then
'        If lbcStations(0).ListIndex = 1 Then
'            gMsgBox "This station cannot be moved until the station information is entered into the station area."
'        Else
'            gMsgBox "These stations cannot be moved until the stations information is entered into the station area."
'        End If
'        mGenOK
'        Exit Sub
    End If
    
    For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
        If lbcStations(0).Selected(ilLoop) Then
            lbcStations(1).AddItem lbcStations(0).List(ilLoop)
            lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
        End If
    Next ilLoop
    
    ilInclCount = lbcStations(1).ListCount
    ilExclCount = lbcStations(0).ListCount
    For ilLoop = 0 To ilInclCount - 1 Step 1
        For ilIdx = 0 To ilExclCount - 1 Step 1
            If lbcStations(1).List(ilLoop) = lbcStations(0).List(ilIdx) Then
                lbcStations(0).RemoveItem (ilIdx)
                ilExclCount = ilExclCount - 1
                'mGenOK
                Exit For
            End If
        Next ilIdx
    Next ilLoop
    
    lbcStations(1).ListIndex = -1
    lbcStations(0).ListIndex = -1
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-cmdMoveRight_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub Form_Initialize()
    
    Me.Width = Screen.Width / 1.05   '1.05  '1.15
    Me.Height = Screen.Height / 1.3 '15    '1.45    '1.25
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmSpotUtil
    gCenterStdAlone Me

End Sub

Private Sub Form_Load()

    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    cmdUpdate.Enabled = False
    imCreateVehBSMode = False
    imCreateVehInChg = False
    imStationMarketBSMode = False
    imStationMarketInChg = False
    imVehicleBSMode = False
    imVehicleInChg = False

    
    igPasswordOk = False
    
    imClear = False
    If StrComp(sgUstClear, "Y", vbTextCompare) = 0 Then
        imClear = True
    End If
    
    imDelete = False
    If StrComp(sgUstDelete, "Y", vbTextCompare) = 0 Then
        imDelete = True
    End If
    
    'If igPasswordOk = False Then
    '    cmdUpdate.Enabled = False
    'End If
        '10000
    lbcWebType.FontSize = 6
    If igDemoMode Then
        lbcWebType.Caption = "Demo Mode"
    ElseIf gIsTestWebServer() Then
        lbcWebType.Caption = "Test Website"
    End If
    frmSpotUtil.Caption = "Affiliate Spot Utility - " & sgClientName
    gLogMsg "", "AffUtilsLog.Txt", False
    gLogMsg "   *** Starting Spot Utility   ***", "AffUtilsLog.Txt", False
    mFillVehicle
    edcEndDate.Text = ""
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-Form_Load: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    gLogMsg "", "AffUtilsLog.Txt", False
    gLogMsg "   *** Ending Spot Utility Program   ***", "AffUtilsLog.Txt", False
    Erase tmOverlapInfo
    adrst.Close
    igPasswordOk = False
    Set frmSpotUtil = Nothing
End Sub



Private Sub lbcStations_DblClick(Index As Integer)

    On Error GoTo ErrHand
    
    If lbcStations(0).ListIndex >= 0 Then
        lbcStations(1).AddItem lbcStations(0).Text
        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(lbcStations(0).ListIndex)
        lbcStations(0).RemoveItem (lbcStations(0).ListIndex)
        lbcStations(1).ListIndex = -1
        lbcStations(0).ListIndex = -1
    End If

    If lbcStations(1).ListIndex >= 0 Then
        lbcStations(0).AddItem lbcStations(1).Text
        lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(lbcStations(1).ListIndex)
        lbcStations(1).RemoveItem (lbcStations(1).ListIndex)
        lbcStations(0).ListIndex = -1
        lbcStations(1).ListIndex = -1
        If rbcSpots(1).Value = True And lbcStations(0).ListCount > 0 Then
            cmdAll.Visible = True
            If Trim$(txtActiveDate.Text) = "" Then
                txtActiveDate.Text = edcStartDate.Text
            End If
            txtActiveDate.Visible = True
        End If
        
    End If
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-lbcStations_DblClick: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub rbcSpots_Click(Index As Integer)

    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    cmdUpdate.Enabled = False
    
    lbcStations(0).Clear
    lbcStations(1).Clear
    edcStartDate.Text = ""
    txtActiveDate.Text = ""
    txtNumberOfDays.Text = ""
    cboVehicle.Clear
    Call mFillVehicle
    lblInclude.Caption = "Stations to Include"
    lblInclude.ForeColor = vbBlack
    lblExclude.ForeColor = vbBlack
    lblExclude.Caption = "Stations to Exclude"
    
    
    If rbcSpots(0).Value Then
        If StrComp(sgUstClear, "Y", vbTextCompare) <> 0 Then
            CSPWord.Show vbModal
            If igPasswordOk Then
                cmdUpdate.Enabled = True
            End If
        Else
            cmdUpdate.Enabled = True
        End If
    End If
    
    If rbcSpots(1).Value Then
        If StrComp(sgUstDelete, "Y", vbTextCompare) <> 0 Then
            CSPWord.Show vbModal
            If igPasswordOk Then
                cmdUpdate.Enabled = True
            End If
        Else
            cmdUpdate.Enabled = True
        End If
    End If
    
    If rbcSpots(2).Value Then
        CSPWord.Show vbModal
        If igPasswordOk Then
            cmdUpdate.Enabled = True
        End If
    End If
    
    If rbcSpots(3).Value Then
        CSPWord.Show vbModal
        If igPasswordOk Then
            cmdUpdate.Enabled = True
        End If
    End If
    
    If rbcSpots(2).Value Then
        txtActiveDate.Visible = False
        cmdAll.Visible = False
        edcStartDate.Visible = False
        lblStart.Visible = False
        lblEnd.Visible = False
        txtNumberOfDays.Visible = False
        edcEndDate.Visible = False
        lblEndDate.Visible = False
    Else
        txtActiveDate.Visible = True
        cmdAll.Visible = True
        edcStartDate.Visible = True
        lblStart.Visible = True
        lblEnd.Visible = True
        txtNumberOfDays.Visible = True
        edcEndDate.Visible = True
        lblEndDate.Visible = True
    End If
    
    Exit Sub
    
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-rbcSpots_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub

Private Sub mShowSelectiveStations()
    
    Dim shtt_rst As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim ilFromVefCode As Integer
    Dim ilCreateVefCode As Integer
    Dim slDate As String
    Dim llStartDate As Long
    Dim llTodayDate As Long
    Dim slName As String
    Dim llRow As Long
    Dim slEndDate As String
    Dim slRange As String

    
    On Error GoTo ErrHand
   
    lbcStations(0).Clear
    
    slName = cboVehicle.Text
    
    If slName = "" Then
        Exit Sub
    End If
    
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    ilFromVefCode = CInt(cboVehicle.ItemData(llRow))
    imVefCode = ilFromVefCode

    SQLQuery = "SELECT DISTINCT attCode, attShfCode, attVefCode, attDropDate, attOffAir, AttOnAir"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = '" & ilFromVefCode & "')"
    Set att_rst = gSQLSelectCall(SQLQuery)

    ReDim lmAttCodes(0 To 1000)
    While Not att_rst.EOF
        If (lmBaseAttCode = att_rst!attCode) Or ((imBaseShttCode = att_rst!attshfcode) And (imVefCode = att_rst!attvefCode)) Then
        Else
            If DateValue(gAdjYear(att_rst!attDropDate)) < DateValue(gAdjYear(att_rst!attOffAir)) Then
                slEndDate = Format$(att_rst!attDropDate, sgShowDateForm)
            Else
                slEndDate = Format$(att_rst!attOffAir, sgShowDateForm)
            End If
            If (DateValue(gAdjYear(att_rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (att_rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                slRange = ""
            Else
                slRange = Format$(Trim$(att_rst!attOnAir), sgShowDateForm)
            End If
            If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (att_rst!attOffAir = "12/31/69") Then
                If slRange <> "" Then
                    slRange = slRange & "-TFN"
                End If
            Else
                If slRange <> "" Then
                    slRange = slRange & "-" & slEndDate    'att_rst!attOffAir
                Else
                    slRange = "Thru " & slEndDate 'att_rst!attOffAir
                End If
            End If
            SQLQuery = "SELECT shttCallLetters, mktName"
            SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode"
            SQLQuery = SQLQuery + " WHERE (shttCode = " & att_rst!attshfcode & ")"
            Set shtt_rst = gSQLSelectCall(SQLQuery)
            If Not shtt_rst.EOF Then
        
                If IsNull(shtt_rst!mktName) = True Then
                    lbcStations(0).AddItem Trim$(shtt_rst!shttCallLetters) & " " & slRange
                Else
                    lbcStations(0).AddItem Trim$(shtt_rst!shttCallLetters) & " , " & Trim$(shtt_rst!mktName) & " " & slRange
                End If
                'lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attshfCode
                lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attCode
            End If
        End If
        att_rst.MoveNext
    Wend
    
    If lbcStations(0).ListCount > 0 Then
        cmdAll.Visible = True
        If Trim$(txtActiveDate.Text) = "" Then
            txtActiveDate.Text = edcStartDate.Text
        End If
        txtActiveDate.Visible = True
    End If
    'mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mShowSelectiveStations"
End Sub

Private Sub mGenOK()

    Dim ilRet As Integer
    
    On Error GoTo ErrHand

    If lbcStations(1).ListCount > 0 Then
        If igPasswordOk Then
            If (rbcSpots(0).Value = True) Or (rbcSpots(1).Value = True And cboVehicle.Text <> "") Or (rbcSpots(2).Value = True) Or (rbcSpots(3).Value = True And cboVehicle.Text <> "") Then
                cmdUpdate.Enabled = True
            End If
        End If
    Else
        'cmdUpdate.Enabled = False
    End If

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSpotUtil-mGenOK: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffUtilsLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub
    
Private Sub rbcSpots_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'rbcSpots(Index).ToolTipText = ""
    If Index = 0 Then
        rbcSpots(Index).ToolTipText = "This option will reset spots as if they were sent to Network Web but not Posted"
    ElseIf Index = 1 Then
        rbcSpots(Index).ToolTipText = "This option will remove spots as if they were never generated on the Affiliate system"
    ElseIf Index = 2 Then
        rbcSpots(Index).ToolTipText = "This option will Remove all spots as if Traffic Logs were never generated"
    ElseIf Index = 3 Then
        rbcSpots(Index).ToolTipText = "This option will remove Network Web spots as if they were never Exported"
    Else
        rbcSpots(Index).ToolTipText = ""
    End If
End Sub

Private Sub txtActiveDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mSpotsClearPosting() As Integer

    Dim ilRet As Integer
    Dim slVef As String
    Dim slSta As String
    Dim slTempSta As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llRet As Long
    ReDim llAtfArray(0 To 0) As Long
    Dim rst As ADODB.Recordset
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slEndDate As String
    Dim slStartDate As String
    Dim llAttCode As Long
    Dim llAstPostedCount As Long
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim blEOF As Boolean
    
    On Error GoTo ErrHand
    mSpotsClearPosting = False
    gLogMsg "", "AffUtilsLog.Txt", False
    If igPasswordOk Or imClear Then
        cmdUpdate.Enabled = True
    Else
        gMsgBox "Sorry, you do not have the necessary rights to Clear Posted Spots", vbOK
        Exit Function
    End If
    
    slStartDate = Trim$(edcStartDate.Text)
    slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
    slVef = gGetVehNameByVefCode(imVefCode)
    gLogMsg "Clearing Posted Spots Running On: " & slVef & "  Beginning: " & slStartDate & " For: " & imNumberDays & " Days.", "AffUtilsLog.Txt", False
    slSta = ""
    
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        gLogMsg "   ", "AffUtilsLog.Txt", False
        slStartDate = Trim$(edcStartDate.Text)
        slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
        slVef = gGetVehNameByVefCode(imVefCode)
        slMoDate = gObtainPrevMonday(slStartDate)
        slSuDate = gObtainNextSunday(slMoDate)
        imNumberDays = txtNumberOfDays.Text

        llAttCode = lbcStations(1).ItemData(ilLoop)
        slTempSta = gGetCallLettersByAttCode(llAttCode)
        slSta = slSta & slTempSta & ", "
        
        Do
            If DateValue(slEndDate) < DateValue(slSuDate) Then
                slSuDate = slEndDate
            End If
            'SQLQuery = "Select COUNT(astCode) from AST"
            'SQLQuery = SQLQuery + " WHERE"
            'SQLQuery = SQLQuery + " astCPStatus = 1 AND astPledgeStatus <> 8 AND astPledgeStatus <> 4"
            'SQLQuery = SQLQuery + " AND astAtfCode = " & llAttCode
            'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            'Set rst = gSQLSelectCall(SQLQuery)
            llAstPostedCount = 0
            SQLQuery = "Select * from AST"
            SQLQuery = SQLQuery + " WHERE"
            SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                blEOF = False
            Else
                blEOF = True
            End If
            Do While Not rst.EOF
                tlDatPledgeInfo.lAttCode = rst!astAtfCode
                tlDatPledgeInfo.lDatCode = rst!astDatCode
                tlDatPledgeInfo.iVefCode = rst!astVefCode
                tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
                tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
                ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                If rst!astCPStatus = 1 Then
                    If tlDatPledgeInfo.iPledgeStatus <> 8 And tlDatPledgeInfo.iPledgeStatus <> 4 Then
                        llAstPostedCount = llAstPostedCount + 1
                    End If
                End If
                rst.MoveNext
            Loop
            
            'If Not rst.EOF Then
            If Not blEOF Then
                'gLogMsg "   " & CStr(rst(0).Value) & " ** Local ** Posted Spots were Cleared from the Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
                gLogMsg "   " & llAstPostedCount & " ** Local ** Posted Spots were Cleared from the Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            End If
           
            'If rst(0).Value <> 0 Then
            If llAstPostedCount <> 0 Then
                SQLQuery = "UPDATE ast SET "
                SQLQuery = SQLQuery + "astCPStatus = 0"
                SQLQuery = SQLQuery + " WHERE"
                SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "SpotUtil-mSpotsClearPosting"
                    mSpotsClearPosting = False
                    Exit Function
                End If
                
                ilRet = mSetCpttStatus(slMoDate, slSuDate, slStartDate, llAttCode)
            End If
            
            'If rst(0).Value <> 0 Then
            'If llAstPostedCount <> 0 Then
            'D.S. 08-15-17 changed below to "If Not blEOF Then". you can use "llAstPostedCount <> 0"
            'because there could be spots posted on the web and not in pervasive.
            If Not blEOF Then
                slStr = "Update Spots Set RecType = 0, PostedFlag = 0, ExportedFlag = 0, statusCode = NULL, exportdate = NULL, MRReason = NULL Where PostedFlag = 1 And attCode = " & llAttCode
                slStr = slStr & " AND (PledgeStartDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND PledgeEndDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                llRet = gExecWebSQLWithRowsEffected(slStr)
            Else
                llRet = 0
            End If
            
            gLogMsg "   " & CStr(llRet) & " ** Web ** Posted Spots were Cleared from the Agreement " & CStr(llAttCode) & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            ilRet = mDeleteMGandMissed(llAttCode, slStartDate, slSuDate, False)
            slMoDate = DateAdd("d", 7, slMoDate)
            slSuDate = gObtainNextSunday(slMoDate)
            slStartDate = slMoDate
            
        Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(slEndDate))
        
    Next ilLoop
    slSta = slSta & "END List"
    gLogMsg "    Stations: " & slSta, "AffUtilsLog.Txt", False
    mSpotsClearPosting = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSpotsClearPosting"
    Exit Function
End Function

Private Function mSpotsDelete() As Integer

    Dim ilRet As Integer
    Dim slVef As String
    Dim slSta As String
    Dim slTempSta As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llRet As Long
    Dim rst As ADODB.Recordset
    Dim rstDel As ADODB.Recordset
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slEndDate As String
    Dim slStartDate As String
    Dim llAttCode As Long
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    mSpotsDelete = False
    gLogMsg "", "AffUtilsLog.Txt", False
    If igPasswordOk Or imClear Then
        cmdUpdate.Enabled = True
    Else
        gMsgBox "Sorry, you do not have the necessary rights to Clear Posted Spots", vbOK
        Exit Function
    End If
    
    slStartDate = Trim$(edcStartDate.Text)
    slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
    slVef = gGetVehNameByVefCode(imVefCode)
    gLogMsg "Deleting Spots Running On: " & slVef & "  Beginning: " & slStartDate & " For: " & imNumberDays & " Days.", "AffUtilsLog.Txt", False
    slSta = ""
    
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        gLogMsg "   ", "AffUtilsLog.Txt", False
        slStartDate = Trim$(edcStartDate.Text)
        slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
        slVef = gGetVehNameByVefCode(imVefCode)
        slMoDate = gObtainPrevMonday(slStartDate)
        slSuDate = gObtainNextSunday(slMoDate)
        imNumberDays = txtNumberOfDays.Text
        llAttCode = lbcStations(1).ItemData(ilLoop)
        slTempSta = gGetCallLettersByAttCode(llAttCode)
        slSta = slSta & slTempSta & ", "
        Do
            If DateValue(slEndDate) < DateValue(slSuDate) Then
                slSuDate = slEndDate
            End If
            
            'D.S. 6/5/19 TTP 9215 GEt rid of any orphaned Missed or MG spots
            SQLQuery = "Select astCode, astStatus, astLkAstCode from AST"
            SQLQuery = SQLQuery + " WHERE"
            SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
            SQLQuery = SQLQuery + " AND astLkAstCode <> 0 "
            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            Set rst = gSQLSelectCall(SQLQuery)
            Do While Not rst.EOF
                SQLQuery = "Delete from ast"
                SQLQuery = SQLQuery + " WHERE"
                SQLQuery = SQLQuery + " astCode = " & rst!astLkAstCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "SpotUtil-mSpotsDelete"
                    mSpotsDelete = False
                    Exit Function
                End If
                rst.MoveNext
            Loop
            'End TTP 9215
            
            SQLQuery = "Select COUNT(astCode) from AST"
            SQLQuery = SQLQuery + " WHERE"
            SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            Set rst = gSQLSelectCall(SQLQuery)
            
            If Not rst.EOF Then
                gLogMsg "   " & CStr(rst(0).Value) & " Spots were Deleted from the Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            End If
           
            If rst(0).Value <> 0 Then
                SQLQuery = "Delete from ast"
                SQLQuery = SQLQuery + " WHERE"
                SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "SpotUtil-mSpotsDelete"
                    mSpotsDelete = False
                    Exit Function
                End If
            End If
            
            ilRet = mSetCpttStatus(slMoDate, slSuDate, slStartDate, llAttCode)
        
            blRet = mDeleteMGandMissed(llAttCode, slStartDate, slSuDate, False)
            
            slStr = "Delete From Spots Where attCode = " & llAttCode
            slStr = slStr & " AND (PledgeStartDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND PledgeEndDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            llRet = gExecWebSQLWithRowsEffected(slStr)
            gLogMsg "   " & CStr(llRet) & " Spot(s) were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            
            slStr = "Delete From SpotRevisions Where attCode = " & llAttCode
            slStr = slStr & " AND (PledgeStartDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND PledgeEndDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            llRet = gExecWebSQLWithRowsEffected(slStr)
            gLogMsg "   " & CStr(llRet) & " SpotRevision(s) were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            
            slMoDate = DateAdd("d", 7, slMoDate)
            slSuDate = gObtainNextSunday(slMoDate)
            slStartDate = slMoDate
            
        Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(slEndDate))
    Next ilLoop
        
    slSta = slSta & "END List"
    gLogMsg "    Station(s): " & slSta, "AffUtilsLog.Txt", False
    mSpotsDelete = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSpotsDelete"
    Exit Function
End Function


Private Function mSpotsDeleteAll() As Integer

    Dim ilRet As Integer
    Dim slVef As String
    Dim slSta As String
    Dim slTempSta As String
    Dim slStr As String
    Dim ilStaCode As Integer
    Dim ilLoop As Integer
    Dim llRet As Long
    Dim rst As ADODB.Recordset
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slEndDate As String
    Dim slStartDate As String
    Dim llAttCode As Long
    Dim slTemp As String
    Dim ilPos As Integer
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    mSpotsDeleteAll = False
    gLogMsg "", "AffUtilsLog.Txt", False
    
    If igPasswordOk Or imClear Then
        cmdUpdate.Enabled = True
    Else
        gMsgBox "Sorry, you do not have the necessary rights to Delete All Spots", vbOK
        Exit Function
    End If

    slVef = gGetVehNameByVefCode(imVefCode)
    gLogMsg "Deleting ALL Spots Running On: " & slVef & "  for the selected stations:", "AffUtilsLog.Txt", False
    slSta = ""
    
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        llAttCode = lbcStations(1).ItemData(ilLoop)
        slTempSta = gGetCallLettersByAttCode(llAttCode)
        slSta = slSta & slTempSta & ", "
        slTemp = lbcStations(1).List(ilLoop)
        slStr = lbcStations(1).List(ilLoop)
        ilPos = InStrRev(slStr, " ", -1, vbTextCompare)
        If ilPos > 0 Then
            slStr = Mid(slStr, ilPos + 1)
            ilPos = InStr(1, slStr, "-", vbTextCompare)
            If ilPos > 0 Then
                slStartDate = Left(slStr, ilPos - 1)
                slEndDate = Mid(slStr, ilPos + 1)
                If (slEndDate = "TFN") Then
                    slEndDate = "2069-12-31"
                End If
            End If
        End If
         
         SQLQuery = "Select COUNT(astCode) from AST"
         SQLQuery = SQLQuery + " WHERE"
         SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
         Set rst = gSQLSelectCall(SQLQuery)
         
         If Not rst.EOF Then
             gLogMsg "   " & CStr(rst(0).Value) & " All Spots were Deleted from the Agreement " & " " & slTemp & " " & llAttCode, "AffUtilsLog.Txt", False
         End If
        
         SQLQuery = "DELETE FROM ast"
         SQLQuery = SQLQuery + " WHERE astAtfCode = " & llAttCode
         'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SpotUtil-mSpotsDeleteAll"
            mSpotsDeleteAll = False
            Exit Function
        End If
         
         SQLQuery = "DELETE FROM cptt"
         SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & llAttCode
         'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SpotUtil-mSpotsDeleteAll"
            mSpotsDeleteAll = False
            Exit Function
        End If
         
         'ilRet = mSetCpttStatus(slMoDate, slSuDate, slStartDate, llAttCode)
         blRet = mDeleteMGandMissed(llAttCode, "12/31/1999", "12/31/2099", False)
         slStr = "Delete FROM Spots WHERE attCode = " & llAttCode
         llRet = gExecWebSQLWithRowsEffected(slStr)
         slStr = "Delete FROM SpotRevisions WHERE attCode = " & llAttCode
         llRet = gExecWebSQLWithRowsEffected(slStr)
         
         gLogMsg "   " & CStr(llRet) & " All Spots were Deleted from the Web for Agreement " & slTemp & " " & llAttCode, "AffUtilsLog.Txt", False
    Next ilLoop
        
    slSta = slSta & "END List"
    gLogMsg "    Stations: " & slSta, "AffUtilsLog.Txt", False
    mSpotsDeleteAll = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSpotsDeleteAll"
    Exit Function
End Function

Private Function mSpotsDeleteWeb() As Integer

    Dim slVef As String
    Dim slSta As String
    Dim slTempSta As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llRet As Long
    Dim rst As ADODB.Recordset
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slEndDate As String
    Dim slStartDate As String
    Dim llAttCode As Long
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    mSpotsDeleteWeb = False
    gLogMsg "", "AffUtilsLog.Txt", False
    If igPasswordOk Or imClear Then
        cmdUpdate.Enabled = True
    Else
        gMsgBox "Sorry, you do not have the necessary rights to Delete Network Web Spots", vbOK
        Exit Function
    End If
    
    slStartDate = Trim$(edcStartDate.Text)
    slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
    slVef = gGetVehNameByVefCode(imVefCode)
    gLogMsg "Deleting Spots Running On: " & slVef & "  Beginning: " & slStartDate & " For: " & imNumberDays & " Days.", "AffUtilsLog.Txt", False
    slSta = ""
    
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        gLogMsg "   ", "AffUtilsLog.Txt", False
        slStartDate = Trim$(edcStartDate.Text)
        slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
        slVef = gGetVehNameByVefCode(imVefCode)
        slMoDate = gObtainPrevMonday(slStartDate)
        slSuDate = gObtainNextSunday(slMoDate)
        imNumberDays = txtNumberOfDays.Text

        llAttCode = lbcStations(1).ItemData(ilLoop)
        slTempSta = gGetCallLettersByAttCode(llAttCode)
        slSta = slSta & slTempSta & ", "
        
        Do
            If DateValue(slEndDate) < DateValue(slSuDate) Then
                slSuDate = slEndDate
            End If
            
            blRet = mDeleteMGandMissed(llAttCode, slStartDate, slSuDate, False)
            
            slStr = "Delete From Spots Where attCode = " & llAttCode
            slStr = slStr & " AND (PledgeStartDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND PledgeEndDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            llRet = gExecWebSQLWithRowsEffected(slStr)
            gLogMsg "   " & CStr(llRet) & " Spots were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
            slStr = "Delete From SpotRevisions Where attCode = " & llAttCode
            slStr = slStr & " AND (PledgeStartDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND PledgeEndDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            llRet = gExecWebSQLWithRowsEffected(slStr)
            gLogMsg "   " & CStr(llRet) & " SpotRevision(s) were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False

            slMoDate = DateAdd("d", 7, slMoDate)
            slSuDate = gObtainNextSunday(slMoDate)
            slStartDate = slMoDate
            
        Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(slEndDate))
    Next ilLoop
        
    slSta = slSta & "END List"
    gLogMsg "    Stations: " & slSta, "AffUtilsLog.Txt", False
    mSpotsDeleteWeb = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSpotsDeleteWeb"
    Exit Function
End Function

Private Function mSetCpttStatus(sMoDate As String, sSuDate As String, slStartDate As String, lAttCode As Long) As Integer

    Dim rst As ADODB.Recordset
    Dim llAstPostedCount As Long
    Dim llAstTotalCount As Long
    Dim llCpttCode As Long
    Dim ilSpotsAired As Integer
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim ilRet As Integer
        
    On Error GoTo ErrHand
    
    mSetCpttStatus = False
    
    
    'Test to see if any spots aired or were they all not aired
    ilSpotsAired = gDidAnySpotsAir(lAttCode, sMoDate, sSuDate)
    If ilSpotsAired Then
        'We know at least one spot aired
       ilSpotsAired = True
    Else
        'no aired spots were found
        ilSpotsAired = False
    End If

    SQLQuery = "Select cpttCode from CPTT"
    SQLQuery = SQLQuery + " WHERE"
    SQLQuery = SQLQuery + " cpttAtfCode = " & lAttCode
    SQLQuery = SQLQuery + " AND (cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    
    If Not rst.EOF Then
        llCpttCode = rst!cpttCode
    End If
    
    '12/13/13: Pledge obtained from DAT
    
    ''find the number of posted spots for the week, if any
    'SQLQuery = "Select COUNT(astCode) from AST"
    'SQLQuery = SQLQuery + " WHERE"
    'SQLQuery = SQLQuery + " astCPStatus = 1 AND astPledgeStatus <> 8 AND astPledgeStatus <> 4"
    'SQLQuery = SQLQuery + " AND astAtfCode = " & lAttCode
    'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(Format$(gObtainNextSunday(sMoDate)), sgSQLDateForm) & "')"
    'Set rst = gSQLSelectCall(SQLQuery)
    
    'If Not rst.EOF Then
    '    llAstPostedCount = rst(0).Value
    'End If
    
    ''find the number of spots posted or not
    'SQLQuery = "Select COUNT(astCode) from AST"
    'SQLQuery = SQLQuery + " WHERE"
    'SQLQuery = SQLQuery + " astPledgeStatus <> 8 AND astPledgeStatus <> 4"
    'SQLQuery = SQLQuery + " AND astAtfCode = " & lAttCode
    'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(gObtainNextSunday(sMoDate), sgSQLDateForm) & "')"
    'Set rst = gSQLSelectCall(SQLQuery)
    
    'If Not rst.EOF Then
    '    llAstTotalCount = rst(0).Value
    'End If
    
    llAstPostedCount = 0
    llAstTotalCount = 0
    SQLQuery = "Select * from AST"
    SQLQuery = SQLQuery + " WHERE"
    SQLQuery = SQLQuery + " astAtfCode = " & lAttCode
    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(Format$(gObtainNextSunday(sMoDate)), sgSQLDateForm) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        tlDatPledgeInfo.lAttCode = rst!astAtfCode
        tlDatPledgeInfo.lDatCode = rst!astDatCode
        tlDatPledgeInfo.iVefCode = rst!astVefCode
        tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
        tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
        If rst!astCPStatus = 1 Then
            If tlDatPledgeInfo.iPledgeStatus <> 8 And tlDatPledgeInfo.iPledgeStatus <> 4 Then
                llAstPostedCount = llAstPostedCount + 1
            End If
        End If
        If tlDatPledgeInfo.iPledgeStatus <> 8 And tlDatPledgeInfo.iPledgeStatus <> 4 Then
            llAstTotalCount = llAstTotalCount + 1
        End If
        rst.MoveNext
    Loop
    
    'Fully Posted
    If llAstTotalCount = llAstPostedCount Then
        SQLQuery = "UPDATE cptt SET "
        If ilSpotsAired Then
            'spots aired
        SQLQuery = SQLQuery + "cpttStatus = 1" & ", "
        Else
            'at least some spot(s) aired
            SQLQuery = SQLQuery + "cpttStatus = 2" & ", "
        End If
        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery + "cpttPostingStatus = 2"
        SQLQuery = SQLQuery + " WHERE cpttCode = " & llCpttCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SpotUtil-mSetCpttStatus"
            mSetCpttStatus = False
            Exit Function
        End If
    End If
    
    'Partial Posted
    If llAstTotalCount > llAstPostedCount And llAstPostedCount <> 0 Then
        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery + "cpttStatus = 0" & ", "
        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery + "cpttPostingStatus = 1"
        SQLQuery = SQLQuery + " WHERE cpttCode = " & llCpttCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SpotUtil-mSetCpttStatus"
            mSetCpttStatus = False
            Exit Function
        End If
    
    End If
    
    'Outstanding
    If llAstPostedCount = 0 Then
        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery + "cpttStatus = 0" & ", "
        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & " cpttAstStatus = 'N'" & ", "
        SQLQuery = SQLQuery + " cpttPostingStatus = 0"
        SQLQuery = SQLQuery + " WHERE cpttCode = " & llCpttCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SpotUtil-mSetCpttStatus"
            mSetCpttStatus = False
            Exit Function
        End If
    
    End If
    gFileChgdUpdate "cptt.mkd", True
    mSetCpttStatus = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSetCpttStatus"
End Function


Private Sub mReset()

End Sub

Private Sub edcStartDate_LostFocus()
    txtActiveDate.Text = edcStartDate.Text
    mSyncDates "StartDate"
End Sub

Private Sub edcEndDate_LostFocus()
    mSyncDates "EndDate"
End Sub


Private Sub txtNumberOfDays_LostFocus()
    mSyncDates "NumberOfDays"
End Sub

Private Sub mSyncDates(sWhoChanged As String)

    Dim slDate As String
    Dim ilNumDays As Integer
    
    If sWhoChanged = "NumberOfDays" Then
        If txtNumberOfDays.Text <> "" And edcStartDate.Text <> "" Then
            slDate = DateValue(gAdjYear(edcStartDate.Text))
            slDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slDate))
            edcEndDate.Text = slDate
        End If
    End If
    
    If sWhoChanged = "EndDate" Then
        If edcEndDate.Text <> "" And edcStartDate.Text <> "" Then
            slDate = DateValue(gAdjYear(edcEndDate.Text)) - DateValue(gAdjYear(edcStartDate.Text))
            ilNumDays = DateDiff("d", DateValue(gAdjYear(edcStartDate.Text)), DateValue(gAdjYear(edcEndDate.Text)))
            txtNumberOfDays.Text = ilNumDays + 1
        End If
    End If
    
    If sWhoChanged = "StartDate" Then
        If edcStartDate.Text <> "" And edcStartDate.Text <> "" Then
            If txtNumberOfDays.Text <> "" Then
                slDate = DateValue(gAdjYear(edcStartDate.Text))
                slDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text, slDate))
                edcEndDate.Text = slDate
            Else
                edcEndDate.Text = ""
            End If
        End If
    End If

End Sub

Private Sub mAlighMulticast()
    Dim slStartDate As String
    Dim slEndDate As String
    If lbcStations(0).ListCount <= 0 Then
        Exit Sub
    End If
    If Not rbcSpots(2).Value Then
        'Selective weeks
        slStartDate = Trim$(edcStartDate.Text)
        slEndDate = gAdjYear(DateAdd("d", txtNumberOfDays.Text - 1, slStartDate))
        gAlignMulticastStations imVefCode, "A", lbcStations(1), lbcStations(0), gDateValue(slStartDate), gDateValue(slEndDate)
    Else
        'All Weeks
        gAlignMulticastStations imVefCode, "A", lbcStations(1), lbcStations(0)
    End If
End Sub


Private Function mDeleteMGandMissed(llAttCode As Long, slStartDate As String, slSuDate As String, slAddMissedorMG As Boolean) As Boolean

    '10/16/18 D.S. **** Handle Deletes from the MakeGoods and MissedSpots Table ****
    
    Dim slDataArray() As String
    ReDim slArray(0 To 1) As String
    Dim llTotRecs As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilIdx As Integer
    Dim ilAttIdx As Integer
    Dim llRet As Long
    Dim slSQLQuery As String
    Dim slTempSta As String
    Dim slMoDate As String
    Dim llShttCode As Long
    Dim llVefCode As Long
    Dim att_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    mDeleteMGandMissed = False
    
    slTempSta = gGetCallLettersByAttCode(llAttCode)
    slMoDate = gObtainPrevMonday(slStartDate)
    '***************************************************** Missed Spots **********************************************************
    
    slSQLQuery = "Delete from MissedSpots Where attCode = " & llAttCode
    slSQLQuery = slSQLQuery & " AND PledgeStartDate >= '" & Format$(slStartDate, "mm-dd-yyyy") & "' AND PledgeEndDate <= '" & Format$(slSuDate, "mm-dd-yyyy") & "'"
    llRet = gExecWebSQLWithRowsEffected(slSQLQuery)
    gLogMsg "   " & CStr(llRet) & " Missed Spots were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
     
     
     '**************************************************** MakeGood Spots **********************************************************
     
    llShttCode = gGetShttCodeFromAttCode(CStr(llAttCode))
    llVefCode = gGetVehCodeFromAttCode(CStr(llAttCode))
                  
    slSQLQuery = "Select attCode from Att where attShfCode = " & llShttCode & " And attvefcode = " & imVefCode
    Set att_rst = gSQLSelectCall(slSQLQuery)
    
    llTotRecs = 0
    While Not att_rst.EOF
        'Check "Spots" table for makegoods within the week we are deleting
        slSQLQuery = "Select orgAstCode from Spots"
        slSQLQuery = slSQLQuery + " Where attCode = " & llAttCode
        slSQLQuery = slSQLQuery + " And astCode >= 2000000000"
        slSQLQuery = slSQLQuery + " And PledgeStartDate >= " & "'" & Format$(slStartDate, "mm-dd-yyyy") & "'"
        slSQLQuery = slSQLQuery + " And (recType = 'M' or recType = '" & "MG" & "')"
        'slSQLQuery = slSQLQuery + " And PledgeEndDate <= " & "'" & Format$(slSuDate, "mm-dd-yyyy") & "'"
        llTotRecs = gExecWebSQLForVendor(slDataArray, slSQLQuery, True)
    
        For ilIdx = 1 To llTotRecs - 1
            slArray = Split(gFixQuote(slDataArray(ilIdx)), ",")
            slArray(0) = gGetDataNoQuotes(slArray(0))     'OrgAst Code
            slSQLQuery = "Delete From Makegoods Where orgAstCode = " & slArray(0)
            llRet = gExecWebSQLWithRowsEffected(slSQLQuery)
            slSQLQuery = "Delete From Spots Where orgAstCode = " & slArray(0)
            llRet = gExecWebSQLWithRowsEffected(slSQLQuery)
            slSQLQuery = "Delete From SpotRevisions Where orgAstCode = " & slArray(0)
            llRet = gExecWebSQLWithRowsEffected(slSQLQuery)
            If slAddMissedorMG Then
                llRet = mInsertIntoMissedSpots(slArray(0))
            End If
        Next ilIdx
        llTotRecs = llTotRecs - 1
        If llTotRecs < 0 Then
            llTotRecs = 0
        End If
        gLogMsg "   " & llTotRecs & " Makegood Spots were Deleted on the Web for Agreement " & llAttCode & " " & slTempSta & "," & " w/o " & slMoDate & ". For the Period " & slStartDate & " - " & slSuDate, "AffUtilsLog.Txt", False
        att_rst.MoveNext
    Wend
    mDeleteMGandMissed = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotUtil-mSpotsDelete"
    Exit Function
End Function




Private Function mInsertIntoMissedSpots(mAstCode As String) As Boolean

    '1/8/19 D.S. **** When spots are Deleted from the Web MakeGoods table, get the original spot information from the Spots file. ****
    '            **** Use the original Spot info to Insert it into the Web MissedSpots Table.                                     ****
    
    Dim slSQLQuery As String
    Dim slDataArray() As String
    Dim slArray() As String
    Dim llTotRecs As Long
    Dim ilIdx As Integer

    
    slSQLQuery = "Select astCode,attCode, SpotSeqNum, CpyRotCode, Advt, Prod, TranType,PledgeStartDate,PledgeEndDate,PledgeStartTime,PledgeEndTime,FeedDate,FeedTime,statusCode,SpotLen,Cart,ISCI,CreativeTitle,AvailName,ActualAirDate,ActualAirtime,ActualDateTime,DescrepancyCode,postDate,exportDate,postedFlag,exportedFlag,SentDate,OrgstatusCode,gsfCode,RecType,ROTEndDate,SRLink,EstimatedTime,IsDayPart,EstimatedDay,BeforeOrAfter,TrueDaysPledged,MRReason,OrgAstCode,NewAstCode,srcAttCode,showVehName,Source,FlightDays,CntrNumber,FlightStartTime,FlightEndTime,AdfCode,Blackout,EmbeddedOrROS,isOverridable from Spots where AstCode = " & mAstCode
    llTotRecs = gExecWebSQLForVendor(slDataArray, slSQLQuery, True)
    slArray = Split(slDataArray(1), ",")
    For ilIdx = 0 To UBound(slArray) Step 1
        slArray(ilIdx) = gGetDataNoQuotes(slArray(ilIdx))
    Next ilIdx
    If llTotRecs = 2 Then
        'Note: The commented out fields are not set for missed spots
        slSQLQuery = "Insert into MissedSpots("
        slSQLQuery = slSQLQuery & " AstCode,"           '0
        slSQLQuery = slSQLQuery & " AttCode,"           '1
        slSQLQuery = slSQLQuery & " SpotSeqNum,"        '2
        slSQLQuery = slSQLQuery & " CpyRotCode,"        '3
        slSQLQuery = slSQLQuery & " Advt,"              '4
        slSQLQuery = slSQLQuery & " Prod,"              '5
        slSQLQuery = slSQLQuery & " TranType,"          '6
        slSQLQuery = slSQLQuery & " PledgeStartDate,"   '7
        slSQLQuery = slSQLQuery & " PledgeEndDate,"     '8
        slSQLQuery = slSQLQuery & " PledgeStartTime,"   '9
        slSQLQuery = slSQLQuery & " PledgeEndTime,"     '10
        slSQLQuery = slSQLQuery & " FeedDate,"          '11
        slSQLQuery = slSQLQuery & " FeedTime,"          '12
        slSQLQuery = slSQLQuery & " StatusCode,"        '13
        slSQLQuery = slSQLQuery & " SpotLen,"           '14
        slSQLQuery = slSQLQuery & " Cart,"              '15
        slSQLQuery = slSQLQuery & " ISCI,"              '16
        slSQLQuery = slSQLQuery & " CreativeTitle,"     '17
        slSQLQuery = slSQLQuery & " AvailName,"         '18
        'slSQLQuery = slSQLQuery & " ActualAirDate,"    '19
        'slSQLQuery = slSQLQuery & " ActualAirTime,"    '20
        slSQLQuery = slSQLQuery & " ActualDateTime,"    '21
        'slSQLQuery = slSQLQuery & " DescrepancyCode,"  '22
        slSQLQuery = slSQLQuery & " PostDate,"          '23
        'slSQLQuery = slSQLQuery & " ExportDate,"       '24
        slSQLQuery = slSQLQuery & " PostedFlag,"        '25
        slSQLQuery = slSQLQuery & " ExportedFlag,"      '26
        slSQLQuery = slSQLQuery & " SentDate,"          '27
        slSQLQuery = slSQLQuery & " OrgStatusCode,"     '28
        slSQLQuery = slSQLQuery & " GsfCode,"           '29
        slSQLQuery = slSQLQuery & " RecType,"           '30
        slSQLQuery = slSQLQuery & " RotEndDate,"        '31
        'slSQLQuery = slSQLQuery & " SRLink,"           '32
        slSQLQuery = slSQLQuery & " EstimatedTime,"     '33
        slSQLQuery = slSQLQuery & " IsDayPart,"         '34
        slSQLQuery = slSQLQuery & " EstimatedDay,"      '35
        slSQLQuery = slSQLQuery & " BeforeOrAfter,"     '36
        slSQLQuery = slSQLQuery & " TrueDaysPledged,"   '37
        slSQLQuery = slSQLQuery & " MRReason,"          '38
        'slSQLQuery = slSQLQuery & " OrgAstCode,"       '39
        'slSQLQuery = slSQLQuery & " NewAstCode,"       '40
        slSQLQuery = slSQLQuery & " SrcAttCode,"        '41
        slSQLQuery = slSQLQuery & " ShowVehName,"       '42
        slSQLQuery = slSQLQuery & " Source,"            '43
        slSQLQuery = slSQLQuery & " FlightDays,"        '44
        slSQLQuery = slSQLQuery & " CntrNumber,"        '45
        slSQLQuery = slSQLQuery & " FlightStartTime,"   '46
        slSQLQuery = slSQLQuery & " FlightEndTime,"     '47
        slSQLQuery = slSQLQuery & " AdfCode,"           '48
        slSQLQuery = slSQLQuery & " BlackOut,"          '49
        slSQLQuery = slSQLQuery & " EmbeddedOrROS,"     '50
        slSQLQuery = slSQLQuery & " IsOverridable)"     '51
        slSQLQuery = slSQLQuery & " VALUES ("
        slSQLQuery = slSQLQuery & slArray(0) & ","
        slSQLQuery = slSQLQuery & slArray(1) & ","
        slSQLQuery = slSQLQuery & slArray(2) & ","
        slSQLQuery = slSQLQuery & slArray(3) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(4) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(5) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(6) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(7) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(8) & "',"
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(9) & "',"
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(10) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(11) & "',"
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(12) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(13) & "',"
        slSQLQuery = slSQLQuery & slArray(14) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(15) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(16) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(17) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(18) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(19) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(20) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(21) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(22) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(23) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(24) & "',"
        slSQLQuery = slSQLQuery & slArray(25) & ","
        slSQLQuery = slSQLQuery & slArray(26) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(27) & "',"
        slSQLQuery = slSQLQuery & slArray(28) & ","
        slSQLQuery = slSQLQuery & slArray(29) & ","
        slSQLQuery = slSQLQuery & slArray(30) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(31) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(32) & "',"
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(33) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(34) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(35) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(36) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(37) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(38) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(39) & "',"
        'slSQLQuery = slSQLQuery & "'" & slArray(40) & "',"
        slSQLQuery = slSQLQuery & slArray(41) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(42) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(43) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(44) & "',"
        slSQLQuery = slSQLQuery & slArray(45) & ","
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(46) & "',"
        slSQLQuery = slSQLQuery & "'" & "1899-12-30 " & slArray(47) & "',"
        slSQLQuery = slSQLQuery & slArray(48) & ","
        slSQLQuery = slSQLQuery & slArray(49) & ","
        slSQLQuery = slSQLQuery & "'" & slArray(50) & "',"
        slSQLQuery = slSQLQuery & "'" & slArray(51) & "'"
        slSQLQuery = slSQLQuery & ")"
        llTotRecs = gExecWebSQLWithRowsEffected(slSQLQuery)
    End If
End Function

