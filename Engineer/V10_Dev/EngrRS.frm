VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrRS 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.PictureBox pbcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2340
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame frcAudioSelection 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   90
      TabIndex        =   25
      Top             =   5310
      Visible         =   0   'False
      Width           =   1740
      Begin VB.ListBox lbcATE 
         Height          =   1425
         ItemData        =   "EngrRS.frx":030A
         Left            =   30
         List            =   "EngrRS.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   195
         Width           =   1635
      End
      Begin VB.Label lacAudioType 
         Caption         =   "Audio Type"
         Height          =   225
         Left            =   30
         TabIndex        =   27
         Top             =   -30
         Width           =   1020
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLibNames 
      Height          =   1080
      Left            =   1905
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1905
      _Version        =   393216
      Cols            =   10
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin V10EngineeringDev.CSI_TimeLength ltcSpec 
      Height          =   195
      Left            =   4650
      TabIndex        =   7
      Top             =   465
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "00:00:00"
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_UseHours    =   -1  'True
      CSI_UseTenths   =   0   'False
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
   Begin V10EngineeringDev.CSI_DayPicker dpcSpec 
      Height          =   210
      Left            =   7785
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   370
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_DayOnColor  =   4638790
      CSI_DayOffColor =   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
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
   Begin V10EngineeringDev.CSI_Calendar cccSpec 
      Height          =   240
      Left            =   3315
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   423
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
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
   Begin VB.CommandButton cmcFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   10230
      TabIndex        =   16
      Top             =   60
      Width           =   1335
   End
   Begin VB.PictureBox pbcPTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   120
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   6570
      Width           =   60
   End
   Begin VB.PictureBox pbcPSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   165
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   5205
      Width           =   60
   End
   Begin VB.TextBox edcGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2265
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   1695
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   435
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1770
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   15
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   480
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   60
   End
   Begin VB.CommandButton cmcStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6330
      TabIndex        =   13
      Top             =   6825
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11235
      Top             =   6615
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   12
      Top             =   6825
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResource 
      Height          =   4620
      Left            =   1965
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   735
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   8149
      _Version        =   393216
      Cols            =   15
      ForeColorFixed  =   -2147483640
      BackColorSel    =   8454016
      BackColorBkg    =   16777215
      BackColorUnpopulated=   8454016
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLinesFixed  =   1
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   525
      Left            =   1905
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   926
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAudio 
      Height          =   1080
      Left            =   1905
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5625
      Visible         =   0   'False
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1905
      _Version        =   393216
      Cols            =   10
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frcBusSelection 
      BorderStyle     =   0  'None
      Height          =   5040
      Left            =   75
      TabIndex        =   19
      Top             =   330
      Visible         =   0   'False
      Width           =   1740
      Begin VB.ListBox lbcBGE 
         Height          =   2010
         ItemData        =   "EngrRS.frx":030E
         Left            =   30
         List            =   "EngrRS.frx":0315
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   210
         Width           =   1635
      End
      Begin VB.ListBox lbcBDE 
         Height          =   2205
         ItemData        =   "EngrRS.frx":0321
         Left            =   30
         List            =   "EngrRS.frx":0328
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   2610
         Width           =   1635
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All"
         Height          =   195
         Left            =   1170
         TabIndex        =   20
         Top             =   2370
         Width           =   585
      End
      Begin VB.Label lacBusGroup 
         Caption         =   "Bus Group"
         Height          =   210
         Left            =   30
         TabIndex        =   24
         Top             =   -30
         Width           =   840
      End
      Begin VB.Label lacBuses 
         Caption         =   "Buses"
         Height          =   255
         Left            =   30
         TabIndex        =   23
         Top             =   2355
         Width           =   555
      End
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   3630
      TabIndex        =   17
      Top             =   5370
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Label lacScreen 
      Caption         =   "Time Finder"
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1665
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Left            =   10425
      Picture         =   "EngrRS.frx":0334
      Top             =   6720
      Width           =   480
   End
End
Attribute VB_Name = "EngrRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrRS - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private hmSEE As Integer

Private imFieldChgd As Integer
Private smState As String
Private smCategory As String
Private smType As String
Private imInChg As Integer
Private imBSMode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private lmCharacterWidth As Long
Private imStop As Integer
Private lmPreTime As Long
Private lmPostTime As Long
Private lmOrigLength As Long

Private bmPrinting As Boolean

Private smBusGroups() As String
Private imIgnoreBDEChg As Integer
Private imIgnoreBGEChg As Integer
Private smCurrBSEStamp As String
Private tmCurrBSE() As BSE
Private imBusCodes() As Integer
Private tmSHE As SHE

Private tmDee As DEE
Private tmDHE As DHE

Private lmTimeArray() As Long
Private imStartRow As Integer   'Convert Start Time to Row number (0 thru 1439)
Private imEndRow As Integer
Private smTips() As String

Private imANECodes() As Integer

'Grid Controls
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private smSvSpecGridValue As String

Private imLastGreenSelStartRow As Integer
Private imLastGreenSelEndRow As Integer
Private imLastGreenSelCol As Integer

Private imLastRedSelStartRow As Integer
Private imLastRedSelEndRow As Integer
Private imLastRedSelCol As Integer


Const TYPEINDEX = 0
Const STARTDATEINDEX = 1
Const ENDDATEINDEX = 2
Const DAYSINDEX = 3
Const STARTTIMEINDEX = 4
Const ENDTIMEINDEX = 5
Const LENGTHINDEX = 6



Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If (smType = "B") Or (smType = "C") Then
        If lbcBDE.SelCount <= 0 Then
            cmcFind.Enabled = False
            Exit Sub
        End If
    End If
    If (smType = "A") Or (smType = "C") Then
        If lbcATE.ListIndex < 0 Then
            cmcFind.Enabled = False
            Exit Sub
        End If
    End If
    If Not mCheckFields(True) Then
        cmcFind.Enabled = False
        Exit Sub
    End If
    cmcFind.Enabled = True
End Sub

Private Sub mEnableBox()
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slStr As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim slDate As String
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(EVENTTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdSpec.Row >= grdSpec.FixedRows) And (grdSpec.Row < grdSpec.Rows) And (grdSpec.Col >= 0) And (grdSpec.Col < grdSpec.Cols) Then
        lmEnableRow = grdSpec.Row
        lmEnableCol = grdSpec.Col
        smSvSpecGridValue = grdSpec.TextMatrix(lmEnableRow, lmEnableCol)
        Select Case grdSpec.Col
            Case TYPEINDEX
                If Trim$(grdSpec.text) = "" Then
                    smType = "C"
                    grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Combination"
                    pbcType_Paint
                End If
            Case STARTDATEINDEX
                cccSpec.CSI_AllowBlankDate = False
                cccSpec.CSI_AllowTFN = False
                If Trim$(grdSpec.text) = "" Then
                    'cccSpec.CSI_DefaultDateType = csiNextMonday
                    slDate = DateAdd("d", 1, Format(gNow(), "ddddd"))
                    slDate = gObtainNextMonday(slDate)
                    cccSpec.text = Format(slDate, "ddddd")
                Else
                    cccSpec.text = grdSpec.text
                End If
            Case ENDDATEINDEX
                cccSpec.CSI_AllowBlankDate = False
                cccSpec.CSI_AllowTFN = False
                If grdSpec.text = "" Then
                    If grdSpec.TextMatrix(lmEnableRow, STARTDATEINDEX) <> "" Then
                        slDate = grdSpec.TextMatrix(lmEnableRow, STARTDATEINDEX)
                    Else
                        slDate = Format(gNow(), "ddddd")
                    End If
                    slDate = gObtainNextSunday(slDate)
                    cccSpec.text = slDate
                Else
                    cccSpec.text = grdSpec.text
                End If
            Case DAYSINDEX
                If Trim$(grdSpec.text) = "" Then
                    If grdSpec.TextMatrix(grdSpec.Row, ENDDATEINDEX) <> "" Then
                        If gIsDate(grdSpec.TextMatrix(grdSpec.Row, ENDDATEINDEX)) Then
                            If grdSpec.TextMatrix(grdSpec.Row, STARTDATEINDEX) <> "" Then
                                If gIsDate(grdSpec.TextMatrix(grdSpec.Row, STARTDATEINDEX)) Then
                                    If gDateValue(grdSpec.TextMatrix(grdSpec.Row, STARTDATEINDEX)) + 6 > gDateValue(grdSpec.TextMatrix(grdSpec.Row, ENDDATEINDEX)) Then
                                        slDay = String(7, "N")
                                        ilStartDay = Weekday(grdSpec.TextMatrix(grdSpec.Row, STARTDATEINDEX), vbMonday)
                                        ilEndDay = Weekday(grdSpec.TextMatrix(grdSpec.Row, ENDDATEINDEX), vbMonday)
                                        If ilStartDay <= ilEndDay Then
                                            For ilDay = ilStartDay To ilEndDay Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                        Else
                                            For ilDay = ilStartDay To 7 Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                            For ilDay = 1 To ilEndDay Step 1
                                                Mid$(slDay, ilDay, 1) = "Y"
                                            Next ilDay
                                        End If
                                        grdSpec.text = gDayMap(slDay)
                                    Else
                                        grdSpec.text = "Mo-Su"
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If grdSpec.TextMatrix(grdSpec.Row, STARTDATEINDEX) <> "" Then
                            grdSpec.text = "M-Su"
                        End If
                    End If
                End If
                dpcSpec.text = grdSpec.text
            Case STARTTIMEINDEX
                slStr = grdSpec.text
                If slStr = "" Then
                    lmOrigLength = -1
                Else
                    If grdSpec.TextMatrix(grdSpec.Row, ENDTIMEINDEX) = "" Then
                        lmOrigLength = -1
                    Else
                        llStartTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, STARTTIMEINDEX), False)
                        llEndTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, ENDTIMEINDEX), True)
                        If llStartTime < llEndTime Then
                            lmOrigLength = llEndTime - llStartTime
                        Else
                            lmOrigLength = -1
                        End If
                    End If
                End If
                If Not gIsLength(slStr) Then
                    ltcSpec.text = ""
                Else
                    ltcSpec.text = ""
                    ltcSpec.text = slStr   'grdSpecEvents.Text
                End If
            Case ENDTIMEINDEX
                slStr = grdSpec.text
                If slStr = "" Then
                    lmOrigLength = -1
                Else
                    If grdSpec.TextMatrix(grdSpec.Row, STARTTIMEINDEX) = "" Then
                        lmOrigLength = -1
                    Else
                        llStartTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, STARTTIMEINDEX), False)
                        llEndTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, ENDTIMEINDEX), True)
                        If llStartTime < llEndTime Then
                            lmOrigLength = llEndTime - llStartTime
                        Else
                            lmOrigLength = -1
                        End If
                    End If
                End If
                If Not gIsLength(slStr) Then
                    ltcSpec.text = ""
                Else
                    ltcSpec.text = ""
                    ltcSpec.text = slStr   'grdSpecEvents.Text
                End If
            Case LENGTHINDEX
                slStr = grdSpec.text
                If slStr = "" Then
                    If (grdSpec.TextMatrix(lmEnableRow, STARTTIMEINDEX) <> "") And (grdSpec.TextMatrix(lmEnableRow, ENDTIMEINDEX) <> "") Then
                        llStartTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, STARTTIMEINDEX), False)
                        llEndTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, ENDTIMEINDEX), True)
                        If llStartTime < llEndTime Then
                            llLength = llEndTime - llStartTime
                            slStr = gLongToLength(llLength, True)
                            grdSpec.text = slStr
                        End If
                    End If
                End If
                If Not gIsLength(slStr) Then
                    ltcSpec.text = ""
                Else
                    ltcSpec.text = ""
                    ltcSpec.text = slStr   'grdSpecEvents.Text
                End If
        End Select
        mSetFocus
    End If
End Sub

Private Sub mSetShow()
    If (lmEnableRow >= grdSpec.FixedRows) And (lmEnableRow < grdSpec.Rows) Then
'        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case TYPEINDEX
                If smType = "C" Then
                    grdLibNames.Top = grdSpec.Top
                    grdLibNames.Left = grdSpec.Left
                Else
                    grdLibNames.Top = grdAudio.Top
                    grdLibNames.Left = grdSpec.Left
                End If
            Case STARTDATEINDEX
            Case ENDDATEINDEX
            Case DAYSINDEX
            Case STARTTIMEINDEX
                mAdjLength
            Case ENDTIMEINDEX
                mAdjLength
            Case LENGTHINDEX
        End Select
        If (smSvSpecGridValue <> "") And (smSvSpecGridValue <> grdSpec.TextMatrix(lmEnableRow, lmEnableCol)) Then
            mRemoveGridData
        End If
    End If
    pbcType.Visible = False
    cccSpec.Visible = False
    dpcSpec.Visible = False
    ltcSpec.Visible = False
    edcGrid.Visible = False
    smSvSpecGridValue = ""
End Sub


Private Function mCheckFields(ilSetErrorFlagOnly As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    
    grdSpec.Redraw = False
    ilError = False
    llRow = grdSpec.FixedRows
    
    slStr = grdSpec.TextMatrix(llRow, TYPEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, TYPEINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = TYPEINDEX
            grdSpec.CellForeColor = vbRed
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, STARTDATEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, STARTDATEINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = STARTDATEINDEX
            grdSpec.CellForeColor = vbRed
        End If
    Else
        If Not gIsDate(slStr) Then
            ilError = True
            If Not ilSetErrorFlagOnly Then
                grdSpec.Row = llRow
                grdSpec.Col = STARTDATEINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, ENDDATEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, ENDDATEINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = ENDDATEINDEX
            grdSpec.CellForeColor = vbRed
        End If
    Else
        If Not gIsDate(slStr) Then
            ilError = True
            If Not ilSetErrorFlagOnly Then
                grdSpec.Row = llRow
                grdSpec.Col = ENDDATEINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, DAYSINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, DAYSINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = DAYSINDEX
            grdSpec.CellForeColor = vbRed
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, STARTTIMEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, STARTTIMEINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = STARTTIMEINDEX
            grdSpec.CellForeColor = vbRed
        End If
    Else
        If Not gIsTime(slStr) Then
            ilError = True
            If Not ilSetErrorFlagOnly Then
                grdSpec.Row = llRow
                grdSpec.Col = STARTTIMEINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, ENDTIMEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, ENDTIMEINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = ENDTIMEINDEX
            grdSpec.CellForeColor = vbRed
        End If
    Else
        If Not gIsTime(slStr) Then
            ilError = True
            If Not ilSetErrorFlagOnly Then
                grdSpec.Row = llRow
                grdSpec.Col = ENDTIMEINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    End If
    slStr = grdSpec.TextMatrix(llRow, LENGTHINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If Not ilSetErrorFlagOnly Then
            grdSpec.TextMatrix(llRow, LENGTHINDEX) = "Missing"
            grdSpec.Row = llRow
            grdSpec.Col = LENGTHINDEX
            grdSpec.CellForeColor = vbRed
        End If
    Else
        If Not gIsLength(slStr) Then
            ilError = True
            If Not ilSetErrorFlagOnly Then
                grdSpec.Row = llRow
                grdSpec.Col = LENGTHINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    End If
    grdSpec.Redraw = True
    If ilError Then
        mCheckFields = False
        Exit Function
    Else
        mCheckFields = True
        Exit Function
    End If
End Function


Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    
    gGrid_AlignAllColsLeft grdSpec
    mGridColumnWidth
    'Set Titles
    grdSpec.TextMatrix(0, TYPEINDEX) = "Type"
    grdSpec.TextMatrix(0, STARTDATEINDEX) = "Start Date"
    grdSpec.TextMatrix(0, ENDDATEINDEX) = "End Date"
    grdSpec.TextMatrix(0, DAYSINDEX) = "Days"
    grdSpec.TextMatrix(0, STARTTIMEINDEX) = "Start Time"
    grdSpec.TextMatrix(0, ENDTIMEINDEX) = "End Time"
    grdSpec.TextMatrix(0, LENGTHINDEX) = "Length"
    grdSpec.Row = 1
    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        grdSpec.Col = ilCol
        grdSpec.CellAlignment = flexAlignLeftCenter
    Next ilCol
    'grdSpec.Height = 2 * grdSpec.RowHeight(0) + 15
    'gGrid_IntegralHeight grdSpec
    grdSpec.Height = grdSpec.RowHeight(0) + grdSpec.RowHeight(1)
    'gGrid_Clear grdSpec, True
    
    
    'gGrid_AlignAllColsLeft grdResource
    mGridColumnWidth
    'Set Titles
    'Set Titles
'    For ilCol = BUSNAMEINDEX To BUSCTRLINDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Bus"
'    Next ilCol
'    For ilCol = TIMEINDEX To DURATIONINDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Time"
'    Next ilCol
'    For ilCol = AUDIONAMEINDEX To AUDIOCTRLINDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Audio"
'    Next ilCol
'    For ilCol = BACKUPNAMEINDEX To BACKUPCTRLINDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Backup"
'    Next ilCol
'    For ilCol = PROTNAMEINDEX To PROTCTRLINDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Protection"
'    Next ilCol
'    For ilCol = RELAY1INDEX To RELAY2INDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Relay"
'    Next ilCol
'    For ilCol = SILENCETIMEINDEX To SILENCE4INDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Silence"
'    Next ilCol
'    For ilCol = NETCUE1INDEX To NETCUE2INDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Netcue"
'    Next ilCol
'    For ilCol = TITLE1INDEX To TITLE2INDEX Step 1
'        grdResource.TextMatrix(0, ilCol) = "Title"
'    Next ilCol
'    grdResource.TextMatrix(1, BUSNAMEINDEX) = "Name"
'    grdResource.TextMatrix(1, BUSCTRLINDEX) = "C"
'    grdResource.TextMatrix(1, TIMEINDEX) = "Time"
'    grdResource.TextMatrix(1, STARTTYPEINDEX) = "Start"
'    grdResource.TextMatrix(1, FIXEDINDEX) = "Fix"
'    grdResource.TextMatrix(1, ENDTYPEINDEX) = "End"
'    grdResource.TextMatrix(1, DURATIONINDEX) = "Dur"
'    grdResource.TextMatrix(0, MATERIALINDEX) = "Mat"
'    grdResource.TextMatrix(1, MATERIALINDEX) = "Type"
'    grdResource.TextMatrix(1, AUDIONAMEINDEX) = "Name"
'    grdResource.TextMatrix(1, AUDIOITEMIDINDEX) = "Item"
'    grdResource.TextMatrix(1, AUDIOCTRLINDEX) = "C"
'    grdResource.TextMatrix(1, BACKUPNAMEINDEX) = "Name"
'    grdResource.TextMatrix(1, BACKUPCTRLINDEX) = "C"
'    grdResource.TextMatrix(1, PROTNAMEINDEX) = "Name"
'    grdResource.TextMatrix(1, PROTITEMIDINDEX) = "Item"
'    grdResource.TextMatrix(1, PROTCTRLINDEX) = "C"
'    grdResource.TextMatrix(1, RELAY1INDEX) = "1"
'    grdResource.TextMatrix(1, RELAY2INDEX) = "2"
'    grdResource.TextMatrix(0, FOLLOWINDEX) = "Fol-"
'    grdResource.TextMatrix(1, FOLLOWINDEX) = "low"
'    grdResource.TextMatrix(1, SILENCETIMEINDEX) = "Time"
'    grdResource.TextMatrix(1, SILENCE1INDEX) = "1"
'    grdResource.TextMatrix(1, SILENCE2INDEX) = "2"
'    grdResource.TextMatrix(1, SILENCE3INDEX) = "3"
'    grdResource.TextMatrix(1, SILENCE4INDEX) = "4"
'    grdResource.TextMatrix(1, NETCUE1INDEX) = "Start"
'    grdResource.TextMatrix(1, NETCUE2INDEX) = "Stop"
'    grdResource.TextMatrix(1, TITLE1INDEX) = "1"
'    grdResource.TextMatrix(1, TITLE2INDEX) = "2"
'
'    grdResource.Row = 1
'    For ilCol = 0 To grdResource.Cols - 1 Step 1
'        grdResource.Col = ilCol
'        grdResource.CellAlignment = flexAlignLeftCenter
'    Next ilCol
'    grdResource.Row = 0
'    grdResource.MergeCells = flexMergeRestrictRows
'    grdResource.MergeRow(0) = True
'    grdResource.Row = 0
'    grdResource.Col = BUSNAMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = TIMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = AUDIONAMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = BACKUPNAMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = PROTNAMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = RELAY1INDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = SILENCETIMEINDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = NETCUE1INDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
'    grdResource.Row = 0
'    grdResource.Col = TITLE1INDEX
'    grdResource.CellAlignment = flexAlignCenterCenter
    grdResource.Height = grdAudio.Top - grdResource.Top - 240    '4 * grdResource.RowHeight(0) + 15
    'gGrid_IntegralHeight grdResource
    'gGrid_Clear grdResource, True
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdSpec.ColWidth(TYPEINDEX) = grdSpec.Width / 7
    grdSpec.ColWidth(STARTDATEINDEX) = grdSpec.Width / 7
    grdSpec.ColWidth(ENDDATEINDEX) = grdSpec.Width / 7
    'grdSpec.ColWidth(DAYSINDEX) = grdSpec.Width / 18
    grdSpec.ColWidth(STARTTIMEINDEX) = grdSpec.Width / 7
    grdSpec.ColWidth(ENDTIMEINDEX) = grdSpec.Width / 7
    grdSpec.ColWidth(LENGTHINDEX) = grdSpec.Width / 7
    grdSpec.ColWidth(DAYSINDEX) = grdSpec.Width '- GRIDSCROLLWIDTH
    For ilCol = TYPEINDEX To LENGTHINDEX Step 1
        If ilCol <> DAYSINDEX Then
            If grdSpec.ColWidth(DAYSINDEX) > grdSpec.ColWidth(ilCol) Then
                grdSpec.ColWidth(DAYSINDEX) = grdSpec.ColWidth(DAYSINDEX) - grdSpec.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    
    For ilCol = 1 To grdResource.Cols - 1 Step 1
        grdResource.ColWidth(ilCol) = grdResource.Width / grdResource.Cols
    Next ilCol
    grdResource.ColWidth(0) = grdResource.Width - GRIDSCROLLWIDTH
    For ilCol = 1 To grdResource.Cols - 1 Step 1
        grdResource.ColWidth(0) = grdResource.ColWidth(0) - grdResource.ColWidth(ilCol)
    Next ilCol

    For ilCol = 0 To grdAudio.Cols - 1 Step 1
        grdAudio.ColWidth(ilCol) = grdAudio.Width / grdAudio.Cols
    Next ilCol


End Sub


Private Sub mClearControls()
    gGrid_Clear grdSpec, True
    gGrid_Clear grdResource, True
    gGrid_Clear grdAudio, True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrRS-mPopulate Bus Groups", tgCurrBGE())
    lbcBGE.Clear
    For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
        lbcBGE.AddItem Trim$(tgCurrBGE(ilLoop).sName)
        lbcBGE.ItemData(lbcBGE.NewIndex) = tgCurrBGE(ilLoop).iCode
    Next ilLoop
    
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrRS-mPopulate Bus", tgCurrBDE())
    lbcBDE.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        lbcBDE.AddItem Trim$(tgCurrBDE(ilLoop).sName)
        lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilLoop).iCode
    Next ilLoop
    
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrRS-mPopulat Audio Typee", tgCurrATE())
    lbcATE.Clear
    For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
        lbcATE.AddItem Trim$(tgCurrATE(ilLoop).sName)
        lbcATE.ItemData(lbcATE.NewIndex) = tgCurrATE(ilLoop).iCode
    Next ilLoop
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrRS-mPopulate Audio Names", tgCurrANE())
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrRS-mPopulate Audio Source", tgCurrASE())
    
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrRS-mPopulate Library Names", tgCurrLibDNE())
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "T", sgCurrTempDNEStamp, "EngrRS-mPopulate Template Names", tgCurrTempDNE())
    
    
End Sub

Private Sub cccSpec_Change()
    Dim slOrigStartDate As String
    Dim slOrigEndDate As String
    Dim slCurStartDate As String
    Dim slCurEndDate As String
    Dim ilClearDays As Integer
    
    If StrComp(Trim$(grdSpec.text), Trim$(cccSpec.text), vbTextCompare) <> 0 Then
        slOrigStartDate = grdSpec.TextMatrix(grdSpec.FixedRows, STARTDATEINDEX)
        slOrigEndDate = grdSpec.TextMatrix(grdSpec.FixedRows, ENDDATEINDEX)
        grdSpec.text = cccSpec.text
        grdSpec.CellForeColor = vbBlack
        slCurStartDate = grdSpec.TextMatrix(grdSpec.FixedRows, STARTDATEINDEX)
        slCurEndDate = grdSpec.TextMatrix(grdSpec.FixedRows, ENDDATEINDEX)
        ilClearDays = False
        If IsDate(slOrigStartDate) Then
            If IsDate(slOrigEndDate) Then
                If IsDate(slCurStartDate) Then
                    If IsDate(slCurEndDate) Then
                        If gDateValue(slOrigEndDate) > gDateValue(slOrigStartDate) + 6 Then
                            If gDateValue(slCurEndDate) < gDateValue(slCurStartDate) + 6 Then
                                ilClearDays = True
                            End If
                        Else
                            ilClearDays = True
                        End If
                    End If
                End If
            End If
        End If
        If ilClearDays Then
            grdSpec.TextMatrix(grdSpec.Row, DAYSINDEX) = ""
        End If
    End If
    mSetCommands
End Sub

Private Sub ckcAll_Click()
    Dim iValue As Integer
    Dim lRg As Long
    Dim lRet As Long
    
    If ckcAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    imIgnoreBDEChg = True
    lRg = CLng(lbcBDE.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcBDE.hwnd, LB_SELITEMRANGE, iValue, lRg)
    imIgnoreBDEChg = False
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrRS
End Sub



Private Sub cmcCancel_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub cmcFind_Click()
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slDays As String
    Dim ilRow As Integer
    Dim ilDay As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilBDE As Integer
    Dim llTimeIndex As Long
    Dim ilIndex As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim llDNE As Long
    Dim slLibName As String
    Dim ilPos As Integer
    Dim ilStartHour As Integer
    Dim ilEndHour As Integer
    Dim blConflict As Boolean
    
    If Not mCheckFields(False) Then
        Exit Sub
    End If
    gSetMousePointer grdSpec, grdResource, vbHourglass
    imLastGreenSelStartRow = -1
    imLastRedSelStartRow = -1
    ilRow = grdSpec.FixedRows
    slStartDate = grdSpec.TextMatrix(ilRow, STARTDATEINDEX)
    llStartDate = gDateValue(slStartDate)
    slEndDate = grdSpec.TextMatrix(ilRow, ENDDATEINDEX)
    llEndDate = gDateValue(slEndDate)
    If llEndDate < llStartDate Then
        gSetMousePointer grdSpec, grdResource, vbDefault
        Beep
        MsgBox "End Date prior to Start Date", vbOKOnly + vbInformation, "Date Error"
        grdSpec.Row = grdSpec.FixedRows
        grdSpec.Col = ENDDATEINDEX
        grdSpec.CellForeColor = vbRed
        mEnableBox
        Exit Sub
    End If
    
    slStartTime = grdSpec.TextMatrix(ilRow, STARTTIMEINDEX)
    ilStartHour = Hour(slStartTime)
    llStartTime = 10 * gTimeToLong(slStartTime, False)
    slEndTime = grdSpec.TextMatrix(ilRow, ENDTIMEINDEX)
    ilEndHour = Hour(slEndTime)
    llEndTime = 10 * gTimeToLong(slEndTime, True) - 1
    If llEndTime < llStartTime Then
        gSetMousePointer grdSpec, grdResource, vbDefault
        MsgBox "End Time prior to Start Time", vbOKOnly + vbInformation, "Time Error"
        Beep
        grdSpec.Row = grdSpec.FixedRows
        grdSpec.Col = ENDTIMEINDEX
        grdSpec.CellForeColor = vbRed
        mEnableBox
        Exit Sub
    End If
    cmcStop.Enabled = True
    cmcCancel.Enabled = False
    If ilStartHour < ilEndHour Then
        ilEndHour = ilEndHour - 1
    End If
    imStartRow = 0  'llStartTime \ 600
    imEndRow = 1439 'llEndTime \ 600
    slStr = grdSpec.TextMatrix(ilRow, DAYSINDEX)
    slDays = gCreateDayStr(slStr)
    'Create array of Bus codes to check
    imStop = False
    imcPrint.Enabled = False
    lacProcessing.Caption = "Initializing Result Grid"
    lacProcessing.Visible = True
    DoEvents
    mResourceRowsCols llStartTime \ 600, llEndTime \ 600
    'Loop thru all dates and Days checking for holes
    For llDate = llStartDate To llEndDate Step 1
        If imStop Then
            Exit For
        End If
        slDate = Format$(llDate, "ddddd")
        lacProcessing.Caption = "Processing: " & slDate
        DoEvents
        ilDay = Weekday(slDate, vbMonday)
        If Mid$(slDays, ilDay, 1) = "Y" Then
            ilRet = gGetRec_SHE_ScheduleHeaderByDate(slDate, "EngrRS-Get Schedule by Date", tmSHE)
            If Not ilRet Then
                tmSHE.lCode = 0
                ilRet = gGetEventsFromLibrariesHourRange(slDate, ilStartHour, ilEndHour)
            Else
                ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrRS-Get Events", tgCurrSEE())
            End If
            For ilLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
                If tgCurrSEE(ilLoop).sAction <> "D" Then
                    DoEvents
                    If imStop Then
                        Exit For
                    End If
                    'Test if within range
                    If (tgCurrSEE(ilLoop).lTime >= llStartTime) And (tgCurrSEE(ilLoop).lTime <= llEndTime) Then
                        'Test if matching Bus
                        ilASE = gBinarySearchASE(tgCurrSEE(ilLoop).iAudioAseCode, tgCurrASE())
                        
                        DoEvents
                        If imStop Then
                            Exit For
                        End If
                        'Turn time Off by looping from llStartTime to llEndTime+lDuration
                        ilIndex = (tgCurrSEE(ilLoop).lTime) \ 600
                        slLibName = ""
                        If tgCurrSEE(ilLoop).lDeeCode > 0 Then
                            If tgCurrSEE(ilLoop).lDeeCode <> tmDee.lCode Then
                                ilRet = gGetRec_DEE_DayEvent(tgCurrSEE(ilLoop).lDeeCode, "EngrSchd-mMoveSEERecToCtrls: DEE", tmDee)
                            End If
                            If tmDee.lDheCode <> tmDHE.lCode Then
                                ilRet = gGetRec_DHE_DayHeaderInfo(tmDee.lDheCode, "EngrSchd-mMoveSEERecToCtrls: DHE", tmDHE)
                            End If
                            If tmDHE.sType <> "T" Then
                                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrLibDNE)
                                If llDNE <> -1 Then
                                    slLibName = Trim$(tgCurrLibDNE(llDNE).sName)
                                End If
                            Else
                                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrTempDNE)
                                If llDNE <> -1 Then
                                    slLibName = Trim$(tgCurrTempDNE(llDNE).sName)
                                End If
                            End If
                        End If
                        If (smType = "B") Or (smType = "C") Then
                            For ilBDE = 0 To UBound(imBusCodes) - 1 Step 1
                                If imStop Then
                                    Exit For
                                End If
                                For llTimeIndex = tgCurrSEE(ilLoop).lTime To tgCurrSEE(ilLoop).lTime + tgCurrSEE(ilLoop).lDuration Step 600
                                    If ilIndex >= 1440 Then
                                        Exit For
                                    End If
                                    If tgCurrSEE(ilLoop).iBdeCode = imBusCodes(ilBDE) Then
                                        lmTimeArray(ilIndex, ilBDE, 0) = 0
                                        ilPos = InStr(1, smTips(ilIndex, ilBDE), slLibName, vbTextCompare)
                                        If ilPos <= 0 Then
                                            If Trim$(smTips(ilIndex, ilBDE)) = "" Then
                                                smTips(ilIndex, ilBDE) = slLibName
                                            Else
                                                smTips(ilIndex, ilBDE) = smTips(ilIndex, ilBDE) & "; " & slLibName
                                            End If
                                        End If
                                    End If
                                    ilIndex = ilIndex + 1
                                Next llTimeIndex
                                'Turn time Off by looping from llStartTime to llEndTime+lDuration
                                ilIndex = (tgCurrSEE(ilLoop).lTime - lmPreTime) \ 600
                                For llTimeIndex = tgCurrSEE(ilLoop).lTime - lmPreTime To tgCurrSEE(ilLoop).lTime + tgCurrSEE(ilLoop).lDuration + lmPostTime Step 600
                                    If ilIndex >= 1440 Then
                                        Exit For
                                    End If
                                    For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                                        blConflict = False
                                        If imANECodes(ilANE) = tgCurrASE(ilASE).iPriAneCode Then
                                            'lmTimeArray(ilIndex, ilBDE, ilANE + 1) = 0
                                            blConflict = True
                                        End If
                                        If imANECodes(ilANE) = tgCurrSEE(ilLoop).iProtAneCode Then
                                            'lmTimeArray(ilIndex, ilBDE, ilANE + 1) = 0
                                            blConflict = True
                                        End If
                                        If imANECodes(ilANE) = tgCurrSEE(ilLoop).iBkupAneCode Then
                                            'lmTimeArray(ilIndex, ilBDE, ilANE + 1) = 0
                                            blConflict = True
                                        End If
                                        'If lmTimeArray(ilIndex, ilBDE, ilANE + 1) = 0 Then
                                        If blConflict Then
                                            lmTimeArray(ilIndex, ilBDE, ilANE + 1) = 0
                                            ilPos = InStr(1, smTips(ilIndex, ilBDE), slLibName, vbTextCompare)
                                            If ilPos <= 0 Then
                                                If Trim$(smTips(ilIndex, ilBDE)) = "" Then
                                                    smTips(ilIndex, ilBDE) = slLibName
                                                Else
                                                    smTips(ilIndex, ilBDE) = smTips(ilIndex, ilBDE) & "; " & slLibName
                                                End If
                                            End If
                                        End If
                                    Next ilANE
                                    ilIndex = ilIndex + 1
                                Next llTimeIndex
                            Next ilBDE
                        Else
                            'Turn time Off by looping from llStartTime to llEndTime+lDuration
                            ilIndex = (tgCurrSEE(ilLoop).lTime - lmPreTime) \ 600
                            For llTimeIndex = tgCurrSEE(ilLoop).lTime - lmPreTime To tgCurrSEE(ilLoop).lTime + tgCurrSEE(ilLoop).lDuration + lmPostTime Step 600
                                If ilIndex >= 1440 Then
                                    Exit For
                                End If
                                If imStop Then
                                    Exit For
                                End If
                                For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                                    blConflict = False
                                    If imANECodes(ilANE) = tgCurrASE(ilASE).iPriAneCode Then
                                        blConflict = True
                                    End If
                                    If imANECodes(ilANE) = tgCurrSEE(ilLoop).iProtAneCode Then
                                        blConflict = True
                                    End If
                                    If imANECodes(ilANE) = tgCurrSEE(ilLoop).iBkupAneCode Then
                                        blConflict = True
                                    End If
                                    If blConflict Then
                                        lmTimeArray(ilIndex, ilANE, 0) = 0
                                        ilPos = InStr(1, smTips(ilIndex, ilANE), slLibName, vbTextCompare)
                                        If ilPos <= 0 Then
                                            If Trim$(smTips(ilIndex, ilANE)) = "" Then
                                                smTips(ilIndex, ilANE) = slLibName
                                            Else
                                                smTips(ilIndex, ilANE) = smTips(ilIndex, ilANE) & "; " & slLibName
                                            End If
                                        End If
                                    End If
                                Next ilANE
                                ilIndex = ilIndex + 1
                            Next llTimeIndex
                        End If
                    End If
                End If
            Next ilLoop
            If imStop Then
                Exit For
            End If
        End If
    Next llDate
    cmcCancel.Enabled = True
    cmcStop.Enabled = False
    If imStop Then
        lacProcessing.Caption = "Processing Stopped"
        gSetMousePointer grdSpec, grdResource, vbDefault
        gSetMousePointer grdSpec, grdResource, vbDefault
        Exit Sub
    Else
        lacProcessing.Caption = "Loading Information into Grid"
    End If
    DoEvents
    mPopResourceGrid
    lacProcessing.Caption = ""
    lacProcessing.Visible = False
    gSetMousePointer grdSpec, grdResource, vbDefault
End Sub

Private Sub cmcFind_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub cmcStop_Click()
    imStop = True
End Sub

Private Sub dpcSpec_OnChange()
    If StrComp(Trim$(grdSpec.text), Trim$(dpcSpec.text), vbTextCompare) <> 0 Then
        grdSpec.text = dpcSpec.text
        grdSpec.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
'    Select Case grdSpec.Col
'        Case CATEGORYINDEX
'        Case NAMEINDEX
'            If grdSpec.Text <> edcGrid.Text Then
'                imFieldChgd = True
'            End If
'            grdSpec.Text = edcGrid.Text
'            grdSpec.CellForeColor = vbBlack
'        Case DESCRIPTIONINDEX
'            If grdSpec.Text <> edcGrid.Text Then
'                imFieldChgd = True
'            End If
'            grdSpec.Text = edcGrid.Text
'            grdSpec.CellForeColor = vbBlack
'        Case AUTOCODEINDEX
'            If grdSpec.Text <> edcGrid.Text Then
'                imFieldChgd = True
'            End If
'            grdSpec.Text = edcGrid.Text
'            grdSpec.CellForeColor = vbBlack
'        Case STATEINDEX
'    End Select
'    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        grdSpec.Height = grdSpec.RowHeight(0) + grdSpec.RowHeight(1)
    End If
    imFirstActivate = False
End Sub

Private Sub Form_Click()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    'Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    Me.Move Me.Left, Me.Top, 0.97 * Screen.Width, 0.82 * Screen.Height
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrRS
    'gCenterFormModal EngrRS
    gCenterForm EngrRS
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(RESOURCEJOB) = True
End Sub

Private Sub Form_Resize()
    Dim ilRow As Integer
    
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdResource.Height = grdAudio.Top - lacProcessing.Height - grdResource.Top - 60  '4 * grdResource.RowHeight(0) + 15
    ''gGrid_IntegralHeight grdResource
    ''gGrid_FillWithRows grdResource
    'grdSpec.Height = 2 * grdSpec.RowHeight(0) + 15
    'gGrid_IntegralHeight grdSpec
    grdSpec.Height = grdSpec.RowHeight(0) + grdSpec.RowHeight(1)
    lacProcessing.Top = grdResource.Top + grdResource.Height + 15
    grdAudio.Height = cmcCancel.Top - grdAudio.Top - 300
    grdAudio.Top = lacProcessing.Top + lacProcessing.Height - 30
    'lbcATE.Top = grdAudio.Top
    frcBusSelection.Top = grdResource.Top
    frcAudioSelection.Top = grdAudio.Top
    gGrid_IntegralHeight grdAudio
    gGrid_FillWithRows grdAudio
    grdSpec.Left = grdResource.Left
    grdLibNames.Top = grdSpec.Top
    grdLibNames.Left = grdSpec.Left
    grdLibNames.Height = grdResource.Top - grdSpec.Top - 120
    cmcFind.Top = grdSpec.Top
    imcPrint.Top = cmcCancel.Top
    lmCharacterWidth = CLng(pbcTab.TextWidth("n"))
    pbcPSTab.Left = -400
    pbcPTab.Left = -400
    pbcPTab.Left = -400
    pbcTab.Left = -400
    pbcClickFocus.Left = -400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmSEE
    
    igJobShowing(RESOURCEJOB) = False
    Set EngrRS = Nothing
End Sub





Private Sub mInit()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdAudio, grdResource, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    lmEnableRow = -1
    imLastGreenSelStartRow = -1
    imLastRedSelStartRow = -1
    lmOrigLength = -1
    imFirstActivate = True
    imInChg = True
    imIgnoreBGEChg = False
    imIgnoreBDEChg = False
    imStop = False
    bmPrinting = False
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim smBusGroups(0 To 0) As String
    mPopulate
    imInChg = False
    imFieldChgd = False
    mSetCommands
'    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(EVENTTYPELIST) = 2) Then
'        cmcDone.Enabled = True
'    Else
'        cmcDone.Enabled = False
'    End If
    gSetMousePointer grdAudio, grdResource, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAudio, grdResource, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Relay Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Relay Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub grdAudio_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub grdResource_Click()
    If (grdResource.Col < grdResource.FixedCols) Or (grdResource.Col >= grdResource.Cols) Then
        Exit Sub
    End If

End Sub

Private Sub grdResource_EnterCell()
    mSetShow
End Sub

Private Sub grdResource_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
    If (grdResource.Col < grdResource.FixedCols) Or (grdResource.Col >= grdResource.Cols) Then
        Exit Sub
    End If
End Sub

Private Sub grdResource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    Dim llRow As Long
    Dim llCol As Long
    
    
    grdResource.ToolTipText = ""
    llRow = grdResource.MouseRow
    llCol = grdResource.MouseCol
    If llRow < grdResource.FixedRows Then
        Exit Sub
    End If
    If llCol < grdResource.FixedCols Then
        Exit Sub
    End If
    On Error Resume Next
    grdResource.ToolTipText = smTips(llRow - grdResource.FixedRows, llCol \ 2)
    DoEvents
End Sub

Private Sub grdResource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    Dim ilANE As Integer
    Dim ilCount As Integer
    Dim ilCol As Integer
    Dim ilIndex As Integer
    Dim ilRow As Integer
    Dim ilWidth As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilSelectedRow As Integer
    Dim slTime As String
    Dim llTime As Long
    Dim ilPos As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilStartPos As Integer
    Dim slName As String
    ReDim slLibNames(0 To 0) As String
    
    'If (lmEnableRow < grdSpec.FixedRows) And (lmEnableRow >= grdSpec.Rows) Then
    '    Exit Sub
    'End If
    'Determine if in header
    If y < grdResource.RowHeight(0) Then
        Exit Sub
    End If
    If x < grdResource.ColWidth(0) Then
        Exit Sub
    End If
'    ilFound = gGrid_DetermineRowCol(grdResource, x, y)
'    'If Not ilFound Then
'    '    pbcClickFocus.SetFocus
'    '    Exit Sub
'    'End If
    grdResource.Row = grdResource.MouseRow
    grdResource.Col = grdResource.MouseCol
    If (grdResource.Col < grdResource.FixedCols) Or (grdResource.Col >= grdResource.Cols) Then
        Exit Sub
    End If
    If Trim$(grdResource.TextMatrix(1, 0)) = "" Then
        Exit Sub
    End If
    grdResource.Redraw = False
    ilSelectedRow = grdResource.MouseRow
    ilStartRow = grdResource.Row
    ilEndRow = grdResource.Row
    ilCol = grdResource.Col
    If (grdResource.CellBackColor = vbRed) Or (grdResource.CellBackColor = vbYellow) Then
        If (imLastRedSelStartRow >= grdResource.FixedRows) And (imLastRedSelEndRow <= grdResource.Rows) And (imLastRedSelCol >= 0) And (imLastRedSelCol <= grdResource.Cols) Then
            grdResource.Col = imLastRedSelCol
            For ilRow = imLastRedSelStartRow To imLastRedSelEndRow Step 1
                grdResource.Row = ilRow
                grdResource.CellBackColor = vbRed
            Next ilRow
        End If
        'If smType = "C" Then
        '    grdAudio.Redraw = False
        '    grdAudio.Rows = 2   'Clear merge rows
        '    grdAudio.Redraw = True
        'End If
        imcPrint.Enabled = False
        grdResource.Row = ilStartRow
        grdResource.Row = ilEndRow
        grdResource.Col = ilCol
        Do
            grdResource.Row = ilStartRow
            If (grdResource.CellBackColor = vbGreen) Or (grdResource.CellBackColor = &H8000000F) Or (grdResource.CellBackColor = vbWhite) Then
                ilStartRow = ilStartRow + 1
                Exit Do
            End If
            ilStartRow = ilStartRow - 1
        Loop While ilStartRow > grdResource.FixedRows
        If ilStartRow < grdResource.FixedRows Then
            imLastRedSelStartRow = -1
            grdResource.Redraw = True
            grdLibNames.Visible = False
            Exit Sub
        End If
        Do
            grdResource.Row = ilEndRow
            If (grdResource.CellBackColor = vbGreen) Or (grdResource.CellBackColor = &H8000000F) Or (grdResource.CellBackColor = vbWhite) Then
                ilEndRow = ilEndRow - 1
                Exit Do
            End If
            ilEndRow = ilEndRow + 1
        Loop While ilEndRow < grdResource.Rows - 1
        If ilEndRow >= grdResource.Rows - 1 Then
            imLastRedSelStartRow = -1
            grdResource.Redraw = True
            grdLibNames.Visible = False
            Exit Sub
        End If
        If (imLastRedSelStartRow = ilStartRow) And (imLastRedSelEndRow = ilEndRow) And (imLastRedSelCol = grdResource.Col) Then
            imLastRedSelStartRow = -1
            grdResource.Redraw = True
            grdLibNames.Visible = False
            Exit Sub
        End If
        For ilRow = ilStartRow To ilEndRow Step 1
            grdResource.Row = ilRow
            grdResource.CellBackColor = vbYellow
        Next ilRow
        imLastRedSelStartRow = ilStartRow
        imLastRedSelEndRow = ilEndRow
        imLastRedSelCol = grdResource.Col
        ilCol = grdResource.Col \ 2
        grdResource.Redraw = True
        For ilRow = ilStartRow To ilEndRow Step 1
            ilStartPos = 1
            slStr = Trim$(smTips(ilRow - grdResource.FixedRows, ilCol))
            If slStr <> "" Then
                Do
                    ilPos = InStr(ilStartPos, slStr, ";", vbTextCompare)
                    If ilPos <= 0 Then
                        ilPos = Len(slStr) + 1
                    End If
                    slName = Trim$(Mid(slStr, ilStartPos, ilPos - ilStartPos))
                    ilStartPos = ilPos + 1
                    ilFound = False
                    For ilLoop = 0 To UBound(slLibNames) - 1 Step 1
                        If StrComp(slLibNames(ilLoop), slName, vbTextCompare) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        slLibNames(UBound(slLibNames)) = slName
                        ReDim Preserve slLibNames(0 To UBound(slLibNames) + 1) As String
                    End If
                Loop While ilStartPos < Len(slStr)
            End If
        Next ilRow
        grdLibNames.Redraw = False
        grdLibNames.Cols = UBound(slLibNames) + 1
        grdLibNames.ColWidth(0) = grdResource.ColWidth(0)
        If grdLibNames.Cols < 15 Then
            If grdLibNames.Cols - 1 > 0 Then
                ilWidth = (grdLibNames.Width - 30 * (grdLibNames.Cols - 1) - grdLibNames.ColWidth(0) - GRIDSCROLLWIDTH) / (grdLibNames.Cols - 1)
            Else
                ilWidth = 0
            End If
        Else
            ilWidth = (grdLibNames.Width - 45 * 15 - grdLibNames.ColWidth(0) - GRIDSCROLLWIDTH) / 15
        End If
        For ilCol = 1 To grdLibNames.Cols - 1 Step 1
            grdLibNames.ColWidth(ilCol) = ilWidth
        Next ilCol
        grdLibNames.RowHeight(1) = 90
        For ilLoop = 0 To UBound(slLibNames) - 1 Step 1
            grdLibNames.Row = 0
            grdLibNames.Col = ilLoop + 1
            grdLibNames.text = Trim$(slLibNames(ilLoop))
            grdLibNames.CellAlignment = flexAlignLeftCenter
        Next ilLoop
        'Merge Column zero so that hours can show
        For ilCol = 0 To UBound(slLibNames) - 1 Step 1
            grdLibNames.MergeCol(ilCol + 1) = False
        Next ilCol
        grdLibNames.Rows = 2   'Clear merge rows
        grdLibNames.Rows = ilEndRow - ilStartRow + grdLibNames.FixedRows + 1
        For ilRow = grdLibNames.FixedRows To grdLibNames.Rows - 1 Step 1
            grdLibNames.RowHeight(ilRow) = 90
        Next ilRow
        
        llTime = 600 * CLng(imLastRedSelStartRow - 1)
        For ilLoop = ilStartRow To ilEndRow Step 5
            grdLibNames.MergeCol(0) = True
            slStr = gLongToStrTimeInTenth(llTime)
            ilPos = InStr(1, slStr, ".", vbTextCompare)
            If ilPos > 0 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            slTime = Format$(slStr, "hh:mm")
            If ilLoop + 2 <= ilEndRow Then
                For ilRow = ilLoop To ilLoop + 2 Step 1
                    grdLibNames.TextMatrix(ilRow - ilStartRow + grdLibNames.FixedRows, 0) = slTime
                    grdLibNames.MergeRow(ilRow - ilStartRow + grdLibNames.FixedRows) = True
                Next ilRow
                If ilLoop + 3 <= ilEndRow Then
                    grdLibNames.MergeRow(ilLoop + 3 - ilStartRow + grdLibNames.FixedRows) = False
                End If
                grdLibNames.MergeCells = 4
                llTime = llTime + 5 * 600
            Else
                Exit For
            End If
        Next ilLoop
        ilCol = grdResource.Col \ 2
        For ilLoop = 0 To UBound(slLibNames) - 1 Step 1
            For ilRow = ilStartRow To ilEndRow Step 1
                slStr = smTips(ilRow - grdResource.FixedRows, ilCol)
                ilPos = InStr(1, slStr, slLibNames(ilLoop), vbTextCompare)
                grdLibNames.Row = ilRow - ilStartRow + grdLibNames.FixedRows
                grdLibNames.Col = ilLoop + 1
                If ilPos >= 1 Then
                    grdLibNames.CellBackColor = vbRed
                Else
                    grdLibNames.CellBackColor = vbWhite
                End If
            Next ilRow
        Next ilLoop
        grdLibNames.Redraw = True
        grdLibNames.Visible = True
    Else
        If smType = "C" Then
            If (imLastGreenSelStartRow >= grdResource.FixedRows) And (imLastGreenSelEndRow <= grdResource.Rows) And (imLastGreenSelCol >= 0) And (imLastGreenSelCol <= grdResource.Cols) Then
                grdResource.Col = imLastGreenSelCol
                For ilRow = imLastGreenSelStartRow To imLastGreenSelEndRow Step 1
                    grdResource.Row = ilRow
                    grdResource.CellBackColor = vbGreen
                Next ilRow
            End If
            grdAudio.Redraw = False
            grdAudio.Rows = 2   'Clear merge rows
            For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                grdAudio.Row = 1
                grdAudio.Col = ilANE + 1
                grdAudio.CellBackColor = vbWhite
            Next ilANE
            grdAudio.TextMatrix(1, 0) = ""
            grdAudio.Redraw = True
            imcPrint.Enabled = False
            grdResource.Row = ilStartRow
            grdResource.Row = ilEndRow
            grdResource.Col = ilCol
            Do
                grdResource.Row = ilStartRow
                If (grdResource.CellBackColor = vbRed) Or (grdResource.CellBackColor = vbYellow) Or (grdResource.CellBackColor = vbWhite) Then
                    ilStartRow = ilStartRow + 1
                    Exit Do
                End If
                ilStartRow = ilStartRow - 1
            Loop While ilStartRow > grdResource.FixedRows
            If ilStartRow < grdResource.FixedRows Then
                imLastRedSelStartRow = -1
                grdResource.Redraw = True
                grdLibNames.Visible = False
                Exit Sub
            End If
            Do
                grdResource.Row = ilEndRow
                If (grdResource.CellBackColor = vbRed) Or (grdResource.CellBackColor = vbYellow) Or (grdResource.CellBackColor = vbWhite) Then
                    ilEndRow = ilEndRow - 1
                    Exit Do
                End If
                ilEndRow = ilEndRow + 1
            Loop While ilEndRow < grdResource.Rows - 1
            If ilEndRow >= grdResource.Rows - 1 Then
                imLastRedSelStartRow = -1
                grdResource.Redraw = True
                grdLibNames.Visible = False
                Exit Sub
            End If
            If (imLastGreenSelStartRow = ilStartRow) And (imLastGreenSelEndRow = ilEndRow) And (imLastGreenSelCol = grdResource.Col) Then
                imLastGreenSelStartRow = -1
                grdResource.Redraw = True
                If smType = "C" Then
                    For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                        grdAudio.Row = grdAudio.FixedRows
                        grdAudio.Col = ilANE + 1
                        grdAudio.CellBackColor = vbWhite
                    Next ilANE
                    grdAudio.TextMatrix(grdAudio.FixedRows, 0) = ""
                    grdAudio.Redraw = True
                End If
                Exit Sub
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                If (ilRow >= grdResource.FixedRows) And (ilRow < grdResource.Rows) Then
                    grdResource.Row = ilRow
                    grdResource.CellBackColor = &H8000000F
                End If
            Next ilRow
            imLastGreenSelStartRow = ilStartRow
            imLastGreenSelEndRow = ilEndRow
            imLastGreenSelCol = grdResource.Col
            ilCol = grdResource.Col \ 2
            grdResource.Redraw = True
            grdAudio.Redraw = False
            grdAudio.Rows = 2   'Clear merge rows
            If ilEndRow - ilStartRow + grdAudio.FixedRows + 1 > 2 Then
                grdAudio.Rows = ilEndRow - ilStartRow + grdAudio.FixedRows + 1
            End If
            For ilRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
                grdAudio.RowHeight(ilRow) = 90
            Next ilRow
            
            llTime = 600 * CLng(imLastGreenSelStartRow - 1)
            For ilLoop = ilStartRow To ilEndRow Step 5
                grdAudio.MergeCol(0) = True
                slStr = gLongToStrTimeInTenth(llTime)
                ilPos = InStr(1, slStr, ".", vbTextCompare)
                If ilPos > 0 Then
                    slStr = Left$(slStr, ilPos - 1)
                End If
                slTime = Format$(slStr, "hh:mm")
                If ilLoop + 2 <= ilEndRow Then
                    For ilRow = ilLoop To ilLoop + 2 Step 1
                        grdAudio.TextMatrix(ilRow - ilStartRow + grdAudio.FixedRows, 0) = slTime
                        grdAudio.MergeRow(ilRow - ilStartRow + grdAudio.FixedRows) = True
                    Next ilRow
                    If ilLoop + 3 <= ilEndRow Then
                        grdAudio.MergeRow(ilLoop + 3 - ilStartRow + grdAudio.FixedRows) = False
                    End If
                    grdAudio.MergeCells = 4
                    llTime = llTime + 5 * 600
                Else
                    Exit For
                End If
            Next ilLoop
            For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                For ilRow = ilStartRow To ilEndRow Step 1
                    ilIndex = gBinarySearchANE(imANECodes(ilANE), tgCurrANE())
                    If ilIndex <> -1 Then
                        grdAudio.Row = ilRow - ilStartRow + grdAudio.FixedRows
                        grdAudio.Col = ilANE + 1
                        If lmTimeArray(ilRow - 1, imLastGreenSelCol \ 2, ilANE + 1) = 1 Then
                            grdAudio.CellBackColor = vbGreen
                        Else
                            grdAudio.CellBackColor = vbRed
                        End If
                    End If
                Next ilRow
            Next ilANE
            grdAudio.Redraw = True
        End If
        imcPrint.Enabled = True
    End If
    
End Sub





Private Sub grdSpec_Click()
    If grdSpec.Col >= grdSpec.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdSpec_EnterCell()
    mSetShow
End Sub

Private Sub grdSpec_GotFocus()
    If grdSpec.Col >= grdSpec.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdSpec_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    grdSpec.Redraw = False
End Sub

Private Sub grdSpec_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdSpec.RowHeight(0) Then
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdSpec, x, y)
    If Not ilFound Then
        grdSpec.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdSpec.Col >= grdSpec.Cols Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    
'    llRow = grdSpec.Row
'    If grdSpec.TextMatrix(llRow, CATEGORYINDEX) = "" Then
'        grdSpec.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdSpec.TextMatrix(llRow, CATEGORYINDEX) = ""
'        grdSpec.Row = llRow + 1
'        grdSpec.Col = CATEGORYINDEX
'        grdSpec.Redraw = True
'    End If
    grdSpec.Redraw = True
    mEnableBox
End Sub

Private Sub imcPrint_Click()
    Dim ilCurrentLineNo As Integer
    Dim ilLinesPerPage As Integer
    Dim slRecord As String
    Dim slHeading As String
    Dim ilBDE As Integer
    Dim slPrevTime As String
    Dim ilANE As Integer
    Dim ilIndex As Integer
    Dim ilRow As Integer
    Dim ilRet As Integer
    Dim ilPageNo As Integer
    Dim ilTimeRow As Integer
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilPos As Integer
    Dim slHour As String
    Dim slMin As String
    Dim slStr As String
    
    If bmPrinting Then
        Exit Sub
    End If
    bmPrinting = True
    lacProcessing.Visible = True
    lacProcessing.Caption = "Printing"
    DoEvents
    ilPageNo = 0
    ilCurrentLineNo = 0
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    ilRet = 0
    On Error GoTo imcPrtErr:
    slHeading = "     Resource Search Information for " & Trim$(tgUIE.sShowName) & " on " & Format$(Now, "ddddd") & " at " & Format$(Now, "ttttt")
    GoSub mHeading1
    If ilRet <> 0 Then
        bmPrinting = False
        Printer.EndDoc
        On Error GoTo 0
        lacProcessing.Visible = False
        Exit Sub
    End If
    slRecord = "     Selection: " & grdSpec.TextMatrix(grdSpec.FixedRows, STARTDATEINDEX) & "-" & grdSpec.TextMatrix(grdSpec.FixedRows, ENDDATEINDEX) _
                               & " " & grdSpec.TextMatrix(grdSpec.FixedRows, STARTTIMEINDEX) & "-" & grdSpec.TextMatrix(grdSpec.FixedRows, ENDTIMEINDEX)
    GoSub mLineOutput
    If ilRet <> 0 Then
        bmPrinting = False
        Printer.EndDoc
        On Error GoTo 0
        lacProcessing.Visible = False
        Exit Sub
    End If
    'Output Information
    For ilBDE = 0 To UBound(imBusCodes) - 1 Step 1
        If ilBDE = imLastGreenSelCol \ 2 Then
            slPrevTime = ""
            For ilRow = 0 To UBound(smTips, 1) - 1 Step 1
                If (ilRow >= imLastGreenSelStartRow - grdResource.FixedRows) And (ilRow <= imLastGreenSelEndRow - grdResource.FixedRows) Then
                    'If Trim$(smTips(ilRow, ilBDE)) <> "" Then
                        If StrComp(slPrevTime, Trim$(smTips(ilRow, ilBDE)), vbTextCompare) <> 0 Then
                            slPrevTime = Trim$(smTips(ilRow, ilBDE))
                            slRecord = "     " & grdResource.TextMatrix(0, 2 * ilBDE + 1) & " " & slPrevTime
                            GoSub mLineOutput
                            If ilRet <> 0 Then
                                bmPrinting = False
                                Printer.EndDoc
                                On Error GoTo 0
                                lacProcessing.Visible = False
                                Exit Sub
                            End If
                            'Show Audio
                            ilPos = InStr(1, slPrevTime, "-", vbTextCompare)
                            If ilPos > 0 Then
                                slStr = Left$(slPrevTime, ilPos - 1)
                                ilHour = Hour(slStr)
                                ilMin = Minute(slStr)
                            Else
                                ilHour = 0
                                ilMin = 0
                            End If
                            For ilTimeRow = ilRow To UBound(smTips) - 1 Step 1
                                If StrComp(slPrevTime, Trim$(smTips(ilTimeRow, ilBDE)), vbTextCompare) = 0 Then
                                    slRecord = ""
                                    slHour = Trim$(Str$(ilHour))
                                    If Len(slHour) = 1 Then
                                        slHour = "0" & slHour
                                    End If
                                    slMin = Trim$(Str$(ilMin))
                                    If Len(slMin) = 1 Then
                                        slMin = "0" & slMin
                                    End If
                                    slRecord = "         " & slHour & ":" & slMin
                                    ilMin = ilMin + 1
                                    If ilMin >= 60 Then
                                        ilHour = ilHour + 1
                                        ilMin = 0
                                    End If
                                    For ilANE = 0 To UBound(imANECodes) - 1 Step 1
                                        If lmTimeArray(ilTimeRow, ilBDE, ilANE + 1) = 1 Then
                                            ilIndex = gBinarySearchANE(imANECodes(ilANE), tgCurrANE())
                                            If ilIndex <> -1 Then
                                                If slRecord = "" Then
                                                    slRecord = "              " & Trim$(tgCurrANE(ilIndex).sName)
                                                Else
                                                    slRecord = slRecord & " " & Trim$(tgCurrANE(ilIndex).sName)
                                                End If
                                                If Len(slRecord) > 80 Then
                                                    GoSub mLineOutput
                                                    If ilRet <> 0 Then
                                                        bmPrinting = False
                                                        Printer.EndDoc
                                                        On Error GoTo 0
                                                        lacProcessing.Visible = False
                                                        Exit Sub
                                                    End If
                                                    slRecord = ""
                                                End If
                                            End If
                                        End If
                                    Next ilANE
                                    If slRecord <> "" Then
                                        GoSub mLineOutput
                                        If ilRet <> 0 Then
                                            bmPrinting = False
                                            Printer.EndDoc
                                            On Error GoTo 0
                                            lacProcessing.Visible = False
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next ilTimeRow
                        End If
                    'End If
                    Exit For
                End If
            Next ilRow
            Exit For
        End If
    Next ilBDE
    Printer.EndDoc
    On Error GoTo 0
    lacProcessing.Visible = False
    bmPrinting = False
    Exit Sub
mHeading1:  'Output file name and date
    ilPageNo = ilPageNo + 1
    Printer.Print slHeading & "    Page # " & ilPageNo
    If ilRet <> 0 Then
        Return
    End If
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Return
mLineOutput:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        If ilRet <> 0 Then
            Return
        End If
        ilCurrentLineNo = 0
        GoSub mHeading1
        If ilRet <> 0 Then
            Return
        End If
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Return
imcPrtErr:
    ilRet = Err.Number
    MsgBox "Printing Error # " & Str$(ilRet) & " " & Err.Description
    Resume Next
End Sub

Private Sub lbcATE_Click()
    mRemoveGridData
    mSetCommands
End Sub

Private Sub lbcATE_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub lbcBDE_Click()
    Dim iValue As Integer
    Dim lRg As Long
    Dim lRet As Long
    
    If imIgnoreBDEChg Then
        Exit Sub
    End If
    mRemoveGridData
    If lbcBDE.SelCount <= 1 Then
        imIgnoreBGEChg = True
        ckcAll.Value = vbUnchecked
        lRg = CLng(lbcBGE.ListCount - 1) * &H10000 Or 0
        iValue = False
        lRet = SendMessageByNum(lbcBGE.hwnd, LB_SELITEMRANGE, iValue, lRg)
        ReDim smBusGroups(0 To 0) As String
        imIgnoreBGEChg = False
    End If
    mSetCommands
End Sub

Private Sub lbcBDE_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub lbcBGE_Click()
    If imIgnoreBGEChg Then
        Exit Sub
    End If
    grdLibNames.Visible = False
    mRemoveGridData
    gSetMousePointer grdSpec, grdResource, vbHourglass
    mSetBusesFromBusGroup
    mSetCommands
    gSetMousePointer grdSpec, grdResource, vbDefault
End Sub

Private Sub lbcBGE_GotFocus()
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
    End If
End Sub

Private Sub lbcBGE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBGE, y)
    If (llRow < lbcBGE.ListCount) And (lbcBGE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBGE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
            If ilCode = tgCurrBGE(ilLoop).iCode Then
                lbcBGE.ToolTipText = Trim$(tgCurrBGE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub ltcSpec_OnChange()
    Dim slStr As String
    
    slStr = ltcSpec.text
    If grdSpec.text <> slStr Then
        grdSpec.text = slStr
        grdSpec.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcPSTab_GotFocus()
    If GetFocus() <> pbcPSTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Then
'        mESetShow
'        If grdResource.Col = BUSNAMEINDEX Then
'            If grdResource.Row > grdResource.FixedRows Then
'                grdResource.Row = grdResource.Row - 1
'                grdResource.Col = TITLE2INDEX
'                mEEnableBox
'            Else
'                cmcCancel.SetFocus
'            End If
'        Else
'            grdResource.Col = grdResource.Col - 1
'            mEEnableBox
'        End If
    Else
'        grdResource.Col = BUSNAMEINDEX
'        grdResource.Row = grdResource.FixedRows
'        mEEnableBox
    End If
End Sub

Private Sub pbcPTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcPTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Then
        If grdResource.Col = LENGTHINDEX Then
            If grdResource.Row >= grdResource.Rows - 1 Then
'                If grdSpec.Col = STATEINDEX Then
'                    llRow = grdSpec.Rows
'                    Do
'                        llRow = llRow - 1
'                    Loop While grdSpec.TextMatrix(llRow, CATEGORYINDEX) = ""
'                    llRow = llRow + 1
'                    If (grdSpec.Row + 1 < llRow) Then
'                        lmTopRow = -1
'                        grdSpec.Row = grdSpec.Row + 1
'                        If Not grdSpec.RowIsVisible(grdSpec.Row) Then
'                            grdSpec.TopRow = grdSpec.TopRow + 1
'                        End If
'                        grdSpec.Col = CATEGORYINDEX
'                        'grdSpec.TextMatrix(grdSpec.Row, CODEINDEX) = 0
'                        If Trim$(grdSpec.TextMatrix(grdSpec.Row, CATEGORYINDEX)) <> "" Then
'                            mEnableBox
'                        Else
'                            imFromArrow = True
'                            pbcArrow.Move grdSpec.Left - pbcArrow.Width - 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + (grdSpec.RowHeight(grdSpec.Row) - pbcArrow.Height) / 2
'                            pbcArrow.Visible = True
'                            pbcArrow.SetFocus
'                        End If
'                    Else
'                        If Trim$(grdSpec.TextMatrix(lmEnableRow, CATEGORYINDEX)) <> "" Then
'                            lmTopRow = -1
'                            If grdSpec.Row + 1 >= grdSpec.Rows Then
'                                grdSpec.AddItem ""
'                            End If
'                            grdSpec.Row = grdSpec.Row + 1
'                            If Not grdSpec.RowIsVisible(grdSpec.Row) Then
'                                grdSpec.TopRow = grdSpec.TopRow + 1
'                            End If
'                            grdSpec.Col = CATEGORYINDEX
'                            grdSpec.TextMatrix(grdSpec.Row, CODEINDEX) = 0
'                            'mEnableBox
'                            imFromArrow = True
'                            pbcArrow.Move grdSpec.Left - pbcArrow.Width - 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + (grdSpec.RowHeight(grdSpec.Row) - pbcArrow.Height) / 2
'                            pbcArrow.Visible = True
'                            pbcArrow.SetFocus
'                        Else
'                            pbcClickFocus.SetFocus
'                        End If
'                    End If
'                    Exit Sub
'                Else
'                    mEnableBox
'                End If
            Else
'                grdResource.Row = grdResource.Row + 1
'                grdResource.Col = BUSNAMEINDEX
'                mEEnableBox
            End If
        Else
'            grdResource.Col = grdResource.Col + 1
'            mEEnableBox
        End If
    Else
'        grdResource.Col = BUSNAMEINDEX
'        grdResource.Row = grdResource.FixedRows
'        mEEnableBox
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
        If grdSpec.Col = TYPEINDEX Then
            pbcClickFocus.SetFocus
        Else
            grdSpec.Col = grdSpec.Col - 1
            mEnableBox
        End If
    Else
        grdSpec.TopRow = grdSpec.FixedRows
        grdSpec.Col = TYPEINDEX
        grdSpec.Row = grdSpec.FixedRows
        mEnableBox
    End If
End Sub


Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If (pbcType.Visible) Or (edcGrid.Visible) Or (cccSpec.Visible) Or (dpcSpec.Visible) Or (ltcSpec.Visible) Then
        mSetShow
        If grdSpec.Col = LENGTHINDEX Then
            pbcClickFocus.SetFocus
        Else
            grdSpec.Col = grdSpec.Col + 1
            mEnableBox
        End If
    Else
        grdSpec.TopRow = grdSpec.FixedRows
        grdSpec.Col = LENGTHINDEX
        grdSpec.Row = grdSpec.FixedRows
        mEnableBox
    End If
End Sub

Private Sub mSetFocus()
    Select Case grdSpec.Col
        Case TYPEINDEX
            pbcType.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            pbcType.Visible = True
            pbcType.SetFocus
        Case STARTDATEINDEX
            cccSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            cccSpec.Visible = True
            cccSpec.SetFocus
        Case ENDDATEINDEX
            cccSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            cccSpec.Visible = True
            cccSpec.SetFocus
        Case DAYSINDEX
            dpcSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            dpcSpec.MaxLength = 0
            dpcSpec.Visible = True
            dpcSpec.SetFocus
        Case STARTTIMEINDEX  'Date
            ltcSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            ltcSpec.Width = gSetCtrlWidth("TIME", lmCharacterWidth, ltcSpec.Width, 0)
            ltcSpec.Visible = True
            ltcSpec.SetFocus
        Case ENDTIMEINDEX  'Date
            ltcSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            ltcSpec.Width = gSetCtrlWidth("TIME", lmCharacterWidth, ltcSpec.Width, 0)
            ltcSpec.Visible = True
            ltcSpec.SetFocus
        Case LENGTHINDEX  'Date
            ltcSpec.Move grdSpec.Left + grdSpec.ColPos(grdSpec.Col) + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 30
            ltcSpec.Width = gSetCtrlWidth("TIME", lmCharacterWidth, ltcSpec.Width, 0)
            ltcSpec.Visible = True
            ltcSpec.SetFocus
    End Select
End Sub






Private Sub mSetBusesFromBusGroup()
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilBusGroup As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilBGECode As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    Dim ilBSE As Integer
    Dim ilBus As Integer
    Dim ilRet As Integer
    ReDim ilNewGroupBus(0 To 0) As Integer
    ReDim ilOldBusSel(0 To 0) As Integer
    ReDim ilOldGroupBus(0 To 0) As Integer
    
    'Get current selected buses
    For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
        If lbcBDE.Selected(ilLoop) Then
            ilOldBusSel(UBound(ilOldBusSel)) = lbcBDE.ItemData(ilLoop)
            ReDim Preserve ilOldBusSel(0 To UBound(ilOldBusSel) + 1) As Integer
        End If
    Next ilLoop
    
    'Get list of buses that could have been highlighted from the Groups that were previously selected
    For ilBusGroup = LBound(smBusGroups) To UBound(smBusGroups) - 1 Step 1
        slStr = Trim$(smBusGroups(ilBusGroup))
        If slStr <> "" Then
            llRow = gListBoxFind(lbcBGE, slStr)
            If llRow >= 0 Then
                ilBGECode = lbcBGE.ItemData(llRow)
                ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
                For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
                    ilFound = False
                    For ilBus = 0 To UBound(ilOldGroupBus) - 1 Step 1
                        If ilOldGroupBus(ilBus) = tmCurrBSE(ilBSE).iBdeCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilBus
                    If Not ilFound Then
                        ilOldGroupBus(UBound(ilOldGroupBus)) = tmCurrBSE(ilBSE).iBdeCode
                        ReDim Preserve ilOldGroupBus(0 To UBound(ilOldGroupBus) + 1) As Integer
                    End If
                Next ilBSE
                'If ilFound Then
                '    Exit For
                'End If
            End If
        End If
    Next ilBusGroup
    ReDim smBusGroups(0 To 0) As String
    'Get Buses from current selected Groups
    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
        If lbcBGE.Selected(ilLoop) Then
            smBusGroups(UBound(smBusGroups)) = lbcBGE.List(ilLoop)
            ReDim Preserve smBusGroups(0 To UBound(smBusGroups) + 1) As String
            ilFound = False
            ilBGECode = lbcBGE.ItemData(ilLoop)
            ilRet = gGetRecs_BSE_BusSelGroup("G", smCurrBSEStamp, ilBGECode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
            For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
                ilFound = False
                For ilBus = 0 To UBound(ilNewGroupBus) - 1 Step 1
                    If tmCurrBSE(ilBSE).iBdeCode = ilNewGroupBus(ilBus) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilBus
                If Not ilFound Then
                    ilNewGroupBus(UBound(ilNewGroupBus)) = tmCurrBSE(ilBSE).iBdeCode
                    ReDim Preserve ilNewGroupBus(0 To UBound(ilNewGroupBus) + 1) As Integer
                End If
            Next ilBSE
        End If
    Next ilLoop
    'De-select items from old bus groups
    imIgnoreBDEChg = True
    For ilLoop = 0 To UBound(ilOldGroupBus) - 1 Step 1
        For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
            If ilOldGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                lbcBDE.Selected(ilBDE) = False
                Exit For
            End If
        Next ilBDE
    Next ilLoop
    For ilLoop = 0 To UBound(ilNewGroupBus) - 1 Step 1
        ilFound = False
        For ilBus = 0 To UBound(ilOldGroupBus) - 1 Step 1
            If ilNewGroupBus(ilLoop) = ilOldGroupBus(ilBus) Then
                ilFound = True
                Exit For
            End If
        Next ilBus
        If Not ilFound Then
            For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                If ilNewGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                    lbcBDE.Selected(ilBDE) = True
                    Exit For
                End If
            Next ilBDE
        Else
            'Was it Selected previously
            ilFound = False
            For ilBus = 0 To UBound(ilOldBusSel) - 1 Step 1
                If ilNewGroupBus(ilLoop) = ilOldBusSel(ilBus) Then
                    ilFound = True
                    Exit For
                End If
            Next ilBus
            If ilFound Then
                For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
                    If ilNewGroupBus(ilLoop) = lbcBDE.ItemData(ilBDE) Then
                        lbcBDE.Selected(ilBDE) = True
                        Exit For
                    End If
                Next ilBDE
            End If
        End If
    Next ilLoop
    imIgnoreBDEChg = False
End Sub

Private Sub mResourceRowsCols(ilStartRow As Integer, ilEndRow As Integer)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilATE As Integer
    Dim ilATECode As Integer
    Dim ilANE As Integer
    Dim llWidth As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilTime As Integer
    Dim ilIndex As Integer
    
    'Create array of Bus codes to check
    grdAudio.Clear
    grdResource.Clear
    ReDim imBusCodes(0 To 0) As Integer
    If (smType = "B") Or (smType = "C") Then
        For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
            If lbcBDE.Selected(ilBDE) Then
                imBusCodes(UBound(imBusCodes)) = lbcBDE.ItemData(ilBDE)
                ReDim Preserve imBusCodes(0 To UBound(imBusCodes) + 1) As Integer
            End If
        Next ilBDE
    End If
    ReDim imANECodes(0 To 0) As Integer
    If (smType = "A") Or (smType = "C") Then
        If lbcATE.ListIndex >= 0 Then
            ilATECode = lbcATE.ItemData(lbcATE.ListIndex)
            For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                If tgCurrANE(ilANE).iAteCode = ilATECode Then
                    For ilRow = 0 To UBound(imANECodes) - 1 Step 1
                        If imANECodes(ilRow) = tgCurrANE(ilANE).iCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilRow
                    If Not ilFound Then
                        imANECodes(UBound(imANECodes)) = tgCurrANE(ilANE).iCode
                        ReDim Preserve imANECodes(0 To UBound(imANECodes) + 1) As Integer
                    End If
                End If
            Next ilANE
            For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    If tgCurrANE(ilANE).iAteCode = ilATECode Then
                        For ilRow = 0 To UBound(imANECodes) - 1 Step 1
                            If imANECodes(ilRow) = tgCurrANE(ilANE).iCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilRow
                        If Not ilFound Then
                            imANECodes(UBound(imANECodes)) = tgCurrANE(ilANE).iCode
                            ReDim Preserve imANECodes(0 To UBound(imANECodes) + 1) As Integer
                        End If
                    End If
                End If
            Next ilASE
        End If
    End If
    If (smType = "B") Or (smType = "C") Then
        grdResource.Cols = 2 * UBound(imBusCodes) + 1
    Else
        grdResource.Cols = 2 * UBound(imANECodes) + 1
    End If
'    If (grdResource.Cols \ 2) < 15 Then
'        llWidth = (grdResource.Width - 45 * (grdResource.Cols) - GRIDSCROLLWIDTH) / (grdResource.Cols \ 2)
'    Else
'        llWidth = (grdResource.Width - 45 * 15 - GRIDSCROLLWIDTH) / 15
'    End If
'    For ilCol = 1 To grdResource.Cols - 1 Step 1
'        If ilCol Mod 2 <> 1 Then
'            grdResource.ColWidth(ilCol) = 15
'        Else
'            grdResource.ColWidth(ilCol) = llWidth
'        End If
'    Next ilCol
'    grdResource.ColWidth(0) = grdResource.Width - GRIDSCROLLWIDTH
'    If (grdResource.Cols \ 2) < 15 Then
'        For ilCol = 1 To grdResource.Cols - 1 Step 1
'            If grdResource.ColWidth(0) - grdResource.ColWidth(ilCol) > 360 Then
'                grdResource.ColWidth(0) = grdResource.ColWidth(0) - grdResource.ColWidth(ilCol)
'            Else
'                Exit For
'            End If
'        Next ilCol
'    Else
'        For ilCol = 1 To grdResource.Cols - 1 Step 1
'            If grdResource.ColWidth(0) > grdResource.ColWidth(ilCol) Then
'                If grdResource.ColWidth(0) - grdResource.ColWidth(ilCol) > 360 Then
'                    grdResource.ColWidth(0) = grdResource.ColWidth(0) - grdResource.ColWidth(ilCol)
'                Else
'                    Exit For
'                End If
'            Else
'                Exit For
'            End If
'        Next ilCol
'    End If
    grdResource.ColWidth(0) = 735
    If grdResource.Cols \ 2 < 15 Then
        llWidth = (grdResource.Width - 30 * (grdResource.Cols - 1) - grdResource.ColWidth(0) - GRIDSCROLLWIDTH) / (grdResource.Cols \ 2)
    Else
        llWidth = (grdResource.Width - 30 * 15 - grdResource.ColWidth(0) - GRIDSCROLLWIDTH) / 15
    End If
    For ilCol = 1 To grdResource.Cols - 1 Step 1
        If ilCol Mod 2 <> 1 Then
            grdResource.ColWidth(ilCol) = 15
        Else
            grdResource.ColWidth(ilCol) = llWidth
        End If
    Next ilCol
    llWidth = 0
    For ilCol = 1 To grdResource.Cols - 1 Step 1
        llWidth = llWidth + grdResource.ColWidth(ilCol)
    Next ilCol
    llWidth = llWidth + GRIDSCROLLWIDTH + 30
    If grdResource.Width - llWidth > grdResource.ColWidth(0) Then
        grdResource.ColWidth(0) = grdResource.Width - llWidth
    End If
    ilCol = 1
    If (smType = "B") Or (smType = "C") Then
        For ilBDE = 0 To lbcBDE.ListCount - 1 Step 1
            If lbcBDE.Selected(ilBDE) Then
                'grdResource.ColWidth(ilCol) = llWidth
                grdResource.TextMatrix(0, ilCol) = Trim$(lbcBDE.List(ilBDE))
                ilCol = ilCol + 1
                'If ilCol < grdResource.Cols Then
                '    grdResource.ColWidth(ilCol) = 15
                    ilCol = ilCol + 1
                'End If
            End If
        Next ilBDE
    Else
        For ilANE = 0 To UBound(imANECodes) - 1 Step 1
            ilIndex = gBinarySearchANE(imANECodes(ilANE), tgCurrANE())
            If ilIndex <> -1 Then
                grdResource.TextMatrix(0, ilCol) = Trim$(tgCurrANE(ilIndex).sName)
                ilCol = ilCol + 1
                'If ilCol < grdResource.Cols Then
                '    grdResource.ColWidth(ilCol) = 15
                    ilCol = ilCol + 1
            End If
        Next ilANE
    End If
    grdResource.Rows = 1441 '2
    For ilRow = grdResource.FixedRows To (imEndRow - imStartRow) + grdResource.FixedRows Step 1
        'If ilRow >= grdResource.Rows Then
        '    grdResource.AddItem ""
        'End If
        grdResource.RowHeight(ilRow) = 15
    Next ilRow
    'Merge Column zero so that hours can show
    If (smType = "B") Or (smType = "C") Then
        For ilCol = 0 To UBound(imBusCodes) - 1 Step 1
            grdResource.MergeCol(ilCol + 1) = False
        Next ilCol
    Else
        For ilCol = 0 To UBound(imANECodes) - 1 Step 1
            grdResource.MergeCol(ilCol + 1) = False
        Next ilCol
    End If
    ilTime = 0
    For ilLoop = 1 To 1381 Step 60
        grdResource.MergeCol(0) = True
        For ilRow = ilLoop To ilLoop + 19 Step 1
            grdResource.TextMatrix(ilRow, 0) = ilTime
            'For ilCol = 0 To UBound(imBusCodes) - 1 Step 1
            '    grdResource.TextMatrix(ilRow, ilCol + 1) = Str$(ilRow)
            'Next ilCol
            grdResource.MergeRow(ilRow) = True
        Next ilRow
        grdResource.MergeRow(ilLoop + 20) = False
        grdResource.MergeCells = 4
        ilTime = ilTime + 1
    Next ilLoop
    grdResource.GridLines = flexGridNone
'    'Remove extra row at bottom
'    grdResource.RemoveItem grdResource.Rows
        
    If (smType = "B") Or (smType = "C") Then
        ReDim lmTimeArray(0 To 1439, 0 To UBound(imBusCodes), 0 To UBound(imANECodes)) As Long
        ReDim smTips(0 To imEndRow - imStartRow, 0 To UBound(imBusCodes)) As String
        'Initialize array
        For ilRow = 0 To 1439 Step 1
            For ilBDE = 0 To UBound(lmTimeArray, 2) Step 1
                For ilANE = 0 To UBound(lmTimeArray, 3) Step 1
                    If (ilRow >= ilStartRow) And (ilRow <= ilEndRow) Then
                        lmTimeArray(ilRow, ilBDE, ilANE) = 1
                    Else
                        lmTimeArray(ilRow, ilBDE, ilANE) = -1
                    End If
                Next ilANE
            Next ilBDE
        Next ilRow
    Else
        ReDim lmTimeArray(0 To 1439, 0 To UBound(imANECodes), 0 To UBound(imBusCodes)) As Long
        ReDim smTips(0 To imEndRow - imStartRow, 0 To UBound(imANECodes)) As String
        'Initialize array
        For ilRow = 0 To 1439 Step 1
            For ilANE = 0 To UBound(lmTimeArray, 2) Step 1
                For ilBDE = 0 To UBound(lmTimeArray, 3) Step 1
                    If (ilRow >= ilStartRow) And (ilRow <= ilEndRow) Then
                        lmTimeArray(ilRow, ilANE, ilBDE) = 1
                    Else
                        lmTimeArray(ilRow, ilANE, ilBDE) = -1
                    End If
                Next ilBDE
            Next ilANE
        Next ilRow
    End If
    If (smType = "C") Then
        grdAudio.Cols = UBound(imANECodes) + 1
        grdAudio.Rows = 2
        grdAudio.ColWidth(0) = grdResource.ColWidth(0)
        If grdAudio.Cols < 15 Then
            llWidth = (grdAudio.Width - 30 * (grdAudio.Cols - 1) - grdAudio.ColWidth(0) - GRIDSCROLLWIDTH) / (grdAudio.Cols - 1)
        Else
            llWidth = (grdAudio.Width - 45 * 15 - grdAudio.ColWidth(0) - GRIDSCROLLWIDTH) / 15
        End If
        For ilCol = 1 To grdAudio.Cols - 1 Step 1
            grdAudio.ColWidth(ilCol) = llWidth
        Next ilCol
        grdAudio.RowHeight(1) = 90
        For ilANE = 0 To UBound(imANECodes) - 1 Step 1
            ilIndex = gBinarySearchANE(imANECodes(ilANE), tgCurrANE())
            If ilIndex <> -1 Then
                grdAudio.Row = 0
                grdAudio.Col = ilANE + 1
                grdAudio.text = Trim$(tgCurrANE(ilIndex).sName)
                grdAudio.CellAlignment = flexAlignLeftCenter
            End If
        Next ilANE
        'Merge Column zero so that hours can show
        For ilCol = 0 To UBound(imANECodes) - 1 Step 1
            grdAudio.MergeCol(ilCol + 1) = False
        Next ilCol
    End If
End Sub

Private Sub mPopResourceGrid()
    Dim ilATE As Integer
    Dim ilATECode As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilTest As Integer
    Dim slLength As String
    Dim llLength As Long
    Dim ilBDE As Integer
    Dim ilANE As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slTime As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim llTopRow As Long
    
    grdResource.Redraw = False
    slLength = grdSpec.TextMatrix(grdSpec.FixedRows, LENGTHINDEX)
    llLength = 10 * gLengthToLong(slLength)
    lmPreTime = 0
    lmPostTime = 0
    If lbcATE.ListIndex >= 0 Then
        ilATECode = lbcATE.ItemData(lbcATE.ListIndex)
        ilATE = gBinarySearchATE(ilATECode, tgCurrATE())
        If ilATE <> -1 Then
            lmPreTime = tgCurrATE(ilATE).lPreBufferTime
            lmPostTime = tgCurrATE(ilATE).lPostBufferTime
        End If
    End If
    llTopRow = -1
    llLength = llLength '+ lmPreTime + lmPostTime
    'ReDim smTips(0 To imEndRow - imStartRow, 0 To UBound(imBusCodes)) As String
    If (smType = "B") Or (smType = "C") Then
        For ilBDE = 0 To UBound(imBusCodes) - 1 Step 1
            If ilBDE = 0 Then
                grdResource.Col = 1
            Else
                grdResource.Col = 2 * ilBDE + 1
            End If
            ilRow = imStartRow
            Do
                grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                If lmTimeArray(ilRow, ilBDE, 0) = 1 Then
                    If llTopRow = -1 Then
                        llTopRow = ilRow - imStartRow + grdResource.FixedRows
                    End If
                    'See if enought time
                    ilStartRow = ilRow
                    ilEndRow = ilRow
                    For ilTest = ilRow + 1 To imEndRow Step 1
                        If (lmTimeArray(ilTest, ilBDE, 0) = 1) Then
                            ilEndRow = ilTest
                        Else
                            Exit For
                        End If
                    Next ilTest
                    llStartTime = 600 * CLng(ilStartRow)
                    llEndTime = 600 * CLng(ilEndRow)
                    If (llEndTime - llStartTime) >= (llLength - 600) Then
                        slStr = gLongToStrTimeInTenth(llStartTime)
                        ilPos = InStr(1, slStr, ".", vbTextCompare)
                        If ilPos > 0 Then
                            slStr = Left$(slStr, ilPos - 1)
                        End If
                        slTime = Format$(slStr, "hh:mm")
                        If llEndTime < 863400 Then
                            slStr = gLongToStrTimeInTenth(llEndTime + 600)
                            ilPos = InStr(1, slStr, ".", vbTextCompare)
                            If ilPos > 0 Then
                                slStr = Left$(slStr, ilPos - 1)
                            End If
                            slTime = slTime & "-" & Format$(slStr, "hh:mm")
                        Else
                            slTime = slTime & "-" & "24:00"
                        End If
                        ilRow = ilStartRow
                        Do
                            grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                            grdResource.CellBackColor = vbGreen
                            'If ilBDE < UBound(imBusCodes) - 1 Then
                            If ilBDE < UBound(imBusCodes) Then
                                grdResource.Col = grdResource.Col + 1
                                grdResource.CellBackColor = vbBlack
                                grdResource.Col = grdResource.Col - 1
                            End If
                            smTips(ilRow - imStartRow, ilBDE) = slTime
                            ilRow = ilRow + 1
                        Loop While ilRow <= ilEndRow
                    Else
                        ilRow = ilStartRow
                        Do
                            grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                            grdResource.CellBackColor = vbRed
                            If ilBDE < UBound(imBusCodes) Then
                                grdResource.Col = grdResource.Col + 1
                                grdResource.CellBackColor = vbBlack
                                grdResource.Col = grdResource.Col - 1
                            End If
                            ilRow = ilRow + 1
                        Loop While ilRow <= ilEndRow
                    End If
                ElseIf lmTimeArray(ilRow, ilBDE, 0) = -1 Then
                    grdResource.CellBackColor = vbWhite
                    If ilBDE < UBound(imBusCodes) Then
                        grdResource.Col = grdResource.Col + 1
                        grdResource.CellBackColor = vbBlack
                        grdResource.Col = grdResource.Col - 1
                    End If
                    ilRow = ilRow + 1
                Else
                    grdResource.CellBackColor = vbRed
                    If ilBDE < UBound(imBusCodes) Then
                        grdResource.Col = grdResource.Col + 1
                        grdResource.CellBackColor = vbBlack
                        grdResource.Col = grdResource.Col - 1
                    End If
                    ilRow = ilRow + 1
                End If
            Loop While ilRow <= imEndRow
        Next ilBDE
    Else
        For ilANE = 0 To UBound(imANECodes) - 1 Step 1
            If ilANE = 0 Then
                grdResource.Col = 1
            Else
                grdResource.Col = 2 * ilANE + 1
            End If
            ilRow = imStartRow
            Do
                grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                If lmTimeArray(ilRow, ilANE, 0) = 1 Then
                    If llTopRow = -1 Then
                        llTopRow = ilRow - imStartRow + grdResource.FixedRows
                    End If
                    'See if enought time
                    ilStartRow = ilRow
                    ilEndRow = ilRow
                    For ilTest = ilRow + 1 To imEndRow Step 1
                        If (lmTimeArray(ilTest, ilANE, 0) = 1) Then
                            ilEndRow = ilTest
                        Else
                            Exit For
                        End If
                    Next ilTest
                    llStartTime = 600 * CLng(ilStartRow)
                    llEndTime = 600 * CLng(ilEndRow)
                    If (llEndTime - llStartTime) >= (llLength - 600) Then
                        slStr = gLongToStrTimeInTenth(llStartTime)
                        ilPos = InStr(1, slStr, ".", vbTextCompare)
                        If ilPos > 0 Then
                            slStr = Left$(slStr, ilPos - 1)
                        End If
                        slTime = Format$(slStr, "hh:mm")
                        If llEndTime < 863400 Then
                            slStr = gLongToStrTimeInTenth(llEndTime + 600)
                            ilPos = InStr(1, slStr, ".", vbTextCompare)
                            If ilPos > 0 Then
                                slStr = Left$(slStr, ilPos - 1)
                            End If
                            slTime = slTime & "-" & Format$(slStr, "hh:mm")
                        Else
                            slTime = slTime & "-" & "24:00"
                        End If
                        ilRow = ilStartRow
                        Do
                            grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                            grdResource.CellBackColor = vbGreen
                            'If ilANE < UBound(imANECodes) - 1 Then
                            If ilANE < UBound(imANECodes) Then
                                grdResource.Col = grdResource.Col + 1
                                grdResource.CellBackColor = vbBlack
                                grdResource.Col = grdResource.Col - 1
                            End If
                            smTips(ilRow - imStartRow, ilANE) = slTime
                            ilRow = ilRow + 1
                        Loop While ilRow <= ilEndRow
                    Else
                        ilRow = ilStartRow
                        Do
                            grdResource.Row = ilRow - imStartRow + grdResource.FixedRows
                            grdResource.CellBackColor = vbRed
                            If ilANE < UBound(imANECodes) Then
                                grdResource.Col = grdResource.Col + 1
                                grdResource.CellBackColor = vbBlack
                                grdResource.Col = grdResource.Col - 1
                            End If
                            ilRow = ilRow + 1
                        Loop While ilRow <= ilEndRow
                    End If
                ElseIf lmTimeArray(ilRow, ilANE, 0) = -1 Then
                    grdResource.CellBackColor = vbWhite
                    If ilANE < UBound(imANECodes) Then
                        grdResource.Col = grdResource.Col + 1
                        grdResource.CellBackColor = vbBlack
                        grdResource.Col = grdResource.Col - 1
                    End If
                    ilRow = ilRow + 1
                Else
                    grdResource.CellBackColor = vbRed
                    If ilANE < UBound(imANECodes) Then
                        grdResource.Col = grdResource.Col + 1
                        grdResource.CellBackColor = vbBlack
                        grdResource.Col = grdResource.Col - 1
                    End If
                    ilRow = ilRow + 1
                End If
            Loop While ilRow <= imEndRow
        Next ilANE
    End If
    If llTopRow <> -1 Then
        If llTopRow - 5 >= grdResource.FixedRows Then
            llTopRow = llTopRow - 5
        End If
        grdResource.TopRow = llTopRow
    End If
    grdResource.Redraw = True
End Sub

Private Sub pbcType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("B") Or (KeyAscii = Asc("B")) Then
        If smType <> "B" Then
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Bus Only"
            mRemoveGridData
        End If
        smType = "B"
        pbcType_Paint
    ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("A")) Then
        If smType <> "A" Then
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Audio Only"
            mRemoveGridData
        End If
        smType = "A"
        pbcType_Paint
    ElseIf KeyAscii = Asc("C") Or (KeyAscii = Asc("C")) Then
        If smType <> "C" Then
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Combination"
            mRemoveGridData
        End If
        smType = "C"
        pbcType_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smType = "B" Then
            imFieldChgd = True
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Audio Only"
            smType = "A"
            pbcType_Paint
            mRemoveGridData
        ElseIf smType = "A" Then
            imFieldChgd = True
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Combination"
            smType = "C"
            pbcType_Paint
            mRemoveGridData
        ElseIf smType = "C" Then
            imFieldChgd = True
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Bus Only"
            smType = "B"
            pbcType_Paint
            mRemoveGridData
        Else
            imFieldChgd = True
            grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Bus Only"
            smType = "B"
            pbcType_Paint
            mRemoveGridData
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smType = "B" Then
        grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Audio Only"
        smType = "A"
        pbcType_Paint
        mRemoveGridData
    ElseIf smType = "A" Then
        grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Combination"
        smType = "C"
        pbcType_Paint
        mRemoveGridData
    ElseIf smType = "C" Then
        grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Bus Only"
        smType = "B"
        pbcType_Paint
        mRemoveGridData
    Else
        grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) = "Bus Only"
        smType = "B"
        pbcType_Paint
        mRemoveGridData
    End If
    mSetCommands
End Sub

Private Sub pbcType_Paint()
    pbcType.Cls
    pbcType.CurrentX = 30  'fgBoxInsetX
    pbcType.CurrentY = 0 'fgBoxInsetY
    If smType = "B" Then
        pbcType.Print "Bus Only"
        frcBusSelection.Visible = True
        frcAudioSelection.Visible = False
        grdAudio.Visible = False
    ElseIf smType = "A" Then
        pbcType.Print "Audio Only"
        frcAudioSelection.Visible = True
        frcBusSelection.Visible = False
        frcAudioSelection.Top = grdResource.Top
        grdAudio.Visible = False
    ElseIf smType = "C" Then
        pbcType.Print "Combination"
        frcBusSelection.Visible = True
        frcAudioSelection.Visible = True
        frcAudioSelection.Top = grdAudio.Top
        grdAudio.Visible = True
    Else
        pbcType.Print ""
        frcBusSelection.Visible = False
        frcAudioSelection.Visible = False
        grdAudio.Visible = False
    End If
End Sub

Private Sub mRemoveGridData()
    grdLibNames.Visible = False
    grdResource.Rows = 2
    grdResource.FixedRows = 1
    grdResource.Clear
    grdAudio.Rows = 2
    grdAudio.FixedRows = 1
    grdAudio.Clear
End Sub

Private Sub mAdjLength()
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slStr As String
    
    If lmOrigLength <> -1 Then
        llStartTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, STARTTIMEINDEX), False)
        llEndTime = gTimeToLong(grdSpec.TextMatrix(lmEnableRow, ENDTIMEINDEX), True)
        If llStartTime < llEndTime Then
            If ((llEndTime - llStartTime) < lmOrigLength) Then
                slStr = gLongToLength((llEndTime - llStartTime), True)
                grdSpec.TextMatrix(lmEnableRow, LENGTHINDEX) = slStr
            ElseIf ((llEndTime - llStartTime) > lmOrigLength) Then
                slStr = grdSpec.TextMatrix(lmEnableRow, LENGTHINDEX)
                If slStr <> "" Then
                    If gLengthToLong(slStr) = lmOrigLength Then
                        slStr = gLongToLength((llEndTime - llStartTime), True)
                        grdSpec.TextMatrix(lmEnableRow, LENGTHINDEX) = slStr
                    End If
                End If
            End If
        End If
        lmOrigLength = -1
    End If

End Sub
