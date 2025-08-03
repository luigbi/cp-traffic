VERSION 5.00
Begin VB.Form CModel 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Model"
   ClientHeight    =   4020
   ClientLeft      =   2865
   ClientTop       =   3315
   ClientWidth     =   5925
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   4950
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "CMODEL.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   9
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   10
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcRateCard 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   150
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3225
      TabIndex        =   26
      Top             =   3630
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
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
      Height          =   165
      Left            =   -105
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2760
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   4590
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4590
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1680
      TabIndex        =   24
      Top             =   3630
      Width           =   1050
   End
   Begin VB.PictureBox plcModel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3120
      Left            =   120
      ScaleHeight     =   3060
      ScaleWidth      =   5610
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   5670
      Begin VB.Frame frcNTR 
         Height          =   615
         Left            =   0
         TabIndex        =   27
         Top             =   1940
         Width           =   5295
         Begin VB.CheckBox ckcNTR 
            Caption         =   "Include NTR"
            Height          =   255
            Left            =   90
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame frcSports 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   105
         TabIndex        =   20
         Top             =   1485
         Width           =   5235
         Begin VB.OptionButton rbcSports 
            Caption         =   "Retain Line, Remove Events"
            Height          =   240
            Index           =   1
            Left            =   1155
            TabIndex        =   23
            Top             =   270
            Width           =   2985
         End
         Begin VB.OptionButton rbcSports 
            Caption         =   "Retain Line and Events On and After Date"
            Height          =   240
            Index           =   0
            Left            =   1155
            TabIndex        =   22
            Top             =   45
            Value           =   -1  'True
            Width           =   3885
         End
         Begin VB.Label lacSports 
            Caption         =   "Sports Lines"
            Height          =   210
            Left            =   0
            TabIndex        =   21
            Top             =   30
            Width           =   1230
         End
      End
      Begin VB.PictureBox plcPriceInfo 
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
         Height          =   750
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   5280
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   750
         Width           =   5280
         Begin VB.OptionButton rbcPrice 
            Caption         =   "Rate Card Price"
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
            Height          =   195
            Index           =   1
            Left            =   3525
            TabIndex        =   15
            Top             =   45
            Width           =   1695
         End
         Begin VB.OptionButton rbcPrice 
            Caption         =   "Existing Price"
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
            Height          =   195
            Index           =   0
            Left            =   1935
            TabIndex        =   14
            Top             =   45
            Width           =   1500
         End
         Begin VB.TextBox edcRateCard 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Left            =   975
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   405
            Width           =   3045
         End
         Begin VB.CommandButton cmcRateCard 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4005
            Picture         =   "CMODEL.frx":2E1A
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   390
            Width           =   195
         End
         Begin VB.Label lacRateCard 
            Appearance      =   0  'Flat
            Caption         =   "Rate Card"
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
            Height          =   195
            Left            =   0
            TabIndex        =   16
            Top             =   390
            Width           =   990
         End
         Begin VB.Label lacPrice 
            Caption         =   "For Air/Pod Spots Use"
            Height          =   180
            Left            =   -15
            TabIndex        =   13
            Top             =   30
            Width           =   1890
         End
      End
      Begin VB.CheckBox ckcBook 
         Caption         =   "Replace Old Research Book(s) with Latest (Default) Book(s)"
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   5220
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Left            =   1110
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   1230
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2325
         Picture         =   "CMODEL.frx":2F14
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   105
         Width           =   195
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   120
         Width           =   990
      End
   End
End
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract model screen code
Option Explicit
Option Compare Text
Dim smStartDate As String
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imBypassFocus As Integer
Dim imLbcArrowSetting As Integer
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer
Dim tmRcf As RCF
Dim imDGSave(0 To 7) As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim ghf_rst As ADODB.Recordset
Dim gsf_rst As ADODB.Recordset


Private Sub ckcBook_GotFocus()
    plcCalendar.Visible = vbFalse
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    igTerminateReturn = 0
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcDate_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim llDate As Long
    Dim slDate As String
    Dim slPriceDate As String
    Dim ilNoDaysChg As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilRcfCode As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilCgf As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim llWkEndDate As Long
    Dim llMoFlStartdate As Long
    Dim llPrevWkEndDate As Long
    Dim llWkPrice As Long
    Dim ilRcf As Integer
    Dim ilFirstWk As Integer
    Dim llPrice As Long
    Dim ilUpper As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llSpots As Long
    Dim ilDay As Integer
    Dim ilDays(0 To 6) As Integer
    Dim ilVef As Integer
    Dim slSQLQuery As String
    Dim ilPrevCff As Integer
    Dim ilPrevCgf As Integer
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim llMoSeasonStart As Long
    Dim llSuSeasonEnd As Long
    Dim llSeasonDate As Long
    Dim llAirDate As Long
    Dim ilNoSpots As Integer
    Dim llRow As Long

    slDate = edcDate.Text
    If Not gValidDate(slDate) Then
        Beep
        edcDate.SetFocus
        Exit Sub
    End If
    If gWeekDayStr(slDate) <> 0 Then
        Beep
        ilRet = MsgBox("Start Date must be a Monday", vbOKOnly + vbExclamation, "Error")
        edcDate.SetFocus
        Exit Sub
    End If
    ilRcfCode = -1
    If Contract.plcScreen <> "Orders" Then
        If rbcPrice(1).Value Then
            If lbcRateCard.ListIndex < 0 Then
                Beep
                edcRateCard.SetFocus
                Exit Sub
            End If
            slNameCode = tmCntRateCard(lbcRateCard.ListIndex).sKey    'Traffic!lbcRateCardCode.List(lbcRateCard.ListIndex)' - 1)    'lbcRateCard.ListCount - lbcRateCard.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            ilRcfCode = Val(slCode)
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                If tgMRcf(ilRcf).iCode = ilRcfCode Then
                    tmRcf = tgMRcf(ilRcf)
                    tgChfCntr.iRcfCode = ilRcfCode
                    Exit For
                End If
            Next ilRcf
        End If
    End If
    Screen.MousePointer = vbHourglass
    ilNoDaysChg = gDateValue(slDate) - gDateValue(smStartDate)
    '10525
    If ckcNTR.Value = vbChecked Then
        bgModelNTR = True
    End If
    '2/17/19: Handle sports
    If frcSports.Visible Then
        If rbcSports(1).Value Then
            llDate = 99999999
        Else
            llDate = gDateValue(slDate)
        End If
        Do While ilClf < UBound(tgClfCntr)
            ilRet = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilRet <> -1 Then
                If tgMVef(ilRet).sType = "G" Then
                    'Determine which season to be used and match by event # instead of dates
                    slSQLQuery = "Select * From ghf_Game_Header Where ghfVefCode = " & tgClfCntr(ilClf).ClfRec.iVefCode
                    'slSQLQuery = slSQLQuery & " And ghfSeasonStartDate <= '" & Format$(slDate, sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & " And ghfSeasonEndDate >= '" & Format$(slDate, sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & " Order By ghfSeasonStartDate"
                    Set ghf_rst = gSQLSelectCall(slSQLQuery)
                    If (Not ghf_rst.EOF) And (rbcSports(0).Value) Then
                        llSeasonStart = gDateValue(ghf_rst!ghfSeasonStartDate)
                        llSeasonEnd = gDateValue(ghf_rst!ghfSeasonEndDate)
                        llMoSeasonStart = gDateValue(gObtainPrevMonday(Format(llSeasonStart, "m/d/yy")))
                        llSuSeasonEnd = gDateValue(gObtainNextSunday(Format(llSeasonEnd, "m/d/yy")))
                    
                        'Match by event #
                        tgClfCntr(ilClf).ClfRec.lghfcode = ghf_rst!ghfCode
                        ilCgf = tgClfCntr(ilClf).iFirstCgf
                        Do While ilCgf <> -1
                            tgCgfCntr(ilCgf).CgfRec.iGameNo = -tgCgfCntr(ilCgf).CgfRec.iGameNo
                            ilCgf = tgCgfCntr(ilCgf).iNextCgf
                        Loop
                        
                        slSQLQuery = "Select * From gsf_Game_Schd Where gsfGhfCode = " & ghf_rst!ghfCode
                        Set gsf_rst = gSQLSelectCall(slSQLQuery)
                        Do While Not gsf_rst.EOF
                            ilCgf = tgClfCntr(ilClf).iFirstCgf
                            Do While ilCgf <> -1
                                If Abs(tgCgfCntr(ilCgf).CgfRec.iGameNo) = gsf_rst!gsfGameNo Then
                                    'Test date?
                                    If gDateValue(gsf_rst!gsfAirDate) >= gDateValue(slDate) Then
                                        tgCgfCntr(ilCgf).CgfRec.iGameNo = Abs(tgCgfCntr(ilCgf).CgfRec.iGameNo)
                                        gPackDate gsf_rst!gsfAirDate, tgCgfCntr(ilCgf).CgfRec.iAirDate(0), tgCgfCntr(ilCgf).CgfRec.iAirDate(1)
                                    End If
                                    Exit Do
                                End If
                                ilCgf = tgCgfCntr(ilCgf).iNextCgf
                            Loop
                            gsf_rst.MoveNext
                        Loop
                        'Fix links
                        tgClfCntr(ilClf).ClfRec.lghfcode = ghf_rst!ghfCode
                        ilCgf = tgClfCntr(ilClf).iFirstCgf
                        tgClfCntr(ilClf).iFirstCgf = -1
                        Do While ilCgf <> -1
                            If tgCgfCntr(ilCgf).CgfRec.iGameNo > 0 Then
                                If tgClfCntr(ilClf).iFirstCgf = -1 Then
                                    tgClfCntr(ilClf).iFirstCgf = ilCgf
                                Else
                                    tgCgfCntr(ilPrevCgf).iNextCgf = ilCgf
                                End If
                                ilPrevCgf = ilCgf
                                ilCgf = tgCgfCntr(ilCgf).iNextCgf
                                tgCgfCntr(ilPrevCgf).iNextCgf = -1
                            Else
                                ilCgf = tgCgfCntr(ilCgf).iNextCgf
                            End If
                        Loop
                        'Fix cff
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        tgClfCntr(ilClf).iFirstCff = -1
                        For llSeasonDate = llMoSeasonStart To llSuSeasonEnd Step 7
                            ilNoSpots = 0
                            ilCgf = tgClfCntr(ilClf).iFirstCgf
                            Do While ilCgf <> -1
                                'Test date
                                gUnpackDateLong tgCgfCntr(ilCgf).CgfRec.iAirDate(0), tgCgfCntr(ilCgf).CgfRec.iAirDate(1), llAirDate
                                If (llAirDate >= llSeasonDate) And (llAirDate <= llSeasonDate + 6) Then
                                    ilNoSpots = ilNoSpots + tgCgfCntr(ilCgf).CgfRec.iNoSpots
                                End If
                                ilCgf = tgCgfCntr(ilCgf).iNextCgf
                            Loop
                            If ilNoSpots > 0 Then
                                If tgClfCntr(ilClf).iFirstCff = -1 Then
                                    tgClfCntr(ilClf).iFirstCff = ilCff
                                Else
                                    tgCffCntr(ilPrevCff).iNextCff = ilCff
                                End If
                                tgCffCntr(ilCff).CffRec.iSpotsWk = ilNoSpots
                                gPackDateLong llSeasonDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                gPackDateLong llSeasonDate + 6, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                tgCffCntr(ilCff).lStartDate = llSeasonDate
                                tgCffCntr(ilCff).lEndDate = llSeasonDate + 6
                                slPriceDate = Format$(llSeasonDate, "m/d/yy")
                                llWkPrice = mComputeRCWkPrice(tmRcf.iCode, tgClfCntr(ilClf).ClfRec.iRdfCode, slPriceDate, tgClfCntr(ilClf).ClfRec.iVefCode, 0)
                                Select Case tgClfCntr(ilClf).ClfRec.sPreempt
                                    Case "P"
                                        imDGSave(3) = 1
                                    Case "N"
                                        imDGSave(3) = 0
                                End Select
                                If rbcPrice(1).Value Then
                                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                        llSpots = 0
                                        For ilDay = gWeekDayLong(llDate) To gWeekDayLong(llWkEndDate) Step 1
                                            llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                        Next ilDay
                                    Else
                                        llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                    End If
                                    For ilDay = 0 To 6 Step 1
                                        If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                                            ilDays(ilDay) = True
                                        Else
                                            ilDays(ilDay) = False
                                        End If
                                    Next ilDay
                                    slStr = mRCPriceAdjust(tmRcf, llWkPrice, True, ilClf + 1, imDGSave(), llOvStartTime, llOvEndTime, tgClfCntr(ilClf).ClfRec.iLen, llSpots, ilDays(), tgClfCntr(ilClf).ClfRec.iRdfCode)
                                    llWkPrice = gStrDecToLong(slStr, 0)
                                    llPrice = 100 * llWkPrice
                                    If tgCffCntr(ilCff).CffRec.sPriceType = "T" Then
                                        tgCffCntr(ilCff).CffRec.lActPrice = llPrice  'Later- might want to store average package price
                                        tgCffCntr(ilCff).CffRec.lPropPrice = llWkPrice
                                    End If
                                End If
                                ilPrevCff = ilCff
                                ilCff = tgCffCntr(ilCff).iNextCff
                                tgCffCntr(ilPrevCff).iNextCff = -1
                                If ilCff = -1 Then
                                    'Create new flight record
                                    tgCffCntr(UBound(tgCffCntr)) = tgCffCntr(ilPrevCff)
                                    ilCff = UBound(tgCffCntr)
                                    ReDim Preserve tgCffCntr(LBound(tgCffCntr) To UBound(tgCffCntr) + 1)
                                    'Exit For
                                End If
                            End If
                        Next llSeasonDate
                    Else
                        'Remove events
                        tgClfCntr(ilClf).iFirstCff = -1
                        tgClfCntr(ilClf).iFirstCgf = -1
                    End If
                End If
            End If
            ilClf = ilClf + 1
        Loop
    End If
    
    'Set Dates
    ilClf = LBound(tgClfCntr)
    Do While ilClf < UBound(tgClfCntr)
        ilRet = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
        If ilRet <> -1 Then
            If tgMVef(ilRet).sType <> "G" Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        If tgCffCntr(ilCff).lStartDate <= tgCffCntr(ilCff).lEndDate Then
                            gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llDate
                            slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                            tgCffCntr(ilCff).lStartDate = llDate + ilNoDaysChg
                            gPackDate slDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                            gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llDate
                            slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                            gPackDate slDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                            tgCffCntr(ilCff).lEndDate = llDate + ilNoDaysChg
                        End If
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
            End If
        End If
        ilClf = ilClf + 1
    Loop

    
    'Set Book
    If ckcBook.Value = vbChecked Then
        ilClf = LBound(tgClfCntr)
        Do While ilClf < UBound(tgClfCntr)
            tgClfCntr(ilClf).ClfRec.iDnfCode = 0
            tgClfCntr(ilClf).iPriDemoBook = 0
            tgClfCntr(ilClf).ClfRec.lCPM = 0
            tgClfCntr(ilClf).ClfRec.lCPP = 0
            tgClfCntr(ilClf).ClfRec.lGrImp = 0
            tgClfCntr(ilClf).lPop = 0
            tgClfCntr(ilClf).iAvgRating = 0
            tgClfCntr(ilClf).lAvgAud = 0
            tgClfCntr(ilClf).iPriDemoAvgRating = 0
            tgClfCntr(ilClf).iResearch = False
            tgClfCntr(ilClf).lCPM = 0
            tgClfCntr(ilClf).lCPP = 0
            tgClfCntr(ilClf).lGrImp = 0
            tgClfCntr(ilClf).lGRP = 0
            tgClfCntr(ilClf).lPriDemoAvgAud = 0
            tgClfCntr(ilClf).lPriDemoCPM = 0
            tgClfCntr(ilClf).lPriDemoCPP = 0
            tgClfCntr(ilClf).lPriDemoGrImp = 0
            tgClfCntr(ilClf).lPriDemoGRP = 0
            tgClfCntr(ilClf).lPriDemoPop = 0
            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                tgClfCntr(ilClf).ClfRec.iDnfCode = tgMVef(ilVef).iDnfCode
            End If
            ilClf = ilClf + 1
        Loop
    End If
    If (ilRcfCode > 0) And (rbcPrice(1).Value) Then 'Use rate card price
        ilClf = LBound(tgClfCntr)
        Do While ilClf < UBound(tgClfCntr)
            ilCff = tgClfCntr(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    If tgCffCntr(ilCff).lStartDate <= tgCffCntr(ilCff).lEndDate Then
                        llFlStartDate = tgCffCntr(ilCff).lStartDate
                        llFlEndDate = tgCffCntr(ilCff).lEndDate
                        llMoFlStartdate = llFlStartDate
                        Do While gWeekDayLong(llMoFlStartdate) <> 0
                            llMoFlStartdate = llMoFlStartdate - 1
                        Loop
                        llDate = llMoFlStartdate
                        llWkEndDate = llDate + 6
                        If llWkEndDate > llFlEndDate Then
                            llWkEndDate = llFlEndDate
                        End If
                        ilFirstWk = True
                        Do
                            slDate = Format$(llDate, "m/d/yy")
                            If (tgClfCntr(ilClf).ClfRec.sType = "A") Or (tgClfCntr(ilClf).ClfRec.sType = "O") Then
                                llWkPrice = 0
                            Else
                                llWkPrice = mComputeRCWkPrice(tmRcf.iCode, tgClfCntr(ilClf).ClfRec.iRdfCode, slDate, tgClfCntr(ilClf).ClfRec.iVefCode, 0)
                            End If
                            If llWkPrice > 0 Then
                                'slStr = mRCPriceAdjust(tmRcf, llWkPrice, False, ilClf + 1, imDGSave(), tgClfCntr(ilClf).ClfRec.iLen)
                                Select Case tgClfCntr(ilClf).ClfRec.sPreempt
                                    Case "P"
                                        imDGSave(3) = 1
                                    Case "N"
                                        imDGSave(3) = 0
                                End Select
                                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                    llSpots = 0
                                    For ilDay = gWeekDayLong(llDate) To gWeekDayLong(llWkEndDate) Step 1
                                        llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                    Next ilDay
                                Else
                                    llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                End If
                                For ilDay = 0 To 6 Step 1
                                    If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                                        ilDays(ilDay) = True
                                    Else
                                        ilDays(ilDay) = False
                                    End If
                                Next ilDay
                                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), False, llOvStartTime
                                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), True, llOvStartTime
                                slStr = mRCPriceAdjust(tmRcf, llWkPrice, True, ilClf + 1, imDGSave(), llOvStartTime, llOvEndTime, tgClfCntr(ilClf).ClfRec.iLen, llSpots, ilDays(), tgClfCntr(ilClf).ClfRec.iRdfCode)
                                llWkPrice = gStrDecToLong(slStr, 0)
                            End If
                            If tgCffCntr(ilCff).CffRec.sPriceType = "T" Then
                                llPrice = 100 * llWkPrice
                                If ilFirstWk Then
                                    tgCffCntr(ilCff).CffRec.lPropPrice = llWkPrice
                                    tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                    ilFirstWk = False
                                Else
                                    If (tgCffCntr(ilCff).CffRec.lPropPrice <> llWkPrice) Or (tgCffCntr(ilCff).CffRec.lActPrice <> llPrice) Then
                                        'Split flight
                                        ilUpper = UBound(tgCffCntr)
                                        ReDim Preserve tgCffCntr(0 To ilUpper + 1) As CFFLIST
                                        tgCffCntr(ilUpper + 1).iStatus = -1 'Not Used
                                        tgCffCntr(ilUpper + 1).lRecPos = 0
                                        tgCffCntr(ilUpper + 1).iNextCff = -1
                                        tgCffCntr(ilUpper) = tgCffCntr(ilCff)
                                        tgCffCntr(ilUpper).iStatus = 0
                                        tgCffCntr(ilCff).iNextCff = ilUpper
                                        'Set End Date
                                        'gPackDateLong llWkEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                        'tgCffCntr(ilCff).lEndDate = llWkEndDate
                                        gPackDateLong llPrevWkEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                        tgCffCntr(ilCff).lEndDate = llPrevWkEndDate
                                        ilCff = ilUpper
                                        gPackDateLong llDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                        tgCffCntr(ilCff).lStartDate = llDate
                                        tgCffCntr(ilCff).CffRec.lPropPrice = llWkPrice
                                        tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                    End If
                                End If
                            End If
                            llPrevWkEndDate = llWkEndDate
                            llDate = llDate + 1
                            Do While gWeekDayLong(llDate) <> 0
                                llDate = llDate + 1
                            Loop
                            llWkEndDate = llDate + 6
                            If llWkEndDate > llFlEndDate Then
                                llWkEndDate = llFlEndDate
                            End If
                        Loop While llDate <= llFlEndDate
                    End If
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
            ilClf = ilClf + 1
        Loop
    End If
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.Rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "Deleted") Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = DateAdd("d", ilNoDaysChg, Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE))
                Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = DateAdd("d", ilNoDaysChg, Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE))
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCHFCODE) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT) = ""   'If comment is to be retained, comment out this line but leave the cpmcxfcode line = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) = "N"
            End If
        End If
    Next llRow
    If (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "") Then
        mSetFGColorForRows
    End If
    Screen.MousePointer = vbDefault
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcRateCard_Click()
    lbcRateCard.Visible = Not lbcRateCard.Visible
    edcRateCard.SelStart = 0
    edcRateCard.SelLength = Len(edcRateCard.Text)
    edcRateCard.SetFocus
End Sub
Private Sub cmcRateCard_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDate_Change()
    Dim slStr As String
    slStr = edcDate.Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
End Sub
Private Sub edcRateCard_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcRateCard, lbcRateCard, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
End Sub
Private Sub edcRateCard_GotFocus()
    plcCalendar.Visible = False
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
        imComboBoxIndex = lbcRateCard.ListIndex
    End If
    imBypassFocus = False
End Sub
Private Sub edcRateCard_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcRateCard_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcRateCard.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRateCard_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcRateCard, imLbcArrowSetting
        edcRateCard.SelStart = 0
        edcRateCard.SelLength = Len(edcRateCard.Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Click()
    plcCalendar.Visible = False
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CModel = Nothing   'Remove data segment
End Sub

Private Sub lbcRateCard_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcRateCard, edcRateCard, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub
Private Sub lbcRateCard_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcRateCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilClf As Integer
    Dim llStartDate As Long
    Dim ilCff As Integer
    Dim blSportsVehicle As Boolean
    Dim ilVef As Integer
    Dim blBookShown As Boolean
    Dim blRCShown As Boolean
    Dim llRow As Long
    '10525
    Dim slSql As String
    Dim rst As ADODB.Recordset
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imCalType = 0   'Standard
    '10525
    bgModelNTR = False
    frcNTR.BorderStyle = 0
    frcNTR.Visible = False
    frcNTR.Enabled = False
    frcNTR.Top = 1940
    ckcNTR.Left = ckcBook.Left
    Dim blSportsShown As Boolean
    Dim ilPlaceCounter As Integer
    Dim ilPlaces(3) As Integer
    
    ilPlaces(0) = ckcBook.Top
    ilPlaces(1) = plcPriceInfo.Top
    ilPlaces(2) = frcSports.Top
    ilPlaces(3) = frcNTR.Top
    mInitBox
    'CModel.Height = cmcDone.Top + 7 * cmcDone.Height / 3
    gCenterModalForm CModel
    blSportsVehicle = False
    llStartDate = 0
    ilClf = LBound(tgClfCntr)
    Do While ilClf < UBound(tgClfCntr)
        ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType = "G" Then
                blSportsVehicle = True
            End If
        End If
        ilCff = tgClfCntr(ilClf).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                If tgCffCntr(ilCff).lStartDate <= tgCffCntr(ilCff).lEndDate Then
                    If (tgCffCntr(ilCff).lStartDate < llStartDate) Or (llStartDate = 0) Then
                        llStartDate = tgCffCntr(ilCff).lStartDate   'gDateValue(slStartDate)
                    End If
                End If
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
        ilClf = ilClf + 1
    Loop
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.Rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "Deleted") Then
                If (gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)) < llStartDate) Or (llStartDate = 0) Then
                    llStartDate = gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE))
                End If
            End If
        End If
    Next llRow
    
    If llStartDate <= 0 Then
        llStartDate = gDateValue(gObtainNextMonday(Format$(gNow(), "m/d/yy")))
    End If
    smStartDate = Format$(llStartDate, "m/d/yy")
    edcDate.Text = smStartDate
    blBookShown = True
    blRCShown = True
    '10525
    blSportsShown = True
    If Contract.plcScreen = "Orders" Then
        plcPriceInfo.Visible = False
        blRCShown = False
        If tgSpf.sGUsePropSys <> "Y" Then
            ckcBook.Visible = False
            ckcBook.Value = vbUnchecked
            blBookShown = False
        End If
    Else
        plcPriceInfo.Visible = True
        rbcPrice(1).Value = True
    End If
    mRateCardPop
    If Contract.plcScreen <> "Orders" Then
        If lbcRateCard.ListCount > 0 Then
            lbcRateCard.ListIndex = 0
        End If
    End If
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) And blSportsVehicle Then
        frcSports.Visible = True
        frcSports.Enabled = True
        If blBookShown = False And blRCShown = False Then
            '10525
           ' frcSports.Top = ckcBook.Top
            frcSports.Top = ilPlaces(0)
        '10525 because book and rc never both shown
        ElseIf blRCShown = False Then
       ' ElseIf blRCShown = True = False Then
'            frcSports.Top = plcPriceInfo.Top
            frcSports.Top = ilPlaces(1)
        End If
    Else
        frcSports.Visible = False
        frcSports.Enabled = False
        blSportsShown = False
    End If
    '10525
    slSql = "Select count(sbfCode) as amount from SBF_Special_Billing where sbfchfCode = " & tgChfCntr.lCode
    Set rst = gSQLSelectCall(slSql)
    If Not rst.EOF Then
       ilRet = rst!Amount
    End If
    If ilRet > 0 Then
        If blBookShown Then
            ilPlaceCounter = 1
        End If
        If blRCShown Then
            ilPlaceCounter = ilPlaceCounter + 1
        End If
        If blSportsShown Then
            ilPlaceCounter = ilPlaceCounter + 1
        End If
        frcNTR.Top = ilPlaces(ilPlaceCounter)
'        Select Case ilPlaceCounter
'            Case 0
'                frcNTR.Top = ckcBook.Top
'            Case 1
'                frcNTR.Top = plcPriceInfo.Top
'            Case 2
'                frcNTR.Top = frcSports.Top
'        End Select
    
'        'NTR alone
'        If blBookShown = False And blRCShown = False And blSportsShown = False Then
'            frcNTR.Top = ckcBook.Top
'        'NTR & Book
'        ElseIf blRCShown = False And blSportsShown = False Then
'            frcNTR.Top = plcPriceInfo.Top
'        'NTR & RC
'        ElseIf blSportsShown = False Then
'            frcNTR.Top = frcSports.Top
'        End If
        frcNTR.Visible = True
        frcNTR.Enabled = True
    End If
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    'Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    'flTextHeight = pbcDate.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcModel.Move 105, 375
    plcCalendar.Move plcModel.Left + edcDate.Left, plcModel.Top + edcDate.Top + edcDate.height
    lbcRateCard.Move plcModel.Left + plcPriceInfo.Left + edcRateCard.Left + 30, plcModel.Top + plcPriceInfo.Top + edcRateCard.Top + edcDate.height
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRateCardPop                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Rate Card list        *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mRateCardPop()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim ilRet As Integer
    For ilLoop = 0 To UBound(tmCntRateCard) - 1 Step 1 '-1
        slNameCode = tmCntRateCard(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            lbcRateCard.AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    lbcRateCard.height = gListBoxHeight(lbcRateCard.ListCount, 6)

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Screen.MousePointer = vbDefault
    Unload CModel
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDate.Text = Format$(llDate, "m/d/yy")
                edcDate.SelStart = 0
                edcDate.SelLength = Len(edcDate.Text)
                imBypassFocus = True
                edcDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub plcPriceInfo_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub plcScreen_Click()
    plcCalendar.Visible = False
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcPrice_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Model Information"
End Sub

Private Sub mRemoveSportsVehicleLines()
    Dim ilClf As Integer
    Dim ilRet As Integer
    Dim ilCff As Integer
    Dim blLineRemoved As Boolean
    Dim ilLoop As Integer
    Dim ilLineNo As Integer
    
    ilClf = LBound(tgClfCntr)
    Do While ilClf < UBound(tgClfCntr)
        '2/8/19: Moved from Contract.frm: mReadClfCffRec to here
        blLineRemoved = False
        ilRet = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
        If ilRet <> -1 Then
            If tgMVef(ilRet).sType = "G" Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    tgCffCntr(ilCff).iStatus = -1
                    ilRet = ilCff
                    ilCff = tgCffCntr(ilCff).iNextCff
                    tgCffCntr(ilRet).iNextCff = -1
                Loop
                blLineRemoved = True
                tgClfCntr(ilClf).iFirstCff = -1
                tgClfCntr(ilClf).iFirstCgf = -1
                tgClfCntr(ilClf).lFirstAvail = -1
                ilLineNo = tgClfCntr(ilClf).ClfRec.iLine
                For ilLoop = ilClf To UBound(tgClfCntr) - 1 Step 1
                    tgClfCntr(ilLoop) = tgClfCntr(ilLoop + 1)
                Next ilLoop
                ReDim Preserve tgClfCntr(0 To UBound(tgClfCntr) - 1) As CLFLIST
                For ilLoop = 0 To UBound(tgClfCntr) - 1 Step 1
                    If tgClfCntr(ilLoop).ClfRec.iLine > ilLineNo Then
                        tgClfCntr(ilLoop).ClfRec.iLine = tgClfCntr(ilLoop).ClfRec.iLine - 1
                    End If
                    If tgClfCntr(ilLoop).ClfRec.iPkLineNo > ilLineNo Then
                        tgClfCntr(ilLoop).ClfRec.iPkLineNo = tgClfCntr(ilLoop).ClfRec.iPkLineNo - 1
                    End If
                Next ilLoop
            End If
        End If
        If Not blLineRemoved Then
            ilClf = ilClf + 1
        End If
    Loop

End Sub

Private Sub rbcSports_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
